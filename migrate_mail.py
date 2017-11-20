import csv
import email
import imaplib
import re
import shlex
import time
import traceback

from socket import gaierror

# Globals
COMPLETED = [] # if pipe breaks during copy, can easily resume position upon reconnect

def imap_connect(username, password, server, port=993):
    imap = imaplib.IMAP4_SSL(server)
    imap.login(username, password)

    return imap

def get_namespace(imap):
    """ Get the namespace and separator of the specified mail account

    Explanation:
        imap.namespace() returns:
            (<code>, [b'(("<namespace>" "<separator>")) NIL NIL'])
        
        imap.namespace()[1][0] returns:
            '(("<namespace>" "<separator>")) NIL NIL'
        
        Using replace and shlex returns:
            [<namespace>, <separator>, NIL, NIL]

    Args:
        imap <imaplib.IMAP4_SSL>: the account to check

    Returns:
        (namespace <str>, separator <str>)
    """
    results = shlex.split(imap.namespace()[1][0].decode() \
        .replace('(','') \
        .replace(')','')
    )

    return (results[0], results[1])

def get_mailboxes(from_server, to_server):
    """ Get the mailboxes of the from_server and to_server.
    
    Determines the appropriately named mailbox on the to_server,
    taking into account differences in namespaces and separators.

    The resulting mailbox names are enclosed in double quotes.

    Args:
        from_server <imaplib.IMAP4_SSL>: to source account
        from_server <imaplib.IMAP4_SSL>: to destination account

    Returns:
        [(from_mailbox <str>, to_mailbox <str>)]

        Example:
            [("INBOX.My Folder.Subfolder", "My Folder/Subfolder")]
    """
    from_ns, from_sep = get_namespace(from_server)
    to_ns, to_sep = get_namespace(to_server)
    
    from_mailboxes = [
        shlex.split(mailbox \
            .decode() \
        )[-1].replace('"','') \
        for mailbox in from_server.list()[1]
    ]

    to_mailboxes = [
        to_ns \
        + mailbox \
            .replace(from_ns, '') \
            .replace(from_sep, to_sep) \
        for mailbox in from_mailboxes
    ]

    all_mailboxes = [('"'+from_mailbox+'"', '"'+to_mailbox+'"') for from_mailbox, to_mailbox in zip(from_mailboxes, to_mailboxes)]
    return all_mailboxes

def get_headers(imap, location):
    header_list = []

    data = imap.uid('search', None, 'ALL')[1]
    uid_list = data[0].split()
    length = len(uid_list)

    for idx, uid in enumerate(uid_list):
        # Progress indicator
        print("Fetching {} headers... {}/{}".format(location, idx+1, length), end='\r')

        data = imap.uid('fetch', uid, '(BODY.PEEK[HEADER])')[1]
        message_id = email.message_from_bytes(data[0][1])['Message-ID'] # Used to check for duplicates

        header_dict = {
            'uid': uid,
            'Message-ID': message_id
        }

        header_list.append(header_dict)

    if( length > 0 ):
        print() # Move cursor to next line
    else:
        print("Fetching %s headers... 0/0" % location)

    return header_list

def get_mail_by_uid(imap, uid):
    data = imap.uid('fetch', uid, '(FLAGS INTERNALDATE RFC822)')[1]
    flags = " ".join([flag.decode() for flag in imaplib.ParseFlags(data[0][0])])
    date_time = imaplib.Internaldate2tuple(data[0][0])

    mail_data = {
        'flags': flags,
        'date_time': date_time,
        'message':  data[0][1],
    }

    return mail_data

def copy_mail(from_account, to_account):
    global COMPLETED
    mailboxes = get_mailboxes(from_account, to_account)

    num_mailboxes = len(mailboxes)

    for mail_index, (from_mailbox, to_mailbox) in enumerate(mailboxes):
        print("{}: Mailbox {} of {}".format(from_mailbox, mail_index+1, num_mailboxes))

        if( from_mailbox not in COMPLETED ):
            data = from_account.select(from_mailbox)[1]
            total_mail = int(data[0])
            if total_mail > 0:
                # Create mailbox on destination if it doesn't exist
                code = to_account.select(to_mailbox)[0]
                if code == 'NO':
                    to_account.create(to_mailbox)
                    to_account.select(to_mailbox)

                # Get all headers
                from_headers = get_headers(from_account, "source")
                to_headers = get_headers(to_account, "destination")

                # Remove Duplicates
                unique_headers = [header for header in from_headers \
                    if header['Message-ID'] not in \
                        [header['Message-ID'] for header in to_headers] \
                ]

                length = len(unique_headers)
                if length > 0:
                    for idx, header in enumerate(unique_headers):
                        print("Copying mail... {}/{}".format(idx+1, length), end='\r')
                        to_account.append(to_mailbox, **get_mail_by_uid(from_account, header['uid']))
                    print()
                else:
                    print('No new mail')

                COMPLETED.append(from_mailbox)
                to_account.close()
            from_account.close()
        else:
            print("Completed, skipping...")
        print() # new line for formatting

def fancy_sleep(message, duration):
    for idx in range(duration, -1, -1):
        print("%s %s " % (message, idx), end="\r")
        time.sleep(1)
    print()

def main():
    global COMPLETED
    dict_list = []
    error_file = open('log/%d.txt' % int(time.time()), 'w')

    with open('data.csv') as datafile:
        reader = csv.DictReader(datafile)
        dict_list = list(reader)

    for data in dict_list:
        COMPLETED = []
        success = False
        try:
            while( not success ):
                print("Connecting to %s" % data['FROM_MAIL'])
                from_account = imap_connect(data['FROM_MAIL'], data['FROM_PASS'], data['FROM_SERVER'])

                print("Connecting to %s" % data['TO_MAIL'])
                to_account = imap_connect(data['TO_MAIL'], data['TO_PASS'], data['TO_SERVER'])

                print("--- From: {}, To: {} ---".format(data['FROM_MAIL'], data['TO_MAIL']))
                try:
                    copy_mail(from_account, to_account)
                    success = True
                except imaplib.IMAP4.abort:
                    print("Error: Connection interrupted.")
                    fancy_sleep("Reconnecting in:", 10)
                    from_account.logout()
                    to_account.logout()

        except (ConnectionRefusedError, gaierror, imaplib.IMAP4.error) as e:
            print("  Unable to connect.")
            error_file.write("%s => %s\n" % (data['FROM_MAIL'], data['TO_MAIL']))
            error_file.write(traceback.format_exc(3) + "\n======\n")
            continue
        
        from_account.logout()
        to_account.logout()

    error_file.close()

if __name__ == "__main__":
    main()
