import csv
import email
import imaplib
import re
import shlex
import time
import traceback

from socket import gaierror

# Globals
"""
<my_dict> = {
    '<mailbox>': {
        'complete': <bool>,
        'data': [{
            'uid': uid,
            'message-id': message-id
        }]
    }
}
"""
COMPLETED = [] # if pipe breaks during copy, can easily resume position upon reconnect

def imap_connect(username, password, server, port=993):
    imap = imaplib.IMAP4_SSL(server)
    imap.login(username, password)

    return imap

def get_namespace(imap):
    """ Get the namespace and separator of the specified mail account

    Example:
        imap.namespace() returns:
            (<code>, [b'(("<namespace>" "<separator>")) NIL NIL'])

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

def prune_protected(mailboxes, ns):
    """ Remove protected folders from a list of mailboxes

    Exchange has several protected folders that are used for non-mail
    purposes, but still appear in the list. Therefore, we want to exclude
    these folders to prevent reading non-mail items.

    Args:
        [<str>]: list of mailboxes - Must not include surrounding quotes
        <str>: namespace for the account

    Returns:
        [<str>]: list of mailboxes with protected folders removed
    """
    protected = ['Calendar', 'Contacts', 'Tasks', 'Journal', 'Deleted Items']
    flagged = []

    for folder in protected:
        pattern = "(?:{})?{}\\b".format(ns,folder)
        for mailbox in mailboxes:
            result = re.match(pattern, mailbox, flags=re.IGNORECASE)
            if result:
                flagged.append(mailbox)
    
    return [mailbox for mailbox in mailboxes if mailbox not in flagged]

def get_mailboxes(imap, with_quotes = True):
    """ Get a list of all mailboxes on the server.

    Args:
        imap <imaplib.IMAP4_SSL>: account to fetch data
        with_quotes <bool>: Whether the resulting array should
            enclose the mailbox names with double quotes

    Returns:
        ["<str>"]: list of mailboxes
            Surrounded by double quotes by default
    """
    mailboxes = [
        shlex.split(mailbox \
            .decode() \
        )[-1].replace('"','') \
        for mailbox in imap.list()[1]
    ]

    ns,_ = get_namespace(imap)
    mailboxes = prune_protected(mailboxes, ns)

    if with_quotes:
        mailboxes = ['"'+mailbox+'"' for mailbox in mailboxes]

    return mailboxes

def convert_mailbox(from_account, to_account, mailbox):
    """ Convert the mailbox to be named appropriately for the destination.
    
    Takes into account differences in namespaces and separators.
    If the mailbox doesn't exist, it will be created.

    Args:
        from_account <imaplib.IMAP4_SSL>: source account
        to_account <imaplib.IMAP4_SSL>: destination account
        mailbox <str>: the mailbox to convert

    Returns:
        "<str>": converted mailbox - surrounded by double quotes
    """
    from_ns, from_sep = get_namespace(from_account)
    to_ns, to_sep = get_namespace(to_account)

    converted = to_ns \
                + mailbox \
                    .replace('"','') \
                    .replace(from_ns, '') \
                    .replace(from_sep, to_sep)    
    converted = '"'+converted+'"'    

    response = to_account.select(converted)[0]
    if response == 'NO':
        to_account.create(converted)
    else:
        to_account.close()

    return converted

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

    from_mailboxes = get_mailboxes(from_account)

    num_mailboxes = len(from_mailboxes)

    for mail_index, from_mailbox in enumerate(from_mailboxes):
        print("{}: Mailbox {} of {}".format(from_mailbox, mail_index+1, num_mailboxes))

        if( from_mailbox not in COMPLETED ):
            data = from_account.select(from_mailbox)[1]
            total_mail = int(data[0])
            if total_mail > 0:
                to_mailbox = convert_mailbox(from_account, to_account, from_mailbox)
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
