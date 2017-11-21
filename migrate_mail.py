import csv
import email
import imaplib
import re
import shlex
import time
import traceback

from socket import gaierror

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

def get_mail_count(imap, mailbox_list):
    total = 0
    num_mailboxes = len(mailbox_list)
    for idx, mailbox in enumerate(mailbox_list):
        print("Counting mail: %d (Mailbox %d of %d) " \
            % (total, idx+1, num_mailboxes), end='\r')
        total += int(imap.select(mailbox)[1][0])
        imap.close()
    print("Counting mail: %d (Mailbox %d of %d) " \
        % (total, idx+1, num_mailboxes))
    return total

def get_all_headers(imap):
    """ Get all headers

    Args:
        imap <imaplib.IMAP4_SSL>: account to fetch data

    Returns:
        A dictionary in the following format:
        headers = {
            'mailbox0': [(uid0, mid0), (uid1, mid1)],
            'mailbox1': [(uid2, mid2), (uid3, mid3)]
        }
        mid = message-id
    """
    headers = dict()

    mailboxes = get_mailboxes(imap)
    total = get_mail_count(imap, mailboxes)
    remaining = total
    for mailbox in mailboxes:
        headers[mailbox] = list()

        imap.select(mailbox)
        uid_list = imap.uid('search', None, 'ALL')[1][0].split()

        for uid in uid_list:
            print('Remaining: %d ' % remaining, end='\r')
            data = imap.uid('fetch', uid, '(BODY.PEEK[HEADER])')[1]
            try:
                message_id = email.message_from_bytes(data[0][1])['Message-ID']
            except TypeError:
                # unable to get message-id
                # to-do: implement logging here
                remaining -= 1
                continue

            package = (uid, message_id)
            headers[mailbox].append(package)
            remaining -= 1
    print('Remaining: %d ' % remaining)
    
    return headers

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

def get_unique_headers(from_account, to_account):
    print("Fetching source headers...")
    from_headers = get_all_headers(from_account)
    print()
    print("Fetching destination headers...")
    to_headers = get_all_headers(to_account)
    print()

    to_mids = list()
    for mailbox, data in to_headers.items():
        for uid, mid in data:
            to_mids.append(mid)

    unique = dict()
    total = 0
    for mailbox, data in from_headers.items():
        unique[mailbox] = list()
        for uid, mid in data:
            if mid not in to_mids:
                unique[mailbox].append(uid)
                total += 1

    return unique, total

def copy_mail(from_account, to_account):
    unique_headers, total = get_unique_headers(from_account, to_account)

    remaining = total

    if total > 0:
        print("Copying mail...")
        print("Total: %d" % total)
        for mailbox, uid_list in unique_headers.items():
            to_mailbox = convert_mailbox(from_account, to_account, mailbox)
            from_account.select(mailbox)
            
            for uid in uid_list:
                print('Remaining: %d ' % remaining, end='\r')
                to_account.append(to_mailbox, **get_mail_by_uid(from_account, uid))
                remaining -= 1
        print('Remaining: %d ' % remaining)
    else:
        print('No mail to copy')
    
def fancy_sleep(message, duration):
    for idx in range(duration, -1, -1):
        print("%s %s " % (message, idx), end="\r")
        time.sleep(1)
    print()

def main():
    error_file = open('log/%d.txt' % int(time.time()), 'w')

    with open('data.csv') as datafile:
        reader = csv.DictReader(datafile)
        data_list = list(reader)

    for data in data_list:
        success = False
        try:
            while( not success ):
                print("--- From: {}, To: {} ---".format(data['FROM_MAIL'], data['TO_MAIL']))
                
                print("Connecting to %s" % data['FROM_MAIL'])
                from_account = imap_connect(data['FROM_MAIL'], data['FROM_PASS'], data['FROM_SERVER'])

                print("Connecting to %s" % data['TO_MAIL'])
                to_account = imap_connect(data['TO_MAIL'], data['TO_PASS'], data['TO_SERVER'])

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
