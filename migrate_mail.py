import csv
import email
import imaplib
import re
import shlex
import time
import traceback

from socket import gaierror

def imap_connect(username, password, server, port=993):
    """ Connect to the server using IMAP SSL

    Args:
        username <str>
        password <str>
        server <str>
        port <int>: default 993 for SSL
    
    Returns:
        <imaplib.IMAP4_SSL>: Reference to the connection
    """
    imap = imaplib.IMAP4_SSL(server)
    imap.login(username, password)

    return imap

def authenticate(verbose = True):
    """ Authenticates accounts for both the server and the destination.

    Args:
        verbose <bool>: if true, will display messages.

    Globals:
        CREDENTIALS: a dictionary of credentials.
            Set to the most recent accounts.
            Configured in main().
    """
    global CREDENTIALS

    if verbose: print("Connecting to %s" % CREDENTIALS['FROM_MAIL'])
    from_account = imap_connect(
        CREDENTIALS['FROM_MAIL'],
        CREDENTIALS['FROM_PASS'],
        CREDENTIALS['FROM_SERVER']
    )

    if verbose: print("Connecting to %s" % CREDENTIALS['TO_MAIL'])
    to_account = imap_connect(
        CREDENTIALS['TO_MAIL'],
        CREDENTIALS['TO_PASS'],
        CREDENTIALS['TO_SERVER']
    )
    
    print() # New line for formatting

    return from_account, to_account

def fancy_sleep(message, duration):
    """ Timeout with a countdown that displays on a single line.

    Args:
        message <str>: The message to display
        duration <int>: The countdown duration
    """
    for idx in range(duration, -1, -1):
        print("%s %s " % (message, idx), end="\r")
        time.sleep(1)
    print()

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

def get_mail_count(imap, mailbox_list):
    """ Gets the total number of emails on specified account.

    Args:
        imap <imaplib.IMAP4_SSL>: the account to check
        mailbox_list [<str>]: a list of mailboxes
            Must be surrounded by double quotes
    
    Returns:
        <int>: total emails
    """
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

def get_message_ids(imap):
    """ Get all Message-IDs for the specified account

    Args:
        imap <imaplib.IMAP4_SSL>: account to fetch data

    Returns:
        A dictionary in the following format:
        package = {
            'mailbox0': [(uid0, mid0), (uid1, mid1)],
            'mailbox1': [(uid2, mid2), (uid3, mid3)]
        }
        mid = message-id
    """
    package = dict()

    mailboxes = get_mailboxes(imap)
    total = get_mail_count(imap, mailboxes)
    remaining = total
    for mailbox in mailboxes:
        package[mailbox] = list()

        imap.select(mailbox)
        uid_list = imap.uid('search', None, 'ALL')[1][0].split()

        for uid in uid_list:
            print('Remaining: %d ' % remaining, end='\r')
            data = imap.uid('fetch', uid, '(BODY.PEEK[HEADER])')[1]
            try:
                message_id = email.message_from_bytes(data[0][1])['Message-ID']
                if message_id == None:
                    remaining -= 1
                    continue
            except TypeError:
                # unable to get message-id
                # to-do: implement logging here
                remaining -= 1
                continue

            data = (uid, message_id)
            package[mailbox].append(data)
            remaining -= 1
    print('Remaining: %d ' % remaining)
    
    return package

def get_mail_by_uid(imap, uid):
    """ The full email as determined by the UID.
    
    Note: Must have used imap.select(<mailbox>) before
    running this function.

    Args:
        imap <imaplib.IMAP4_SSL>: the server to fetch

    Returns:
        Dictionary, with the keys 'flags', 'date_time', and
        'message'. Contents are self-explanatory...
    """
    data = imap.uid('fetch', uid, '(FLAGS INTERNALDATE RFC822)')[1]
    flags = " ".join([flag.decode() for flag in imaplib.ParseFlags(data[0][0])])
    date_time = imaplib.Internaldate2tuple(data[0][0])

    mail_data = {
        'flags': flags,
        'date_time': date_time,
        'message':  data[0][1],
    }

    return mail_data

def get_unique_uids(from_account, to_account):
    """ Get all UIDs and their corresponding mailbox for emails
    that do not appear on the destination.

    Args:
        from_account <imaplib.IMAP4_SSL>: source
        to_account <imaplib.IMAP4_SSL>: destination

    Returns:
        {}, <int>
        Dictionary in the format:
            '<mailbox>': [uids]
        <int>: total number of uids
    """
    print("Fetching source headers...")
    from_mid_package = get_message_ids(from_account)
    print()
    print("Fetching destination headers...")
    to_mid_package = get_message_ids(to_account)
    print()

    to_mids = list()
    for mailbox, data in to_mid_package.items():
        for uid, mid in data:
            to_mids.append(mid)

    unique = dict()
    total = 0
    for mailbox, data in from_mid_package.items():
        unique[mailbox] = list()
        for uid, mid in data:
            if mid not in to_mids:
                unique[mailbox].append(uid)
                total += 1

    return unique, total

def copy_mail(from_account, to_account):
    """ Copies all emails from the source to the destination.

    Will not copy any mail that already appears in the destination.

    Args:
        from_account <imaplib.IMAP4_SSL>: source
        to_account <imaplib.IMAP4_SSL>: destination
    """
    unique_uid_package, total = get_unique_uids(from_account, to_account)

    remaining = total

    print("Copying mail...")
    print("Total: %d" % total)
    if total > 0:
        for mailbox, uid_list in unique_uid_package.items():
            to_mailbox = convert_mailbox(from_account, to_account, mailbox)
            from_account.select(mailbox)
            
            for uid in uid_list:
                print('Remaining: %d ' % remaining, end='\r')
                success = False
                while(not success):
                    try:
                        # The pipe was breaking here when transferring to Exchange.
                        # Hence, we catch the exception and reconnect, which
                        # appears to fix the problem.
                        to_account.append(to_mailbox, **get_mail_by_uid(from_account, uid))
                        remaining -= 1
                        success = True
                    except (imaplib.IMAP4.error, imaplib.IMAP4.abort):
                        print("\nError: Attempting to reconnect")
                        fancy_sleep("Reconnecting in:", 10)
                        from_account.logout()
                        to_account.logout()
                        from_account, to_account = authenticate(False)
                        from_account.select(mailbox)
                        to_account.select(to_mailbox)
        print('Remaining: %d ' % remaining)
    else:
        print('No mail to copy')
    
def main():
    # CREDENTIALS used as a global to allow to reconnect if the pipe
    # breaks when copying emails.
    global CREDENTIALS
    
    error_file = open('log/errors.txt', 'a')

    with open('data.csv') as datafile:
        reader = csv.DictReader(datafile)
        data_list = list(reader)

    for data in data_list:
        CREDENTIALS = data
        try:
            print("--- From: {}, To: {} ---".format(data['FROM_MAIL'], data['TO_MAIL']))
            from_account, to_account = authenticate()
            copy_mail(from_account, to_account)
        except (ConnectionRefusedError, gaierror, imaplib.IMAP4.error) as e:
            print("  Unable to connect.")
            error_file.write("%s => %s\n" % (data['FROM_MAIL'], data['TO_MAIL']))
            error_file.write(traceback.format_exc(3) + "\n======\n")
            continue
        
        from_account.logout()
        to_account.logout()
        print()
    error_file.close()

if __name__ == "__main__":
    main()
