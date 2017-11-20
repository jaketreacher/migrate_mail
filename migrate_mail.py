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

def get_mailbox_info(mailbox_data):
    data = re.search('(?P<flags>\(.*\)) "(?P<sep>.)" (?P<name>.*)', 
            mailbox_data[0].decode()).groupdict()
    sep = data['sep']
    namespace = data['name'] if (sep == '.') else ''
    return sep, namespace

def convert_mailbox_format(mailbox_data):
    data_list = []
    for box in mailbox_data:
        data_list.append(
            re.search('(?P<flags>\(.*\)) "(?P<sep>.)" (?P<name>.*)', 
            box.decode()).groupdict()
        )
    return data_list

def change_namespace(name, old_ns, new_ns, old_sep):
    name = name.replace('"','')
    if (name.upper() != "INBOX"):
        if (old_ns != '') and (new_ns == ''):
            name = name.replace(old_ns + old_sep, '', 1)
        if (old_ns == '') and (new_ns != ''):
            name = new_ns + old_sep + name
        if (old_ns != '') and (new_ns != ''):
            name = name.replace(old_ns + old_sep, new_new + old_sep, 1)
    return '"' + name + '"'

def get_mail_table(from_server, to_server):
    from_mailbox_data = from_server.list()[1]
    to_mailbox_data = to_server.list()[1]

    _, from_namespace = get_mailbox_info(from_mailbox_data)
    new_sep, to_namespace = get_mailbox_info(to_mailbox_data)

    mailbox_list = convert_mailbox_format(from_mailbox_data)
    mail_table = dict()

    for mailbox in mailbox_list:
        # Change the namespace and replace the separators
        key = '"' + mailbox['name'].replace('"','') + '"'
        data = change_namespace(
                    mailbox['name'],
                    from_namespace,
                    to_namespace,
                    mailbox['sep']
                ).replace(mailbox['sep'], new_sep)
        mail_table[key] = data
    
    return mail_table

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
    mailboxes = ['"' + shlex.split(item.decode())[-1] + '"' for item in from_account.list()[1]]

    num_mailboxes = len(mailboxes)
    # mail table used as a hash to convert old mailbox to new format, if required
    mail_table = get_mail_table(from_account, to_account)

    for mail_index, mailbox in enumerate(mailboxes):
        print("{}: Mailbox {} of {}".format(mailbox, mail_index+1, num_mailboxes))

        if( mailbox not in COMPLETED ):
            data = from_account.select(mailbox)[1]
            total_mail = int(data[0])
            if total_mail > 0:
                # Create mailbox on destination if it doesn't exist
                code = to_account.select(mail_table[mailbox])[0]
                if code == 'NO':
                    to_account.create(mail_table[mailbox])
                    to_account.select(mail_table[mailbox])

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
                        to_account.append(mail_table[mailbox], **get_mail_by_uid(from_account, header['uid']))
                    print()
                else:
                    print('No new mail')

                COMPLETED.append(mailbox)
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
