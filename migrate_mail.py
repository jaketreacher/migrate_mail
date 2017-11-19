import csv
import email
import imaplib
import shlex
import time

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
        mail_table[mailbox['name']] = \
            change_namespace(
                mailbox['name'],
                from_namespace,
                to_namespace,
                mailbox['sep']
            ).replace(mailbox['sep'], new_sep)
    
    return mail_table

def get_all_mail(imap, location):
    typ, data = imap.uid('search',None, 'ALL')
    mail_list = []
    uid_list = data[0].split()
    length = len(uid_list)

    for idx, uid in enumerate(uid_list):
        print("Fetching {}... {}/{}".format(location, idx+1, length), end='\r') # Progress indicator

        data = imap.uid('fetch', uid, '(FLAGS INTERNALDATE RFC822)')[1]
        flags = " ".join([flag.decode() for flag in imaplib.ParseFlags(data[0][0])])
        date = imaplib.Internaldate2tuple(data[0][0])
        message_id = email.message_from_bytes(data[0][1])['Message-ID'] # Used to check for duplicates

        mail_dict = {
            'uid': uid,
            'data': data[0][1],
            'flags': flags,
            'date': date,
            'Message-ID': message_id
        }

        mail_list.append(mail_dict)

    if( length > 0 ):
        print() # Move cursor to next line
    else:
        print("Fetching %s... 0/0" % location)

    return mail_list

def copy_mail(from_account, to_account):
    mailboxes = ['"' + shlex.split(item.decode())[-1] + '"' for item in from_account.list()[1]]

    num_mailboxes = len(mailboxes)
    # mail table used as a hash to convert old mailbox to new format, if required
    mail_table = get_mail_table(from_account, to_account)

    for mail_index, mailbox in enumerate(mailboxes):
        code, data = from_account.select(mailbox)
        total_mail = int(data[0])
        print("{}: {} mail, {}/{}".format(mailbox, total_mail, mail_index+1, num_mailboxes))
        if total_mail > 0:
            # Create mailbox on destination if it doesn't exist
            code = to_account.select(mail_table[mailbox])[0]
            if code == 'NO':
                to_account.create(mail_table[mailbox])
                to_account.select(mail_table[mailbox])

            # Get all mail
            from_mail = get_all_mail(from_account, "source")
            to_mail = get_all_mail(to_account, "destination")

            # Remove Duplicates
            unique_mail = [mail for mail in from_mail \
                if mail['Message-ID'] not in \
                    [mail['Message-ID'] for mail in to_mail] \
            ]

            length = len(unique_mail)
            if length > 0:
                for idx, mail in enumerate(unique_mail):
                    print("Copying mail... {}/{}".format(idx+1, length), end='\r')
                    to_account.append(mail_table[mailbox], mail['flags'], mail['date'], mail['data'])
                print()
            else:
                print('No new mail')

            to_account.close()
            print() # new line for formatting
        from_account.close()

def main():
    dict_list = []
    error_file = open('errors.txt', 'w')

    with open('data.csv') as datafile:
        reader = csv.DictReader(datafile)
        dict_list = list(reader)

    for data in dict_list:
        try:
            print("Connecting to %s" % data['FROM_MAIL'])
            from_account = imap_connect(data['FROM_MAIL'], data['FROM_PASS'], data['FROM_SERVER'])

            print("Connecting to %s" % data['TO_MAIL'])
            to_account = imap_connect(data['TO_MAIL'], data['TO_PASS'], data['TO_SERVER'])
        except Exception as e:
            print("  Unable to connect.")
            error_file.write("%s => %s || (%s)\n" % (data['FROM_MAIL'], data['TO_MAIL'], e))
            continue

        print("--- From: {}, To: {} ---".format(data['FROM_MAIL'], data['TO_MAIL']))
        copy_mail(from_account, to_account)

        from_account.logout()
        to_account.logout()

    error_file.close()

if __name__ == "__main__":
    main()
