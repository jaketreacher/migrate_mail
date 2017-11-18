import csv
import email
import imaplib
import shlex
import time

def imap_connect(username, password, server, port=993):
    imap = imaplib.IMAP4_SSL(server)
    imap.login(username, password)

    return imap

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
    for mail_index, mailbox in enumerate(mailboxes):
        code, data = from_account.select(mailbox)
        total_mail = int(data[0])
        print("{}: {} mail, {}/{}".format(mailbox, total_mail, mail_index+1, num_mailboxes))
        if total_mail > 0:
            # Create mailbox on TO if it doesn't exit
            code = to_account.select(mailbox)[0]
            if code == 'NO':
                to_account.create(mailbox)
                to_account.select(mailbox)

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
                    to_account.append(mailbox, mail['flags'], mail['date'], mail['data'])
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
