import csv
import imaplib
import shlex
import time

def imap_connect(username, password, server, port=993):
    imap = imaplib.IMAP4_SSL(server)
    imap.login(username, password)

    return imap

def get_all_mail(imap):
    typ, data = imap.search(None, 'ALL')
    mail_list = []
    for num in data[0].split():
        data = imap.fetch(num, '(FLAGS INTERNALDATE RFC822)')[1]
        flags = " ".join([flag.decode() for flag in imaplib.ParseFlags(data[0][0])])
        date = imaplib.Internaldate2tuple(data[0][0])
        
        mail_dict = {
            'data': data[0][1],
            'flags': flags,
            'date': date
        }

        mail_list.append(mail_dict)
    
    return mail_list



def copy_mail(from_account, to_account):
    mailboxes = ['"' + shlex.split(item.decode())[-1] + '"' for item in from_account.list()[1]]

    for mailbox in mailboxes:
        code, data = from_account.select(mailbox)
        count = int(data[0])
        print("{}: {}".format(mailbox, count))
        if count > 0:
            # Create mailbox on TO if it doesn't exit
            code = to_account.select(mailbox)[0]
            if code == 'NO':
                to_account.create(mailbox)
                to_account.select(mailbox)

            # Get all mail
            print("Fetching mail... ")
            from_mail_list = get_all_mail(from_account)
            to_mail_list = get_all_mail(to_account)

            # Remove Duplicates - not working
            # from_mail_list = [ mail for mail in from_mail_list if mail not in to_mail_list]

            length = len(from_mail_list)
            if length > 0:
                print("Copying mail... ")
                for idx, mail in enumerate(from_mail_list):
                    print("{}/{}".format(idx+1, length), end='\r')
                    to_account.append(mailbox, mail['flags'], mail['date'], mail['data'])
                print()
            else:
                print('No new mail')

            to_account.close()
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
