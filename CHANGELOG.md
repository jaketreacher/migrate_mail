# Change Log

## 0.4.0 (2017-11-21)
- Resume from broken pipe improved
    - Resume from where you left off, rather than from the start of the current mailbox
- Check if duplicate exists _anywhere_ on destination
    - Was previously only checking mail in the same directory
- Ignore protected directories
    - Exchange lists these directories, and accessing items in them results in errors
    - `Calendar`, `Contacts`, `Tasks`, `Journal`, `Deleted Items`

## 0.3.1 (2017-11-19)
- Fix error when resuming from broken pipe
- Fix print typos

## 0.3.0 (2017-11-19)
- Automatically convert namespace/separators between servers
- Attempt to reconnect if pipe breaks
- Fetch headers when checking for duplicates

## 0.2.0 (2017-11-18)
- If unable to connect to account, will skip that account and log the error rather than exiting the program.

## 0.1.0 (2017-11-18)
- Read files from .csv
- Copy all mail from the source to the destination