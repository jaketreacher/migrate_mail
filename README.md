# Migrate Mail v 0.4.0

## Synopsis

A command line script to migrate multiple mail accounts between servers.

## Usage
Setup `data.csv` with the appropriate details.  
Run `python3 migrate_mail.py`

## Notes
1. The following mailboxes are _protected_, and will be ignored:
    - `Calendar`, `Contacts`, `Tasks`, `Journal`, `Deleted Items`  
    If you need to copy mail from these directories, consider renaming them.
2. Refrain from using either the source or destination when mail is being copies. It may result in unexpected errors.

## To-Do
* Better logging

## License

Copyright (c) Jake Treacher. All rights reserved.  
Licensed under the [MIT](https://github.com/jaketreacher/migrate_mail/blob/master/LICENSE.txt) License.  
