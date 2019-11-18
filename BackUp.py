from shutil import copyfile
import os.path
import datetime
from datetime import timedelta

def DoBackup():
    unique_end = datetime.datetime.strftime(datetime.datetime.now(),"%d-%b-%Y %H%M%S")
    src = r'' # source file string
    dest = rf'{unique_end}.xlsm' # destination file string

    print(f'Copying: {src}')
    try:
        copyfile(src,dest)
        print(f'Copy completed, file stored in {dest}')
    except Exception as e:
        print('Copy failed')
        print(str(e))

def main():
    backup_choice = ''
    while backup_choice != 'y' or backup_choice != 'n':
        backup_choice = input('Do you want to back up the Collection Log now? (y/n): ')
        if backup_choice.lower() == 'y':
            DoBackup()
            break
        elif backup_choice.lower() == 'n':
            break
        else:
            print('invalid input\n')


if __name__ == "__main__":
    main()