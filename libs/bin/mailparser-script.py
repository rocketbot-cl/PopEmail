#!d:\python36\python.exe
# EASY-INSTALL-ENTRY-SCRIPT: 'mail-parser==3.9.3','console_scripts','mailparser'
__requires__ = 'mail-parser==3.9.3'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('mail-parser==3.9.3', 'console_scripts', 'mailparser')()
    )
