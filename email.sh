# A helper bash script for https://github.com/thephm/email_md
# Run this in WSL ubunu shell
# To install WSL: https://learn.microsoft.com/en-us/windows/wsl/install

#!/bin/bash

# Check if the email address and password was provided
if [ $# -lt 2 ]; then
  echo "Usage: $0 <email_address> <password>"
  exit 1
fi

# Get the email address from the command line
EMAIL="$1"
PASSWORD="$2"

# IMAP server
IMAP=imap.fastmail.com

# this is your slug
SLUG=bob

# location of the Python script
PY_DIR=/mnt/c/data/github/email_md

# configuration for signal_sqlite_md
CONFIG_DIR=/mnt/c/data/dev-output/config

# location to put the output Markdown files from signal_sqlite_md
OUTPUT_DIR=/mnt/c/data/dev-output

cd $PY_DIR
python3 email_md.py -c $CONFIG_DIR -s $CONFIG_DIR -d -o $OUTPUT_DIR -m $SLUG -i $IMAP -e $EMAIL -p $PASSWORD -b 1900-01-01
