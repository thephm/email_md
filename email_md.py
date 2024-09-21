import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import re
from dateutil import parser

import markdownify
from markdownify import markdownify as md

import time
import json
import csv
from datetime import datetime, timezone
import tzlocal 

import sys
sys.path.insert(1, '../../github/message_md/') 
import message_md
import config
import markdown
import message
import person
import attachment

import logging

import warnings
from bs4 import BeautifulSoup, MarkupResemblesLocatorWarning 

INBOX = "INBOX"
SENT = "SENT"

# attribution to https://thepythoncode.com/article/reading-emails-in-python

# use your email provider's IMAP server, you can look for your provider's IMAP server on Google
# or check this page: https://www.systoolsgroup.com/imap/
# use `i <imapServer>`

def get_email_address(text):
    
    import re

    if not isinstance(text, str):
        return text

    # regular expression pattern for extracting email addresses
    email_pattern = r'[\w\.-]+@[\w\.-]+'

    # Use the findall method to extract email addresses from the text
    email_addresses = re.findall(email_pattern, text)

    if email_addresses: 
        result = email_addresses[0]
    else:
        result = False

    return result

# -----------------------------------------------------------------------------
#
# Shortens the body of the email by removing quoted replies (they will be in 
# separate dated files already) and other useless text.
#
# Example:
#
# On Mon, Oct 9, 2023 at 1:06 PM Bob Smith <bob@smith> wrote:
#
# Returns:
# 
# String with the quoted reply text "From: ..."",
# "<!-- ... !-->, and "On ... YYYY ... wrote:" removed 
#
# -----------------------------------------------------------------------------
def remove_reply(text):

    result = text

    patterns = [
        r'^On\s.*?wrote:.*$', # match "On <date/time> <person> wrote:"
        r'Le mer\.|a écrit\s*:' # match "Le mer. <date/time> <person> écrit :"
        r'(?:\*\*From:\*\*|From:).*',  # match "From:" line (with/out asterisks) and everything 
        r'<!--.*?-->',  # matches <!-- ... -->
        r'\\_\\_\\_\\_' # matches \\_\\_\\_\\_
    ]

    for pattern_expression in patterns:
        pattern = re.compile(pattern_expression, re.MULTILINE | re.DOTALL | re.IGNORECASE)  # match across multiple lines, ignore case, and treat ^ and $ as the start/end of each line
        # search for the replied-to part, the "From:" line, or comments
        result = pattern.sub('', result)

        return result.strip()

    return result

# -----------------------------------------------------------------------------
#
# Parse the header of the email into a Message object.
#
# Parameters:
# 
#   - the_email - the actual email
#   - the_message - where the parsed email message goes
#
# Returns: 
#
#   - True if parsed successfully
#   - False if ran into an issue
#
# -----------------------------------------------------------------------------
def parse_header(the_email, the_message):

    result = True   # assume success
    subject = ""

    # decode the email subject
    try:
        subject, encoding = decode_header(the_email["Subject"])[0]
    except Exception as e:
        return False
    
    if encoding and isinstance(subject, bytes):
        try:
            # if it's a bytes, decode to str
            subject = subject.decode(encoding)
        except:
            return False
    
    the_message.subject = subject

    # get the funky IMAP id of the message
    the_message.id = the_email.get('Message-ID')

    # get the date from the header
    date_header = the_email.get('Date')

    # remove extra info after tz offset
    if date_header:
        try:
            date_header = date_header.split(' (', 1)[0]
            parsed_date = parser.parse(date_header)

            # format the date and time
            date_str = parsed_date.strftime('%Y-%m-%d')
            time_str = parsed_date.strftime('%H:%M')

            the_message.timestamp = int(parsed_date.timestamp())
            the_message.date_str = date_str
            the_message.time_str = time_str
        except:
            return False

    # add the person that was passed in `-m slug` to "to_slugs"
    if the_config.me:
        the_message.to_slugs.append(the_config.me.slug)

    # decode email sender
    the_from, encoding = decode_header(the_email.get("From"))[0]
    if encoding and isinstance(the_from, bytes):
        try:
            the_from = the_from.decode(encoding)
        except:
            pass

    if the_from:
        email_addresses = get_email_address(the_from)

    # get the `slug` of the sender
    person = the_config.get_person_by_email(email_addresses)
    if person:
        the_message.from_slug = person.slug
    else: 
        pass

    return result

# -----------------------------------------------------------------------------
#
# Download the attachment from the multi-part email and add a corresponding
# Attachment object to the Messsage object. 
#
# Parameters:
# 
#   - part - the part of the email that has the attachment
#   - theMessage - where the Attachment is added
#
# Returns: nothing
#
# -----------------------------------------------------------------------------
def download_attachment(part, the_message):

    filename = part.get_filename()

    # download the attachment
    if filename:
        # find the place to put it
        folder = os.path.join(the_config.source_folder, the_config.attachments_subfolder)
        file_path = os.path.join(folder, filename)

        try:
            # download attachment and save it
            open(file_path, "wb").write(part.get_payload(decode=True))

            # create and fill the Attachment object
            the_attachment = attachment.Attachment()
            try:
                the_attachment.id = filename
                the_attachment.filename = filename
                the_attachment.type = attachment.get_mime_type(filename)
                the_attachment.custom_filename = filename
                the_message.add_attachment(the_attachment)
            except:
                pass

        except Exception as e:
            logging.error(e)
            pass

# -----------------------------------------------------------------------------
#
# If the email is a multi-part email, parse each part.
#
# Parameters:
# 
#   - the_email - the actual email
#   - the_message - where the parsed email message goes
#
# Returns: nothing
#
# -----------------------------------------------------------------------------
def parse_multi_part(the_email, the_message):

    the_body = ""
    content_disposition = ""
    content_type = ""
    
    # iterate over email parts
    for part in the_email.walk():

        # extract content type of email
        try:
            content_type = part.get_content_type()
        except:
            pass

        try:
            content_disposition = str(part.get("Content-Disposition"))
        except:
            pass

        try:
            the_body = part.get_payload(decode=True).decode()
            if the_body and content_type == "text/plain" and "attachment" not in content_disposition:
                the_message.body = md(the_body)
        except:
            pass

        if "attachment" in content_disposition:
            download_attachment(part, the_message)

# -----------------------------------------------------------------------------
#
# Parse the body of the email. Used when it's not a multi-part email.
#
# Parameters:
# 
#   - the_email - the actual email
#   - the_message - where the parsed email message goes
#
# Returns: nothing
#
# -----------------------------------------------------------------------------
def parse_body(the_email, the_message):

    the_body = ""

    # if the email message is multipart
    if the_email.is_multipart():
        parse_multi_part(the_email, the_message)
    else:
        # extract the content type of the email
        content_type = the_email.get_content_type()

        # get the email body
        try:
            the_body = the_email.get_payload(decode=True).decode()
            if the_body:
                the_message.body = md(the_body)
        
        # sometimes got errors like this, so ignoring the email
        # 'utf-8' codec can't decode byte 0x80 in position 8: invalid start byte
        except Exception as e:
            pass

# -----------------------------------------------------------------------------
#
# Infers whether an email was forwarded from the header fields and subject. 
#
# Parameters:
# 
#   - the_email - the actual email
#   - the_message - where the parsed email message goes
#
# Notes:
#
#   - this is not foolproof but don't need it to be
#
# Returns: 
#
#   - True if inferred it was forwarded
#   - False if it wasn't
#
# -----------------------------------------------------------------------------
def wasForwarded(the_email, the_message):

    result = False

    references = the_email.get('References')
    in_reply_to = the_email.get('In-Reply-To')

    # if References and In-Reply-To are empty, might be a forwarded message
    if not references and not in_reply_to:
        result = True
    
    # if the subject line starts with "Fwd:" or "fwd" or "Fw:" or "fw", then
    # it was likely a forwarded email
    subject = the_message.subject
    pattern = re.compile(r'^fw.*:', re.IGNORECASE)

    # avoid "TypeError: cannot use a string pattern on a bytes-like object"
    if isinstance(subject, str) and not bool(pattern.match(subject)):
        result = True
    
    return result

def isEmailHeader(line):
    # check if the line resembles an email header
    return re.match(r'^\s*(From:|Sent:|To:|Cc:|Subject:)', line, re.IGNORECASE) is not None

def join_lines(body):

    # reassemble lines, preserving paragraph breaks
    lines = body.splitlines()
    paragraphs = []
    current_paragraph = ''

    for line in lines:
        if isEmailHeader(line):
            # if the line resembles an email header, start a new paragraph
            if current_paragraph:
                paragraphs.append(current_paragraph.strip())
                current_paragraph = ''
            current_paragraph += line.strip() + '\n'
        elif line.strip():  # if the line is not empty
            current_paragraph += line + ' '
        else:  # if the line is empty, it indicates a new paragraph
            if current_paragraph and not current_paragraph.isspace():
                paragraphs.append(current_paragraph.strip())
            current_paragraph = ''

    # add the last paragraph if there's any
    if current_paragraph and not current_paragraph.isspace():
        paragraphs.append(current_paragraph.strip())

    # join lines within each paragraph, excluding email headers
    body = '\n\n'.join(para for para in paragraphs)

    return body

# -----------------------------------------------------------------------------
#
# Remove extra stuff from the body of the email like "Sent from my iPhone" and 
# any reply-to text from the email being replied to. 
#
# Parameters:
# 
#   - the_email - the actual email
#   - the_message - where the parsed email message goes
#
# Returns: 
#
#   - True if successful
#   - False if ran into an 
#
# -----------------------------------------------------------------------------
def clean_body(the_email, the_message):

    text = the_message.body

    result = False

    # get rid of quotes, a bit drastic but they're annoying
    text = text.replace('>>  >> ', ' ')
    text = text.replace('>>> ', ' ')
    text = text.replace('>> ', ' ')
    text = text.replace('>>  ', ' ')
    text = text.replace('  >  > ', ' ')
    text = text.replace(' >  > ', ' ') 
    text = text.replace(' > ', ' ') 
    text = text.replace('>  > ', ' ') 
    text = text.replace('> ', ' ') 
    text = text.replace('> > >', ' ')

    # get rid of "{margin:0;}"
    text = text.replace("{margin:0;}", "")
    text = text.replace("{margin: 0;}", "")

    # get rid of this
    text.replace("P {margin-top:0;margin-bottom:0;}", "")

    pattern1 = re.compile(re.escape("#") + '.*?' + re.escape("{margin:0;}"), re.DOTALL)
    text = pattern1.sub('', text)

    pattern2 = re.compile(re.escape("#") + '.*?' + re.escape("NoSpacing"), re.DOTALL)
    text = pattern2.sub('', text)

    # get rid of [External]/[Externe]
    pattern3 = re.compile(r'\[External\]/\[Externe\]', re.IGNORECASE)
    text = pattern3.sub('', text)

    # regular expression pattern to match variations of "Sent from my iPhone" etc.
    # this should likely be an option since it could remove meaningful parts of messages.
    # For me, I want less noise, more signal so the risk of excluding a sentence is low
    pattern4 = re.compile(r'Sent from .*|Get (Outlook for iOS|.*? for Android)|Sent via .*', re.IGNORECASE)
    text = pattern4.sub('', text)

    text = remove_reply(text)

    # add backticks around text with "#" so they aren't seen as tags in 
    # Obsidian e.g. `#bob` 
    text = re.sub(r'#([^\s\)\]\.]+)', r'`#\1`', text)

    # remove "\\_\\_\\_\\_\\_\\_\\_" of any length
    pattern5 = re.compile(r'\\_+', re.MULTILINE)
    text = pattern5.sub('', text)

    # remove "\\\\" of any length
    pattern6 = re.compile(r'^\\\\$', re.MULTILINE)
    text = pattern6.sub('', text)

    # add a blank line before "On Feb 22, 2018, at 8:18 PM, Bob Smith wrote:"
    text = re.sub(r'(On .*? wrote:)', r'\n\1\n', text, flags=re.DOTALL)

    # remove backslash and asterisk around "From," "Sent," and "To"
    text = re.sub(r'\*\*(From|Cc|Sent|To|Subject)\:\*\*', r'\n\1:', text)

    try:
        # do a final removal of any HTML
        text = md(text.strip())
        result = True

    # ignoring exceptions from Beautiful Soup like this one:
    #
    #   "MarkupResemblesLocatorWarning: The input looks more like a URL..."
    #
    except Exception as e:
        pass

    # reassemble lines to avoid word splitting
    text = join_lines(text)

    # get rid of "p.MsoNormal,p.MsoNoSpacing{margin:0}"
    text = text.replace('p.MsoNormal,p.MsoNoSpacing{margin:0}', ' ') 
    pattern7 = re.compile(re.escape("p.") + '(.*?)' + re.escape("{margin: ?0;}"), re.DOTALL)
    text = pattern7.sub('', text)

    # remove leading spaces, likely vestiges from other substitutions above
    text = re.sub(r'^\s*', '', text, flags=re.MULTILINE)

    # get rid of extra newlines
    text = text.replace('\n\n\n', '\n\n')
    
    text = text.replace('| | | --- | |', ' ') 
    text = text.replace('| | --- | |', ' ') 

    # add a blank line before lines starting with From:
    text = re.sub(r'^From: ', '\nFrom: ', text, flags=re.MULTILINE)

    # add a blank line after lines starting with Subject:
    text = re.sub(r'(Subject: .*)', r'\1\n', text)

    # add a line before "---------- Forwarded message ---------"
    text = re.sub(r'\n?-*\s*-?Forwarded message?-*\s*-', '\n\n*- Forwarded message *-', text)

    # add a line before "-----Original Message-----"
    text = re.sub(r'\n?-*\s*?Original Message?-\s*-*\n?', '\n\n*-Original Message*-', text)

    # remove lines between and including "-=-=-=-=-=-=-=-=-=-=-=-"
    text = re.sub(r'-=-=-=-=-=-=-=-=-=-=-=-.*?-=-=-=-=-=-=-=-=-=-=-=-\n?', '', text, flags=re.DOTALL)

    # remove everything after "Join Zoom Meeting"
    text = re.sub(r'Join Zoom Meeting.*', 'Join Zoom Meeting', text, flags=re.DOTALL)

    # FINALLY, ready to put the new-and-improved body in the message 🤣
    the_message.body = text

    return result

# parse a specific email and append it to the list Messages collection
def parse_email(this_email, the_message):

    for response in this_email:
        if isinstance(response, tuple):
            
            try:
                # parse a bytes email into a message object
                this_email = email.message_from_bytes(response[1])

                if the_config.debug:
                    logging.info(f"ID: {id}")

                if parse_header(this_email, the_message):
                    if (the_message.from_slug):
                        parse_body(this_email, the_message)
                        clean_body(this_email, the_message)
            except:
                pass

# -----------------------------------------------------------------------------
#
# Load the emails from a specific IMAP server folder
#
# Parameters:
# 
#   - imap - the imap connection
#   - folder - the folder name e.g. "INBOX"
#   - messsages - where the parsed Message objects are appended 
#
# Returns: the number of emails successfully parsed
#
# Notes:
#
# - was going to search by specific email addresses I know about but decided
#   against that approach
#
#   status, results = imap.search(None, '(HEADER FROM "spongebob@gmail.com")')
#
# -----------------------------------------------------------------------------
def fetch_emails(imap, folder, messages):

    count = 0

    # for some reason, some folders have a space but no quotes around them 
    # and others do have quotes. So, if the folder contains a space and 
    # doesn't have double quotes, add them
    if ' ' in folder and not (folder.startswith('"') and folder.endswith('"')):
        folder = '"' + folder + '"'

    try:
        status, emails = imap.select(folder)
    except Exception as e:
        logging.error(e)
        return count
    
    if status != 'OK':
        logging.error(status)
        return count

    # get the total number of emails in the folder
    emails = int(emails[0])

    for i in range(emails, 0, -1):
        # fetch the email message by ID
        response, this_email = imap.fetch(str(i), "(RFC822)")

         # create a holder for the parsed email
        the_message = message.Message()

        parse_email(this_email, the_message)

        if the_message.from_slug:
            count += 1
            messages.append(the_message)

        # remove the double quotes around the folder name
        if folder.startswith('"') and folder.endswith('"'):
            folder = folder[1:-1]

        # let the user know where processing is at
        status = f"Folder: {folder}  " + f"Countdown: {i-1}  "
        status += f"Found: {count}  Date: {the_message.date_str} "
        status += ' ' * (120 - len(status))
        print(status, end="\r")
        
        # stop if this message was sent before the start date
        if the_message.date_str:
            try:
                message_date = datetime.strptime(the_message.date_str, '%Y-%m-%d')
                from_date = datetime.strptime(the_config.from_date, '%Y-%m-%d')
                if message_date < from_date:
                    continue
            except ValueError:
                print("Error: Date string does not match format '%Y-%m-%d'")

        if count == the_config.max_messages:
            return count

    return count

# -----------------------------------------------------------------------------
#
# Load the emails from the IMAP server.
#
# Parameters:
# 
#   - dest_file - not used but needed for the interface
#   - messages - where the Message objects will go
#   - reactions - not used but needed for the interface
#   - the_config - specific settings 
#
# Returns: the number of messages loaded
#
# -----------------------------------------------------------------------------
def load_messages(dest_file, messages, reactions, the_config):

    count = 0
    folders = []

    if not (the_config.imap_server and the_config.email_account and the_config.password):
        return 0

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL(the_config.imap_server)

    # authenticate, get the list of folders, fetch the emails
    try:
        imap.login(the_config.email_account, the_config.password)

        if len(the_config.email_folders) > 0:
            folders = the_config.email_folders
        else:
            # log the list of folders
            logging.info(imap.list()[1])

            for i in imap.list()[1]:
                l = i.decode().split(' "/" ')
                folders.append(l[1])

        # remove any folders specified in `not-email-folders` 
        # setting and any of it's subfolders
        for x_folder in the_config.not_email_folders:
            for y_folder in folders:
                parts = y_folder.split('/')
                if parts[0] == x_folder:
                    folders.remove(y_folder)

        for folder in folders:
            # only fetch emails from folders not in the exclude list
            count += fetch_emails(imap, folder, messages)
            if count >= the_config.max_messages:
                imap.close()
                imap.logout()
                return count

    except Exception as e:
        logging.error(e)

    # close the connection and logout
    try:
        imap.close()
        imap.logout()
    except:
        pass

    return count

# main

# Was getting the warning below and found no easy way to supress it so using: 
# https://stackoverflow.com/questions/36039724/suppress-warning-of-url-in-beautifulsoup
#
# /home/bjansen/.local/lib/python3.10/site-packages/markdownify/__init__.py:96:
# MarkupResemblesLocatorWarning: The input looks more like a URL than markup. 
# You may want to use an HTTP client like requests to get the document behind 
# the URL, and feed that document to Beautiful Soup.
#  soup = BeautifulSoup(html, 'html.parser')
#
warnings.filterwarnings('ignore', category=MarkupResemblesLocatorWarning)

the_messages = []
the_reactions = [] 

the_config = config.Config()

if message_md.setup(the_config, markdown.YAML_SERVICE_EMAIL):

    # needs to be after setup so the command line parameters override the
    # values defined in the settings file
    message_md.get_markdown(the_config, load_messages, the_messages, the_reactions)

    print("\n")