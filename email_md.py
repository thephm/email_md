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

def getEmailAddress(text):
    
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
def removeReply(text):

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
#   - theEmail - the actual email
#   - theMessage - where the parsed email message goes
#
# Returns: 
#
#   - True if parsed successfully
#   - False if ran into an issue
#
# -----------------------------------------------------------------------------
def parseHeader(theEmail, theMessage):

    result = True   # assume success
    subject = ""

    # decode the email subject
    try:
        subject, encoding = decode_header(theEmail["Subject"])[0]
    except Exception as e:
        return False
    
    if encoding and isinstance(subject, bytes):
        try:
            # if it's a bytes, decode to str
            subject = subject.decode(encoding)
        except:
            return False
    
    theMessage.subject = subject

    # get the funky IMAP id of the message
    theMessage.id = theEmail.get('Message-ID')

    # get the date from the header
    dateHeader = theEmail.get('Date')

    # remove extra info after tz offset
    if dateHeader:
        try:
            dateHeader = dateHeader.split(' (', 1)[0]
            parsedDate = parser.parse(dateHeader)

            # format the date and time
            dateStr = parsedDate.strftime('%Y-%m-%d')
            timeStr = parsedDate.strftime('%H:%M')

            theMessage.timeStamp = int(parsedDate.timestamp())
            theMessage.dateStr = dateStr
            theMessage.timeStr = timeStr
        except:
            return False

    # set the "toSlug" to the person that was passed in `-m slug`
    if theConfig.me:
        theMessage.toSlugs.append(theConfig.me.slug)

    # decode email sender
    From, encoding = decode_header(theEmail.get("From"))[0]
    if encoding and isinstance(From, bytes):
        try:
            From = From.decode(encoding)
        except:
            pass

    if From:
        emailAddresses = getEmailAddress(From)

    # get the `slug` of the sender
    person = theConfig.getPersonByEmail(emailAddresses)
    if person:
        theMessage.fromSlug = person.slug
    else: 
        pass

    # Decode email recipients (To:)
#    whoTo, encoding = decode_header(theEmail.get("To"))[0]
#    if encoding and isinstance(whoTo, bytes):
#        whoTo = whoTo.decode(encoding)
#
#    toEmailAddresses = getEmailAddress(whoTo)
#
#   slugs = []
#    for emailAddress in toEmailAddresses:
#        slugs.append(theConfig.getPersonByEmail(emailAddress))
#
    # Decode email recipients (Cc:)
#    cc, encoding = decode_header(theEmail.get("Cc"))[0]
#    if encoding and isinstance(cc, bytes):
#        cc = cc.decode(encoding)
#
#   ccEmailAddresses = getEmailAddress(cc)
#    for emailAddress in ccEmailAddresses:
#        slugs.append(theConfig.getPersonByEmail(emailAddress))

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
def downloadAttachment(part, theMessage):

    filename = part.get_filename()

    # download the attachment
    if filename:
        # find the place to put it
        folder = os.path.join(theConfig.sourceFolder, theConfig.attachmentsSubFolder)
        filepath = os.path.join(folder, filename)

        try:
            # download attachment and save it
            open(filepath, "wb").write(part.get_payload(decode=True))

            # create and fill the Attachment object
            theAttachment = attachment.Attachment()
            try:
                theAttachment.id = filename
                theAttachment.fileName = filename
                theAttachment.type = attachment.getMIMEType(filename)
                theAttachment.customFileName = filename
                theMessage.addAttachment(theAttachment)
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
#   - theEmail - the actual email
#   - theMessage - where the parsed email message goes
#
# Returns: nothing
#
# -----------------------------------------------------------------------------
def parseMultiPart(theEmail, theMessage):

    theBody = ""
    contentDisposition = ""
    contentType = ""
    
    # iterate over email parts
    for part in theEmail.walk():

        # extract content type of email
        try:
            contentType = part.get_content_type()
        except:
            pass

        try:
            contentDisposition = str(part.get("Content-Disposition"))
        except:
            pass

        try:
            theBody = part.get_payload(decode=True).decode()
            if theBody and contentType == "text/plain" and "attachment" not in contentDisposition:
                theMessage.body = md(theBody)
        except:
            pass

        if "attachment" in contentDisposition:
            downloadAttachment(part, theMessage)

# -----------------------------------------------------------------------------
#
# Parse the body of the email. Used when it's not a multi-part email.
#
# Parameters:
# 
#   - theEmail - the actual email
#   - theMessage - where the parsed email message goes
#
# Returns: nothing
#
# -----------------------------------------------------------------------------
def parseBody(theEmail, theMessage):

    theBody = ""

    # if the email message is multipart
    if theEmail.is_multipart():
        parseMultiPart(theEmail, theMessage)
    else:
        # extract the content type of the email
        contentType = theEmail.get_content_type()

        # get the email body
        try:
            theBody = theEmail.get_payload(decode=True).decode()
            if theBody:
                theMessage.body = md(theBody)
        
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
#   - theEmail - the actual email
#   - theMessage - where the parsed email message goes
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
def wasForwarded(theEmail, theMessage):

    result = False

    references = theEmail.get('References')
    in_reply_to = theEmail.get('In-Reply-To')

    # if References and In-Reply-To are empty, might be a forwarded message
    if not references and not in_reply_to:
        result = True
    
    # if the subject line starts with "Fwd:" or "fwd" or "Fw:" or "fw", then
    # it was likely a forwarded email
    subject = theMessage.subject
    pattern = re.compile(r'^fw.*:', re.IGNORECASE)

    # avoid "TypeError: cannot use a string pattern on a bytes-like object"
    if isinstance(subject, str) and not bool(pattern.match(subject)):
        result = True
    
    return result

def isEmailHeader(line):
    # check if the line resembles an email header
    return re.match(r'^\s*(From:|Sent:|To:|Cc:|Subject:)', line, re.IGNORECASE) is not None

def joinLines(body):

    # reassemble lines, preserving paragraph breaks
    lines = body.splitlines()
    paragraphs = []
    currentParagraph = ''

    for line in lines:
        if isEmailHeader(line):
            # if the line resembles an email header, start a new paragraph
            if currentParagraph:
                paragraphs.append(currentParagraph.strip())
                currentParagraph = ''
            currentParagraph += line.strip() + '\n'
        elif line.strip():  # if the line is not empty
            currentParagraph += line + ' '
        else:  # if the line is empty, it indicates a new paragraph
            if currentParagraph and not currentParagraph.isspace():
                paragraphs.append(currentParagraph.strip())
            currentParagraph = ''

    # add the last paragraph if there's any
    if currentParagraph and not currentParagraph.isspace():
        paragraphs.append(currentParagraph.strip())

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
#   - theEmail - the actual email
#   - theMessage - where the parsed email message goes
#
# Returns: 
#
#   - True if successful
#   - False if ran into an 
#
# -----------------------------------------------------------------------------
def cleanBody(theEmail, theMessage):

    text = theMessage.body

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

    text = removeReply(text)

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
    text = joinLines(text)

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
    theMessage.body = text

    return result

# parse a specific email and append it to the list Messages collection
def parseEmail(theEmail, theMessage):

    for response in theEmail:
        if isinstance(response, tuple):
            
            try:
                # parse a bytes email into a message object
                thisEmail = email.message_from_bytes(response[1])

                if theConfig.debug:
                    logging.info(f"ID: {id}")

                if parseHeader(thisEmail, theMessage):
                    if (theMessage.fromSlug):
                        parseBody(thisEmail, theMessage)
                        cleanBody(thisEmail, theMessage)
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
def fetchEmails(imap, folder, messages):

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
        response, thisEmail = imap.fetch(str(i), "(RFC822)")

         # create a holder for the parsed email
        theMessage = message.Message()

        parseEmail(thisEmail, theMessage)

        if theMessage.fromSlug:
            count += 1
            messages.append(theMessage)

        # remove the double quotes around the folder name
        if folder.startswith('"') and folder.endswith('"'):
            folder = folder[1:-1]

        # let the user know where processing is at
        status = f"Folder: {folder}  " + f"Countdown: {i-1}  "
        status += f"Found: {count}  Date: {theMessage.dateStr} "
        status = status + ' ' * (120 - len(status))
        print(status, end="\r")
        
        # stop if this message was sent before the start date
        if theMessage.dateStr:
            messageDate = datetime.strptime(theMessage.dateStr, '%Y-%m-%d')
            fromDate = datetime.strptime(theConfig.fromDate, '%Y-%m-%d')
            if messageDate < fromDate:
                continue
 
        if count == theConfig.maxMessages:
            return count

    return count

# -----------------------------------------------------------------------------
#
# Load the emails from the IMAP server.
#
# Parameters:
# 
#   - destFile - not used but needed for the interface
#   - messages - where the Message objects will go
#   - reactions - not used but needed for the interface
#   - theConfig - specific settings 
#
# Returns: the number of messages loaded
#
# -----------------------------------------------------------------------------
def loadMessages(destFile, messages, reactions, theConfig):

    count = 0
    folders = []

    if not (theConfig.imapServer and theConfig.emailAccount and theConfig.password):
        return 0

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL(theConfig.imapServer)

    # authenticate, get the list of folders, fetch the emails
    try:
        imap.login(theConfig.emailAccount, theConfig.password)

        if len(theConfig.emailFolders) > 0:
            folders = theConfig.emailFolders
        else:
            # log the list of folders
            logging.info(imap.list()[1])

            for i in imap.list()[1]:
                l = i.decode().split(' "/" ')
                folders.append(l[1])

        # remove any folders specified in `not-email-folders` 
        # setting and any of it's subfolders
        for xFolder in theConfig.notEmailFolders:
            for yFolder in folders:
                parts = yFolder.split('/')
                if parts[0] == xFolder:
                    folders.remove(yFolder)

        for folder in folders:
            # only fetch emails from folders not in the exclude list
            count += fetchEmails(imap, folder, messages)
        
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

theMessages = []
theReactions = [] 

theConfig = config.Config()

if message_md.setup(theConfig, markdown.YAML_SERVICE_EMAIL, True):

    theConfig.reversed = False

    # needs to be after setup so the command line parameters override the
    # values defined in the settings file
    message_md.getMarkdown(theConfig, loadMessages, theMessages, theReactions)

    print("\n")