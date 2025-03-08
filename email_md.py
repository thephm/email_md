import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import re
from dateutil import parser
from email.utils import getaddresses

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

HEADER_TO = 'To'
HEADER_CC = 'Cc'
HEADER_DATE = 'Date'
HEADER_MESSAGE_ID = 'Message-ID'
HEADER_FROM = 'From'
HEADER_REFERENCES = 'References'
HEADER_IN_REPLY_TO = 'In-Reply-To'

ATTACHMENT = 'attachment'
CONTENT_DISPOSITION = 'Content-Disposition'
CONTENT_TYPE_TEXT_PLAIN = 'text/plain'

email_not_found = []

# attribution to https://thepythoncode.com/article/reading-emails-in-python

# use your email provider's IMAP server, you can look for your provider's IMAP server on Google
# or check this page: https://www.systoolsgroup.com/imap/
# use `i <imapServer>`

# Parse email addresses from text string, returning first found address in lowercase
def get_email_address(text):
    """
    Parses a string of email addresses separated by ';' into a collection.

    Args:
        text: The set of email addresses

    Returns:
        The first email address found (lowercase) or False if none found
    """

    if not isinstance(text, str):
        return text

    # regular expression pattern for extracting email addresses
    email_pattern = r'[\w\.-]+@[\w\.-]+'

    # Use the findall method to extract email addresses from the text
    email_addresses = re.findall(email_pattern, text)

    if email_addresses: 
        result = email_addresses[0].lower()  # convert to lowercase
    else:
        result = False

    return result

# Remove quoted replies and other unnecessary text from email body
def remove_reply(text):
    """
    Shortens the body of the email by removing quoted replies (they will be in
    separate dated files already) and other useless text.

    Example:
        Input: On Mon, Oct 9, 2023 at 1:06 PM Bob Smith <bob@smith> wrote:
        Output: String with quoted reply text removed

    Args:
        text: The email body text to clean

    Returns:
        Cleaned text with quotes and unnecessary content removed
    """

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

# Parse email addresses from header and add them to Message object
def parse_addresses(the_email, the_message, direction):
    """
    Parse the email addresses from email into a Message object.

    Args:
        the_email: The actual email
        the_message: Where the parsed email message goes
        direction: FROM, TO, or CC

    Returns:
        bool: True if person was found and email was added to them,
              False if person being ignored or email address not found
    """
    
    result = False

    to_from_cc_header = the_email.get(direction)

    if to_from_cc_header:
        # parse "To" addresses into a list of tuples (name, email)
        parsed_addresses = getaddresses([to_from_cc_header])
        
        # extract just the email addresses from the parsed list
        to_emails = [addr[1] for addr in parsed_addresses if addr[1]]

        # add these emails to the_message (adjust as per your data structure)
        the_message.to_emails = to_emails

        for email_address in the_message.to_emails:
            try: 
                person = the_config.get_person_by_email(email_address.lower())
                # if we found someone and not ignoring them e.g. a mailing list
                if person and not person.ignore:
                    if person.slug not in the_message.to_slugs and person.slug not in the_message.from_slug:
                        the_message.to_slugs.append(person.slug)
                        result = True
                elif not person:
                    email_not_found.append(email_address)
            except Exception as e:
                logging.error(f"{the_config.get_str(the_config.STR_NO_PERSON_WITH_EMAIL)}: {email_address}. Error {e}")

    return result

# Extract and parse email header fields into Message object
def parse_header(the_email, the_message):
    """
    Parse the header of the email into a Message object.

    Args:
        the_email: The actual email
        the_message: Where the parsed email message goes

    Returns:
        bool: True if parsed successfully, False if ran into an issue
    """

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
    the_message.id = the_email.get(HEADER_MESSAGE_ID)

    # get the date from the header
    date_header = the_email.get(HEADER_DATE)

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
    
    # get the to and cc email addresses
    result = parse_addresses(the_email, the_message, HEADER_TO)
    parse_addresses(the_email, the_message, HEADER_CC)

    # decode email sender
    the_from, encoding = decode_header(the_email.get(HEADER_FROM))[0]
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

# Save email attachment to disk and add it to Message object
def download_attachment(part, the_message):
    """
    Download the attachment from the multi-part email and add a corresponding
    Attachment object to the Message object.

    Args:
        part: The part of the email that has the attachment
        the_message: Where the Attachment is added

    Returns:
        None
    """
    filename = part.get_filename()

    # download the attachment
    if filename:
        # find the place to put it
        folder = os.path.join(the_config.output_folder, the_config.people_subfolder)
        folder = os.path.join(folder, the_config.media_subfolder)
        file_path = os.path.join(folder, filename)
        
        # if the folder doesn't exist, create it
        if not os.path.exists(folder):
            # create the folder
            os.makedirs(folder)

        try:
            # download attachment and save it
            open(file_path, "wb").write(part.get_payload(decode=True))

            # create and fill the Attachment object
            the_attachment = attachment.Attachment()
            try:
                the_attachment.id = filename
                the_attachment.filename = filename
                the_attachment.type = the_config.get_mime_type(filename)
                the_attachment.custom_filename = filename
                the_message.add_attachment(the_attachment)
            except Exception as e:
                logging.error("download_attachment: {e}")
                pass

        except Exception as e:
            logging.error("{the_config.get_str(STR_COULD_NOT_CREATE_MEDIA_FOLDER)}: {file_path}") 
            pass

# Process all parts of a multi-part email message 
def parse_multi_part(the_email, the_message):
    """
    If the email is a multi-part email, parse each part.

    Args:
        the_email: The actual email
        the_message: Where the parsed email message goes

    Returns:
        None
    """

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
            content_disposition = str(part.get(CONTENT_DISPOSITION))
        except:
            pass

        try:
            the_body = part.get_payload(decode=True).decode()
            if the_body and content_type == CONTENT_TYPE_TEXT_PLAIN and ATTACHMENT not in content_disposition:
                the_message.body = md(the_body)
        except:
            pass

        if ATTACHMENT in content_disposition:
            download_attachment(part, the_message)

# Extract and process email body content
def parse_body(the_email, the_message):
    """
    Parse the body of the email. Used when it's not a multi-part email.

    Args:
        the_email: The actual email
        the_message: Where the parsed email message goes

    Returns:
        None
    """

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

def is_email_header(line):
    # check if the line resembles an email header
    return re.match(r'^\s*(From:|Sent:|To:|Cc:|Subject:)', line, re.IGNORECASE) is not None

# Join lines within paragraphs while preserving email headers
def join_lines(body):
    """
    Join lines within paragraphs, excluding email headers.
    
    Args:
        body (str): The email body contents
        
    Returns:
        str: The resulting body with properly joined lines
    """

    # reassemble lines, preserving paragraph breaks
    lines = body.splitlines()
    paragraphs = []
    current_paragraph = ''

    for line in lines:
        if is_email_header(line):
            # if the line resembles an email header, start a new paragraph
            if current_paragraph:
                paragraphs.append(current_paragraph.strip())
                current_paragraph = ''
            paragraphs.append(current_paragraph.strip())
        elif line.strip():  # if the line is not empty, add it to the current paragraph
            current_paragraph += line.strip() + '\n'
        elif line.strip():  # if the line is not empty
            current_paragraph += line + ' '
        else:  # if the line is empty, it indicates the end of the current paragraph
            if current_paragraph and not current_paragraph.isspace():
                paragraphs.append(current_paragraph.strip())
            current_paragraph = ''

    # add the last paragraph if there's any
    if current_paragraph:
        paragraphs.append(current_paragraph.strip())

    # join lines within each paragraph, excluding email headers
    body = '\n\n'.join(para for para in paragraphs)

    return body

def clean_yahoo_text(text):
    """
    Remove lines containing Yahoo promotional content
    
    Args:
        text (str): Input text to clean
    
    Returns:
        str: Cleaned text
    """
    # Regex to remove lines
    lines = [line for line in text.split('\n') 
             if not (
                 # Remove lines with Yahoo promotional content
                 (re.search(r'yahoo\!?', line, re.IGNORECASE) and 
                  (re.search(r'yahoo\.com', line, re.IGNORECASE) or 
                   re.search(r'platinum|mail\s*plus|sign\s*up|greetings|e-?cards|messages|photos|finance|personals|home\s*page|security|mobile|farechase|travel|autos|beta|discover|sports|small\s*business|spam|search|my\s*yahoo|resources|holiday|video|email|dsl|music', line, re.IGNORECASE))) or
                 # Remove "Do you Yahoo!?" lines
                 re.search(r'do\s*you\s*yahoo\!?\?', line, re.IGNORECASE) or
                 # Remove standalone angle brackets and empty lines
                 re.match(r'^>\s*$', line) or
                 # Remove job and resume related lines
                 re.search(r'(find|post)\s*(a)?\s*(job|resume)', line, re.IGNORECASE) or
                 # Remove lines about accessing Yahoo services
                 re.search(r'(listen|access|check)\s*.*\s*(yahoo\!? (mail|messages))', line, re.IGNORECASE) or
                 # Remove lines about Yahoo service features
                 re.search(r'(new|more)\s*yahoo\!?\s*(photos|mail)\s*-\s*(easier|better|simpler)', line, re.IGNORECASE) or
                 # Remove lines about Yahoo services with specific actions
                 re.search(r'yahoo\!?\s*(finance|mail):\s*(get|file|check|track)', line, re.IGNORECASE) or
                 # Remove lines highlighting Yahoo service improvements
                 re.search(r'(new\s*and\s*improved|improved)\s*yahoo\!?\s*mail\s*-\s*.*\s*(messages|email)', line, re.IGNORECASE) or
                 # Remove registration or declaration lines
                 re.search(r'(declare\s*yourself|register)\s*.*\s*(online|today)', line, re.IGNORECASE) or
                 # Remove ad posting and personal ad lines
                 re.search(r'(post|place)\s*(your)?\s*(free)?\s*(ad|personal)', line, re.IGNORECASE) or
                 # Remove lines with Yahoo URLs for personals or ads
                 re.search(r'http://.*yahoo\.(com|ca)/*(personals|ads)', line, re.IGNORECASE) or
                 # Remove lines about setting Yahoo as home page
                 re.search(r'(start|begin)\s*your\s*day\s*with\s*yahoo\!?', line, re.IGNORECASE) or
                 # Remove lines encouraging making Yahoo home page
                 re.search(r'make\s*it\s*your\s*(default\s*)?home\s*page', line, re.IGNORECASE) or
                 # Remove lines about Yahoo Mail security
                 re.search(r'yahoo\!?\s*mail\s*-\s*(you)\s*care\s*about\s*(security)', line, re.IGNORECASE) or
                 # Remove lines about mobile Yahoo Mail
                 re.search(r'(take|get)\s*yahoo\!?\s*mail\s*.*\s*(mobile\s*phone|phone)', line, re.IGNORECASE) or
                 # Remove lines about Yahoo FareChase or travel services
                 re.search(r'yahoo\!?\s*farechase:\s*(search|find)\s*.*\s*(travel\s*sites|sites)', line, re.IGNORECASE) or
                 # Remove lines about Yahoo Autos
                 re.search(r'(find|search)\s*your\s*next\s*(car|vehicle)\s*at\s*yahoo\!?\s*canada?\s*autos', line, re.IGNORECASE) or
                 # Remove lines about Yahoo Mail beta
                 re.search(r'(everyone\s*is\s*raving|raving)\s*about\s*.*\s*yahoo\!?\s*mail\s*beta', line, re.IGNORECASE) or
                 # Remove "Discover Yahoo!" and similar promotional lines
                 re.search(r'(discover\s*yahoo\!?|get\s*on-the-go)\s*.*\s*(scores|quotes|news)', line, re.IGNORECASE) or
                 # Remove "Check it out!" type lines
                 re.search(r'check\s*it\s*out\!?', line, re.IGNORECASE) or
                 # Remove Yahoo Sports fantasy lines
                 re.search(r'(rekindle\s*the\s*rivalries|sign\s*up\s*for)\s*.*\s*(fantasy\s*football)', line, re.IGNORECASE) or
                 # Remove Yahoo Small Business promotional lines
                 re.search(r'have\s*a\s*(huge|big)\s*year\s*through\s*yahoo\!?\s*small\s*business', line, re.IGNORECASE) or
                 # Remove Yahoo Mail spam protection lines
                 re.search(r'(tired\s*of\s*spam|spam\s*protection)\s*.*\s*yahoo\!?\s*mail', line, re.IGNORECASE) or
                 # Remove Yahoo Search movie showtime lines
                 re.search(r'(find\s*a\s*flick|movie\s*showtime)\s*.*\s*yahoo\!?\s*search', line, re.IGNORECASE) or
                 # Remove My Yahoo promotional lines
                 re.search(r'(meet\s*the|try\s*it)\s*.*\s*my\s*yahoo\!?', line, re.IGNORECASE) or
                 # Remove Yahoo Small Business resources lines
                 re.search(r'yahoo\!?\s*small\s*business\s*-\s*(try|check\s*out)\s*.*\s*(resources?\s*site)', line, re.IGNORECASE) or
                 # Remove holiday greetings lines
                 re.search(r'send\s*your\s*.*\s*(free)\s*.*\s*(holiday\s*greetings|greetings)', line, re.IGNORECASE) or
                 # Remove Yahoo Mail video email lines
                 re.search(r'send\s*.*\s*(free)\s*.*\s*(video\s*emails?)\s*in\s*yahoo\!?\s*mail', line, re.IGNORECASE) or
                 # Remove Yahoo Mail promotional lines with "Try FREE"
                 re.search(r'(try\s*free)\s*yahoo\!?\s*mail\s*-\s*.*\s*(greatest\s*free\s*email)', line, re.IGNORECASE) or
                 # Remove Yahoo Greetings e-cards lines
                 re.search(r'yahoo\!?\s*greetings\s*-\s*send\s*.*\s*(free\s*e-?cards|e-?cards)', line, re.IGNORECASE) or
                 # Remove SBC Yahoo! DSL promotional lines
                 re.search(r'sbc\s*yahoo\!?\s*dsl\s*-\s*now\s*.*\s*\$\d+\.\d+\s*per\s*month', line, re.IGNORECASE) or
                 # Remove Yahoo Mail improved/new feature lines with message size
                 re.search(r'(new\s*and\s*improved)\s*yahoo\!?\s*mail\s*-\s*send\s*\d+\s*mb\s*messages', line, re.IGNORECASE) or
                 # Remove Yahoo Photos new feature lines
                 re.search(r'new\s*yahoo\!?\s*photos\s*-\s*.*\s*(easier|simpler)\s*(uploading|sharing)', line, re.IGNORECASE) or
                 # Remove Yahoo Music launch lines
                 re.search(r'(launch|start)\s*-\s*your\s*yahoo\!?\s*music\s*experience', line, re.IGNORECASE) or
                 # Remove Yahoo Shopping promotional lines
                 re.search(r'>\s*(>\s*)*\s*yahoo\!?[\s\w]*\s*shopping\s*[-:]\s*.*', line, re.IGNORECASE)
             )]
    
    return '\n'.join(lines).strip()

def clean_body(the_email, the_message):
    """
    Remove extra stuff from the body of the email like "Sent from my iPhone" and 
    any reply-to text from the email being replied to.

    Performs extensive text cleaning including:
    - Removing HTML and styling artifacts
    - Cleaning up quoted text
    - Normalizing line breaks and spacing
    - Removing common email client signatures
    - Formatting headers and message boundaries

    Args:
        the_email: The actual email
        the_message: Where the parsed email message goes

    Returns:
        bool: True if successful, False if ran into errors
    """

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
    text = text.replace("P {margin-top:0;margin-bottom:0;}", "")

    try:
        pattern1 = re.compile(re.escape("#") + '.*?' + re.escape("{margin:0;}"), re.DOTALL)
        text = pattern1.sub('', text)
    except:
        pass
    
    try:
        pattern2 = re.compile(re.escape("#") + '.*?' + re.escape("NoSpacing"), re.DOTALL)
        text = pattern2.sub('', text)
    except:
        pass

    try:
        # get rid of [External]/[Externe]
        pattern3 = re.compile(r'\[External\]/\[Externe\]', re.IGNORECASE)
        text = pattern3.sub('', text)
    except:
        pass
    
    try:
        # regular expression pattern to match variations of "Sent from my iPhone" etc.
        pattern4 = re.compile(r'Sent from .*|Get (Outlook for iOS|.*? for Android)|Sent via .*', re.IGNORECASE)
        text = pattern4.sub('', text)
    except:
        pass
    
    try:
        text = remove_reply(text)
    except:
        pass

    try:
        # add backticks around text with "#" so they aren't seen as tags in Obsidian e.g. `#bob`
        text = re.sub(r'#([^\s\)\]\.]+)', r'`#\1`', text)
    except:
        pass

    try:
        # remove "\\_\\_\\_\\_\\_\\_\\_" of any length
        pattern5 = re.compile(r'\\_+', re.MULTILINE)
        text = pattern5.sub('', text)
    except:
        pass

    try:
        # remove "\\\\" of any length
        pattern6 = re.compile(r'^\\\\$', re.MULTILINE)
        text = pattern6.sub('', text)
    except:
        pass
    
    try:
        # add a blank line before "On Feb 22, 2018, at 8:18 PM, Bob Smith wrote:"
        text = re.sub(r'(On .*? wrote:)', r'\n\1\n', text, flags=re.DOTALL)
    except:
        pass

    try:
        # remove backslash and asterisk around "From," "Sent," and "To"
        text = re.sub(r'\*\*(From|Cc|Sent|To|Subject)\:\*\*', r'\n\1:', text, flags=re.IGNORECASE)
    except:
        pass

    # remove any HTML
    try:
        text = md(text.strip())
        result = True
    except Exception as e:
        pass

    try:
        # reassemble lines to avoid word splitting
        text = join_lines(text)
    except:
        pass

    try:
        # get rid of "p.MsoNormal,p.MsoNoSpacing{margin:0}"
        text = text.replace('p.MsoNormal,p.MsoNoSpacing{margin:0}', ' ') 
        pattern7 = re.compile(re.escape("p.") + '(.*?)' + re.escape("{margin: ?0;}"), re.DOTALL)
        text = pattern7.sub('', text)
    except:
        pass

    try:
        # remove leading spaces, likely vestiges from other substitutions above
        text = re.sub(r'^\s*', '', text, flags=re.MULTILINE)
    except:
        pass

    try:
        # get rid of extra newlines
        text = re.sub(r'\n{3,}', '\n\n', text)
    except:
        pass    
    
    try:
        text = text.replace('| | | --- | |', ' ') 
        text = text.replace('| | --- | |', ' ') 
    except:
        pass

    try:
        # add a blank line before lines starting with From:
        text = re.sub(r'^(From:)', r'\n\1', text, flags=re.MULTILINE)
    except:
        pass
    
    try:
        # add a blank line after lines starting with Subject:
        text = re.sub(r'(Subject: .*)\n+', r'\1\n', text)
    except:
        pass

    try:
        # add a line before "---------- Forwarded message ---------"
        text = re.sub(r'\n?-*\s*-?Forwarded message?-*\s*-', '\n\n*- Forwarded message *-', text, flags=re.IGNORECASE)
    except:
        pass

    try:
        # ensure exactly one blank line before and after "Original Message"
        text = re.sub(r'\n*\s*-{0,3}\s*Original Message\s*-{0,3}\s*\n*', '\n\n-- Original Message --\n\n', text)
    except:
        pass

    try:
        # remove lines between and including "-=-=-=-=-=-=-=-=-=-=-=-"
        text = re.sub(r'-=-=-=-=-=-=-=-=-=-=-=-.*?-=-=-=-=-=-=-=-=-=-=-=-\n?', '', text, flags=re.DOTALL)
    except:
        pass

    try:
        # remove everything after "Join Zoom Meeting"
        text = re.sub(r'Join Zoom Meeting.*', 'Join Zoom Meeting', text, flags=re.DOTALL)
    except:
        pass

    try:
        # replace lines containing "======================================================================" with 9 fewer "="
        text = re.sub(r'^(=+)$', lambda m: '=' * (len(m.group(1)) - 9), text, flags=re.MULTILINE)
    except:
        pass

    try:
        # combine lines that don't start with "> " or a prompt and don't end with a period
        text = re.sub(r'([^\.\!\?])\n(?!>|[a-zA-Z]+:\s)', r'\1 ', text)
    except:
        pass

    try:
        # add blank lines between paragraphs
        text = re.sub(r'(\n)(?=\S)', r'\1\n', text)
    except:
        pass

    try:
        # ensure lines between quoted paragraphs also have a quote ">"
        text = re.sub(r'(?<=^> .*)\n(?=> )', '>\n', text, flags=re.MULTILINE)
    except:
        pass

    try:
        # remove text that starts and ends with any quantity of asterisks and contains "This e-mail"
        text = re.sub(r'\*+.*?This e-mail.*?\*+', '', text, flags=re.DOTALL)
    except:
        pass

    try:
        # remove "> > >", "> >", or "> " from the end of lines
        text = re.sub(r'(> > >|> >|> )$', '', text, flags=re.MULTILINE)
    except:
        pass

    try:
        # ensure lines with "--------------------------------" preceded and/or followed by a space or any number of dashes greater than 2 have a blank line before and after them
        text = re.sub(r'\s*-{2,}\s*--------------------------------\s*-{2,}\s*', '\n\n--------------------------------\n\n', text)
    except:
        pass

    try:
        # remove "_____" (or more or fewer underscores)
        text = re.sub(r'_+', '', text)
    except:
        pass

    try:
        # deal with "Date: Tuesday, December 17, 2024 10:20 AM Blah blah"
        date_pattern = re.compile(r'(Date: [A-Za-z]+, [A-Za-z]+ \d{1,2}, \d{4} \d{1,2}:\d{2} [APM]{2})(?=\S)', re.IGNORECASE)
        text = date_pattern.sub(r'\1\n', text)
    except:
        pass

    # remove any extra blank lines that might be introduced
    text = re.sub(r'\n{3,}', '\n\n', text)

    # split "> --- Bob Smith wrote:"
    text = re.sub(
        r"(.*?)(?:\\s*>*)\\s*-{2,}\\s*(.*?\\s*wrote:)",
        r"\\1\\n\\2",
        text,
        flags=re.IGNORECASE | re.MULTILINE | re.DOTALL
    )

    text = re.sub(r'>> Chat with friends online, try MSN Messenger: http://messenger\.msn\.com', '', text)
    
    # remove AVG text
    text = re.sub(r'Checked by AVG.*?message\.', '', text)
    text = re.sub(r'> >---\s*> >Incoming mail is certified Virus Free\..*?Release Date: \d{2}/\d{2}/\d{4}', '', text, flags=re.DOTALL)

    # remove "> Off to school, going on a trip, or moving? Windows Live (MSN) Messenger..."
    text = re.sub(r'Off to school.*?/[^/]+$', '', text)

    # replace
    #   > -- Original Message -- 
    #   > 
    #   > -- 
    # 
    # with
    #   -- Original Message --
    text = re.sub(r'-- Original Message --\s*> --\s*', '-- Original Message --', text)

    # remove "-- " from "-- hi " in lines
    text = re.sub(r">\s*-+\s*hi(?=\s|$)", "> hi", text)
    
    # get rid of extra quoted lines with more flexibility
    text = re.sub(r'\n\n+(?=>)', '\n', text)

    text = clean_yahoo_text(text)

    # FINALLY, ready to put the new-and-improved body in the message 🤣
    the_message.body = text

    return result

# Parse single email and extract header/body into Message object
def parse_email(this_email, the_message):
    """
    Parse a specific email and clean up its contents.

    Processes both the header and body, stores results in the provided Message object.
    Only parses emails from known senders (those with a from_slug).

    Args:
        this_email: The email to be parsed
        the_message: Where the parsed email message goes

    Returns:
        bool: True if email was parsed successfully, 
              False if person was ignored or there was an issue
    """

    result = False

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
                        result = True
                        clean_body(this_email, the_message)
            except:
                pass
        
    return result

# Retrieve and parse all emails from specified IMAP folder
def fetch_emails(imap, folder, messages):
    """
    Load the emails from a specific IMAP server folder.

    Retrieves emails one by one from the specified folder, parses them,
    and appends successfully parsed messages to the messages list.
    Stops when either all emails are processed, max_messages is reached,
    or a message predating from_date is encountered.

    Note: Originally considered filtering by specific email addresses
    (e.g. 'HEADER FROM "spongebob@gmail.com"') but decided against it.

    Args:
        imap: The IMAP connection
        folder: The folder name e.g. "INBOX"
        messages: Where the parsed Message objects are appended

    Returns:
        int: The number of emails successfully parsed
    """

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

        result = parse_email(this_email, the_message)

        if result and the_message.from_slug:
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
        try:
            if the_message.date_str:
                message_date = parser.parse(the_message.date_str)
                from_date = parser.parse(the_config.from_date)
                if message_date < from_date:
                    continue
        except Exception as e:
            logging.error(f"{the_config.get_str(the_config.STR_DATE_STRING_DOES_NOT_MATCH_FORMAT)}: {the_message.date_str}. Error {e}")

        if count >= the_config.max_messages:
            return count

    return count

# Connect to IMAP server and load emails from all accessible folders
def load_messages(dest_file, messages, reactions, the_config):
    """
    Load the emails from the IMAP server.

    Connects to the IMAP server, authenticates, and retrieves emails from all
    accessible folders (except those in not-email-folders). Some parameters
    are unused but required by the interface.

    Args:
        dest_file: Not used but needed for the interface
        messages: Where the Message objects will go
        reactions: Not used but needed for the interface
        the_config: Specific settings

    Returns:
        int: The number of messages loaded
    """

    count = 0
    folders = []

    if not (the_config.imap_server and the_config.email_account and the_config.password):
        return 0

    # create an IMAP4 class with SSL 
    try:
        imap = imaplib.IMAP4_SSL(the_config.imap_server)
    except Exception as e:
        logging.error("load_messages: {e}")
        return 0
    
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

    if len(email_not_found):
        print(the_config.get_str(the_config.STR_THESE_EMAIL_ADDRESSES_NOT_FOUND))

        for email_address in email_not_found:
            print(email_address)