# email_md

Tool to pull IMAP emails and convert them to atomic Markdown files.

## Dependencies

The code in this repo relies heavily on my [message_md](https://github.com/thephm/message_md/tree/main/config) classes which contain generic `Message`, `Person`, `Group` and other classes and the methods to convert messages to Markdown files. Be sure to read the [README](https://github.com/thephm/message_md/blob/main/README.md) and the configuration [guide](https://github.com/thephm/message_md/blob/main/docs/guide.md) for that repo first.

Also relies on markdownify so do this:

```bash
pip install markdownify
```

Refer to your email provider's help pages to find out how to get an application password. If they require two-factor authentication (2FA), then this tool won't work for you.

### Gmail

To setup an application password for Gmail, visit [Sign in with app passwords](https://support.google.com/accounts/answer/185833?hl=en)

For the IMAP server use: `imap.gmail.com`

### Fastmail

Here are the steps to setup an application password for Fastmail.

1. Click "Settings" then "Privacy and Security"
2. Click "New App Password"
3. Give it a name like "email_md"
4. Choose "Mail, Contacts, & Calendars" (or just "IMAP")
5. Click "Generate Password"
6. Under "Your new password for email_md is:" copy that password
7. Put the password somewhere secure like a password keeper, I use BitWarden

Figure out the IMAP server address and port. 

For Fastmail it's `imap.fastmail.com` and Port `993`.

## Setting up the config files

The next step is to configure this tool.

### email server and account

In the `config.json` file, set the following fields. For demonstration purposes, I've used Fastmail's settings:

In this example, emails from only two folders would be fetched: `INBOX` and `Sent Items`:

```
    "imap-server": "imap.fastmail.com",
    "email-folders": "INBOX;Sent Items",
    "not-email-folders": "",
    "email-account": "993",
```

To use all email folders and exclude specific ones (and their sub-folders), use the `not-email-folders` setting:

```
    "imap-server": "imap.fastmail.com",
    "email-folders": "",
    "not-email-folders": "Spam;Shopping;Trash;Health",
    "email-account": "993",
```

Which can also be overriden on the command line:

Command line | Alternate | Description
--- | --- | ---
| `-i` | `--imap` | IMAP server address
| `-r` | `--folders` | IMAP folders to retrieve from
| `-e` | `--email` | email address to retrieve from

### People and groups

You'll need to define each person that you communicate with in `people.json` and the groups in `groups.json`. This way the tool can associate each message with the person that sent it and who it was sent to.

Samples of these configuration files are in the [message_md](https://github.com/thephm/message_md/tree/main/config) repo upon which this tool depends.

This part is tedious the first time and needs to be updated when you add new contacts, i.e. a pain.

## Using email_md

Once you've configured the tool and the `people.json` file is setup, you're ready to run the tool.

The command line options are described in the [message_md](https://github.com/thephm/message_md/tree/main/config) repo.

Example:

```bash
python3 email_md.py -c ../../dev-output/config -s ../../dev-output -o ../../dev-output -d -e spongebob@ownmail.net -p lifeisahighway2! -i imap.fastmail.com -m spongebob -x 20 -b 2024-01-01
```

where:

- `c`onfig settings are in `../../dev-output/config`
- `s`ource folder is `../../dev-output`
- `o`utput the Markdown files to `../../dev-output`
- `d`ebug messages are enabled
- `e`mail address is `spongebob@ownmail.net`
- `p`assword for the email is `lifeisahighway2!`
- ema`i`l server is `imap.fastmail.com`
- `m`y slug is `spongebob`
- ma`x`imum of `20` messages should be converted
- `b`egin the export from `2024-01-01`

