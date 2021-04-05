# Set-OutlookSignatures.ps1
## General feature
Downloads centrally stored signatures, replaces variables, optionally sets default signatures.
Signatures can be applicable to all users, specific groups or specific mail addresses.
## Removing old signatures
The script deletes locally available signatures, if they are no longer available centrally.
Signature created manually by the user are not deleted. The script marks each downloaded signature with a specific HTML tag, which enables this cleaning feature.
## Outlook signature path
The Outlook signature path is retrieved from the users registry, so the script is language independent.
## Mailboxes
The script only considers primary mailboxes (mailboxes added as additional accounts), no secondary mailboxes.
This is the same way Outlook handles mailboxes from a signature perspective.
## Requirements
Requires Outlook and Word, at least version 2010.
## Error handling
Error handling is implemented rudimentarily.
## Execution information
Outlook and the script can run simultaneously. New and changed signatures are available instantly, but changed settings which signatures are to be used as default signature require an Outlook restart. 
## Signatur file format
Only HTML files with the extension .HTM are supported.
The script does not support .MSG, .EML, .MHT, .MHTML or .HTML files.
The files must be UTF8 encoded, or at least only contain UTF8 compatible characters.
The files must be single HTML file, additional files are not supported.
Graphics must either be embedded via a pulic URL, or be part of the HTML code as base64 encoded string.
Possible approaches
- Design the mail body in a HTML editor and use links to images, which can be accessed from the internet
- Design the mail body in a HTML editor that is able to convert image links to inline Data URI Base64 strings
- Design the mail body in Outlook.
    - Copy the mail body, paste it into Word and save it as “Website, filtered” (this reduces lots of Office HTML overhead)
    - This creates a .HTM file and a folder containing pictures (if there are any)
- Use ConvertTo-SingleFileHTML.ps1 (comes with this script) or tools such as https://github.com/zTrix/webpage2html to convert image links to inline Data URI Base64 strings
## Signature file naming
The script copies every signature file name as-is, with one exception: When tags are defined in the file name, these tags are removed.
Tags must be placed before the file extension and be separated from the base filename with a period.
Examples:
- 'ITSV extern.htm' -> 'ITSV extern.htm', no changes
- 'ITSV extern.[defaultNew].htm' -> 'ITSV extern.htm', tag(s) is/are removed
- 'ITSV extern [deutsch].htm' ' -> 'ITSV extern [deutsch].htm', tag(s) is/are not removed, because they are separated from base filename
- 'ITSV extern [deutsch].[defaultNew] [ITSV-SVS Alle Mitarbeiter].htm' ' -> 'ITSV extern [deutsch].htm', tag(s) is/are removed, because they are separated from base filename
### Allowed filename tags
- [defaultNew]
    - Set signature as default signature for new messages
- [defaultReplyFwd]
    - Set signature as default signature für reply for forwarded mails
- [NETBIOS-Domain Group-SamAccountName]
    - Make this signature specific to any member (direct or indirect) of this group
- [SMTP address]
    - Make this signature specific for the assigned mail address (all SMTP addresses of a mailbox are considered, not only the primary one)
Filename tags can be combined, so a signature may be assigned to several groups and several mail addresses at the samt time.
## Signature application order
Signatures are applied in a specific order: Common signatures first, group signatures second, mail address specific signatures last.
Common signatures are signatures with either no tag or only [defaultNew] and/or [defaultReplyFwd].
Within these groups, signatures are applied alphabetically ascending.
Every centrally stored signature is applied only once, as there is only one signature path in Outlook, and subfolders are not allowed - so the file names have to be unique.
The script always starts with the mailboxes in the default Outlook profile, preferrably with the current users personal mailbox.
## Variable replacement
Variables are case sensitive. Variables are replaced everywhere in the .HTM file, including href-Links.
With this feature, you can not only show mail addresses and telephone numbers in the signature, but show them as a link which upon being clicked on opens a new mail message ("mailto:") or dials the number ("tel:") via a locally installed softphone.
Available variables:
- Currently logged on user
    - $CURRENTUSERGIVENNAME$: Given name
    - $CURRENTUSERSURNAME$: Surname
    - $CURRENTUSERNAMEWITHTITLES$: SVSTitelVorne Given Name Surname, SVSTitelHinten (ohne überflüssige Satzzeichen bei fehlenden Einzelattributen)
    - $CURRENTUSERDEPARTMENT$: Department
    - $CURRENTUSERTITLE$: Title
    - $CURRENTUSERSTREETADDRESS$: StreetM
    - $CURRENTUSERPOSTALCODE$: Postal code
    - $CURRENTUSERLOCATION$: Location
    - $CURRENTUSERCOUNTRY$: Country
    - $CURRENTUSERTELEPHONE$: Telephone number
    - $CURRENTUSERFAX$: Facsimile number
    - $CURRENTUSERMOBILE$: Mobile phone
    - $CURRENTUSERMAIL$: Mail address
- Manager of currently logged on user
    - Same variables as logged on user, $CURRENTUSERMANAGER[...]$ instead of $CURRENTUSER[...]$
- Current mailbox
    - Same variables as logged on user, $CURRENTMAILBOX[...]$ instead of $CURRENTUSER[...]$
- Manager of current mailbox
    - Same variables as logged on user, $CURRENTMAILBOXMANAGER[...]$ instead of $CURRENTMAILBOX[...]$
## Outlook Web
If the currently logged on user has configured his personal mailbox in Outlook, the default signature for new emails is configured in Outlook Web automatically.
If the default signature for new mails matches the one used for replies and forwarded mail, this is also set in Outlook.
If different signatures for new and reply/forward are set, only the new signature is copied to Outlook Web.
If only a default signature for replies and forwards is set, only this new signature is copied to Outlook Web.
If there is no default signature in Outlook, Outlook Web settings are not changed.
All this happens with the credentials of the currently logged on user, without any interaction neccessary.
# ConvertTo-SingleFileHTML.ps1
Script to embed locally available files into single HTML file as Base64 strings.
Can only embed files stored on the local computer.
## Input
Input can be a HTML file or folder.
If folder, all .htm and .html files directly in this folder are considered.
## File format
Every single HTML file must be UTF-8 encoded.