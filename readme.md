# 1. Set-OutlookSignatures.ps1
## 1.1. Content
- [1. Set-OutlookSignatures.ps1](#1-set-outlooksignaturesps1)
  - [1.1. Content](#11-content)
  - [1.2. General description](#12-general-description)
  - [1.3. Removing old signatures](#13-removing-old-signatures)
  - [1.4. Outlook signature path](#14-outlook-signature-path)
  - [1.5. Mailboxes](#15-mailboxes)
  - [1.6. Group memberships](#16-group-memberships)
  - [1.7. Parameters](#17-parameters)
    - [1.7.1. SignatureTemplatePath](#171-signaturetemplatepath)
    - [1.7.2. DomainsToCheckForGroups](#172-domainstocheckforgroups)
  - [1.8. Requirements](#18-requirements)
  - [1.9. Error handling](#19-error-handling)
  - [1.10. Run script while Outlook is running](#110-run-script-while-outlook-is-running)
  - [1.11. Signature file format](#111-signature-file-format)
  - [1.12. Signature file naming](#112-signature-file-naming)
    - [1.12.1. Allowed filename tags](#1121-allowed-filename-tags)
  - [1.13. Signature application order](#113-signature-application-order)
  - [1.14. Variable replacement](#114-variable-replacement)
  - [1.15. Outlook Web](#115-outlook-web)
  - [1.16. FAQ](#116-faq)
    - [1.16.1. Why use legacyExchangeDN to find the user behind a mailbox, and not mail or proxyAddresses?](#1161-why-use-legacyexchangedn-to-find-the-user-behind-a-mailbox-and-not-mail-or-proxyaddresses)
## 1.2. General description
Downloads centrally stored signatures, replaces variables, optionally sets default signatures.
Signatures can be applicable to all users, specific groups or specific mail addresses.
Signatures are also set in Outlook Web.
## 1.3. Removing old signatures
The script deletes locally available signatures, if they are no longer available centrally.
Signature created manually by the user are not deleted. The script marks each downloaded signature with a specific HTML tag, which enables this cleaning feature.
## 1.4. Outlook signature path
The Outlook signature path is retrieved from the users registry, so the script is language independent.
The registry setting does not allow for absolute paths, only for paths relative to '%APPDATA%\Microsoft'.
If the relative path set in the registry would be a valid path but does not exist, the script creates it.
## 1.5. Mailboxes
The script only considers primary mailboxes (mailboxes added as additional accounts), no secondary mailboxes.
This is the same way Outlook handles mailboxes from a signature perspective.
The script is created for Exchange environments. Non-Exchange mailboxes can not have group signatures, but common and mailbox specific signatures.
## 1.6. Group memberships
The script considers all groups the currently logged-on user belongs to, as well as all groups the currently processed mailbox belongs to.
For both sets of groups, group membership is searched against the whole Active Directory forest of the currently logged-on user as well as all trusted domains the user can access.
Trusted domains can be modified with the DomainsToCheckForGroups parameter.
Group membership is achieved by querying the tokenGroups attribute, which is not only very fast and resource saving on client and server, but also considers sIDHistory.  
## 1.7. Parameters
### 1.7.1. SignatureTemplatePath
The parameter SignatureTemplatePath tells the script where signature template files are stored.
### 1.7.2. DomainsToCheckForGroups
The parameters tells the script which domains should be used to search for mailbox and user group memberships.
The default value, '*", tells the script to query all trusted domains in the Active Directory forest of the logged-on user.
For a custom list of domains/forests, specify them as comma-separated list of strings: "domain-a.local", "dc=example,dc=com", "domain-b.internal".
The Active Directory forest of the current user is always considered.
## 1.8. Requirements
Requires Outlook and Word, at least version 2010.
The script must run in the security context of the currently logged-on user.
The script must run in PowerShell Full Language mode. Constrained Language mode is not supported, as some features such as BASE64 conversions are not available in this mode or require very slow workarounds.
The path to the signature template files (SignatureTemplatePath) must be accessible by the currently logged-on user. The template files must be at least readable for the currently logged-on user.
## 1.9. Error handling
Error handling is implemented rudimentarily.
## 1.10. Run script while Outlook is running
Outlook and the script can run simultaneously.
New and changed signatures can be used instantly in Outlook.
Changing which signatures are to be used as default signature require an Outlook restart. 
## 1.11. Signature file format
Only Word files with the extension .DOCX are supported as signature template files.
## 1.12. Signature file naming
The script copies every signature file name as-is, with one exception: When tags are defined in the file name, these tags are removed.
Tags must be placed before the file extension and be separated from the base filename with a period.
Examples:
- 'Company external German.docx' -> 'Company external German.htm', no changes
- 'Company external German.[defaultNew].docx' -> 'Company external German.htm', tag(s) is/are removed
- 'Company external [English].docx' ' -> 'Company external [English].htm', tag(s) is/are not removed, because there is no dot before
- 'Company external [English].[defaultNew] [Company-AD All Employees].docx' ' -> 'Company external [English].htm', tag(s) is/are removed, because they are separated from base filename
### 1.12.1. Allowed filename tags
- [defaultNew]
    - Set signature as default signature for new mails
- [defaultReplyFwd]
    - Set signature as default signature for replies and forwarded mails
- [NETBIOS-Domain Group-SamAccountName]
    - Make this signature specific for an Outlook mailbox or the currently logged-on user being a member (direct or indirect) of this group
    - Groups must be available in Active Directory. Groups like 'Everyone' and 'Authenticated Users' only exist locally, not in Active Directory.
- [SMTP address]
    - Make this signature specific for the assigned mail address (all SMTP addresses of a mailbox are considered, not only the primary one)
Filename tags can be combined, so a signature may be assigned to several groups and several mail addresses at the samt time.
## 1.13. Signature application order
Signatures are applied in a specific order: Common signatures first, group signatures second, mail address specific signatures last.
Common signatures are signatures with either no tag or only [defaultNew] and/or [defaultReplyFwd].
Within these groups, signatures are applied alphabetically ascending.
Every centrally stored signature is applied only once, as there is only one signature path in Outlook, and subfolders are not allowed - so the file names have to be unique.
The script always starts with the mailboxes in the default Outlook profile, preferrably with the current users personal mailbox.
## 1.14. Variable replacement
Variables are case sensitive.
Variables are replaced everywhere in the signature files, including href-Links. With this feature, you can not only show mail addresses and telephone numbers in the signature, but show them as links which open a new mail message ("mailto:") or dial the number ("tel:") via a locally installed softphone when clicked.
Custom attributes are supported, see SVSTitelVorne and SVSTitelHinten in the following list.
Available variables:
- Currently logged-on user
    - $CURRENTUSERGIVENNAME$: Given name
    - $CURRENTUSERSURNAME$: Surname
    - $CURRENTUSERNAMEWITHTITLES$: SVSTitelVorne GivenName Surname, SVSTitelHinten (without unnecessary punctuation marks when one of the attributes is not availabe)
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
- Manager of currently logged-on user
    - Same variables as logged-on user, $CURRENTUSERMANAGER[...]$ instead of $CURRENTUSER[...]$
- Current mailbox
    - Same variables as logged-on user, $CURRENTMAILBOX[...]$ instead of $CURRENTUSER[...]$
- Manager of current mailbox
    - Same variables as logged-on user, $CURRENTMAILBOXMANAGER[...]$ instead of $CURRENTMAILBOX[...]$
## 1.15. Outlook Web
If the currently logged-on user has configured his personal mailbox in Outlook, the default signature for new emails is configured in Outlook Web automatically.
If the default signature for new mails matches the one used for replies and forwarded mail, this is also set in Outlook.
If different signatures for new and reply/forward are set, only the new signature is copied to Outlook Web.
If only a default signature for replies and forwards is set, only this new signature is copied to Outlook Web.
If there is no default signature in Outlook, Outlook Web settings are not changed.
All this happens with the credentials of the currently logged-on user, without any interaction neccessary.
## 1.16. FAQ
### 1.16.1. Why use legacyExchangeDN to find the user behind a mailbox, and not mail or proxyAddresses?
The legacyExchangeDN attribute is used to find the user behind a mailbox, because mail and proxyAddresses are not unique in certain Exchange scenarios:
- A separate Active Directory forest for users and Exchange mailboxes: In this case, the mail attribute is usually set in the user forest, although there are no mailboxes in this forest.
- One common mail domain across multiple Exchange organizations: In this case, the address book is very like synchronized between Active Directory forests by using contacts or mail-enabled users, which both will have the SMTP address of the mailbox in the proxyAddresses attribute.
The disadvantage of using legacyEchangeDn is that no group membership information can be retrieved for Exchange mailboxes configured as IMAP or POP accounts in Outlook. This scenario is very rare in Exchange/Outlook enterprise environments. These mailboxes can still receive common and mailbox specific signatures.
