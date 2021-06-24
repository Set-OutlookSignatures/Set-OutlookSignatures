<!-- title: Set-OutlookSignatures.ps1 -->
<!-- omit in toc -->
# <a href="https://github.com/GruberMarkus/Set-OutlookSignatures"><img src="https://raw.githubusercontent.com/GruberMarkus/Set-OutlookSignatures/main/logo/Set-OutlookSignatures%20Logo.png" width="500" title="Set-OutlookSignatures.ps1" alt="Set-OutlookSignatures.ps1"></a>  
Central Outlook for Windows management and deployment script for text signatures and Out of Office (OOF) auto reply messages.  
 
Signatures and OOF messages can be  
- customized with a broad range of variables, including photos from Active Directory,  
- applied to all mailboxes, specific groups or specific addresses,  
- assigned time ranges within which they are valid,  
- set in Outlook Web for the currently logged-on user,  
- centrally managed only or exist along user created signatures (signatures only),  
- copied to an alternate path for easy access on mobile devices not directly supported by this script (signatures only).  
- set as default signature for new mails, or for replies and forwards (signatures only),  
- set as default OOF message for internal or external recipients (OOF messages only).  
  
The script is designed to work in big and complex environments (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects).  
  
  
# Table of Contents  <!-- omit in toc -->
- [1. Requirements](#1-requirements)
- [2. Parameters](#2-parameters)
  - [2.1. SignatureTemplatePath](#21-signaturetemplatepath)
  - [2.2. ReplacementVariableConfigFile](#22-replacementvariableconfigfile)
  - [2.3. DomainsToCheckForGroups](#23-domainstocheckforgroups)
  - [2.4. DeleteUserCreatedSignatures](#24-deleteusercreatedsignatures)
  - [2.5. SetCurrentUserOutlookWebSignature](#25-setcurrentuseroutlookwebsignature)
  - [2.6. SetCurrentUserOOFMessage](#26-setcurrentuseroofmessage)
  - [2.7. OOFTemplatePath](#27-ooftemplatepath)
  - [2.8. AdditionalSignaturePath](#28-additionalsignaturepath)
- [3. Outlook signature path](#3-outlook-signature-path)
- [4. Mailboxes](#4-mailboxes)
- [5. Group membership](#5-group-membership)
- [6. Removing old signatures](#6-removing-old-signatures)
- [7. Error handling](#7-error-handling)
- [8. Run script while Outlook is running](#8-run-script-while-outlook-is-running)
- [9. Signature and OOF file format](#9-signature-and-oof-file-format)
  - [9.1. Signature and OOF file naming](#91-signature-and-oof-file-naming)
  - [9.2. Allowed filename tags](#92-allowed-filename-tags)
- [10. Signature and OOF application order](#10-signature-and-oof-application-order)
- [11. Variable replacement](#11-variable-replacement)
  - [11.1. Photos from Active Directory](#111-photos-from-active-directory)
- [12. Outlook Web](#12-outlook-web)
- [13. FAQ](#13-faq)
  - [13.1. Why use legacyExchangeDN to find the user behind a mailbox, and not mail or proxyAddresses?](#131-why-use-legacyexchangedn-to-find-the-user-behind-a-mailbox-and-not-mail-or-proxyaddresses)
  - [13.2. Which ports are required?](#132-which-ports-are-required)
  - [13.3. Why is Out of Office abbreviated OOF and not OOO?](#133-why-is-out-of-office-abbreviated-oof-and-not-ooo)
  - [13.4. What about the new signature roaming feature Microsoft announced?](#134-what-about-the-new-signature-roaming-feature-microsoft-announced)
  - [13.5. Why DOCX as template format and not HTML? Signatures in Outlook sometimes look different than my DOCX templates.](#135-why-docx-as-template-format-and-not-html-signatures-in-outlook-sometimes-look-different-than-my-docx-templates)
  
  
# 1. Requirements  
Requires Outlook and Word, at least version 2010.  
The script must run in the security context of the currently logged-on user.  
The script must run in PowerShell Full Language mode. Constrained Language mode is not supported, as some features such as BASE64 conversions are not available in this mode or require very slow workarounds. If you use AppLocker or a comparable solution, you may need to digitally sign the PowerShell script.  
The paths to the template files (SignatureTemplatePath, OOFTemplatePath) must be accessible by the currently logged-on user. The template files must be at least readable for the currently logged-on user.  
# 2. Parameters  
## 2.1. SignatureTemplatePath  
The parameter SignatureTemplatePath tells the script where signature template files are stored.  
Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\Signature templates').  
WebDAV paths are supported (https only): 'https<area>://server.domain/SignatureSite/SignatureTemplates' or '\\\\server.domain@SSL\SignatureSite\SignatureTemplates'  
The currently logged-on user needs at least read access to the path.  
Default value: '.\Signature templates'  
## 2.2. ReplacementVariableConfigFile  
The parameter ReplacementVariableConfigFile tells the script where the file defining replacement variables is located.  
Local and remote paths are supported. Local paths can be absolute ('C:\config\default replacement variables.txt') or relative to the script path ('.\config\default replacement variables.txt').  
WebDAV paths are supported (https only): 'https<area>://server.domain/SignatureSite/config/default replacement variables.txt' or '\\\\server.domain@SSL\SignatureSite\config\default replacement variables.txt'  
The currently logged-on user needs at least read access to the file.  
Default value: '.\config\default replacement variables.txt'  
## 2.3. DomainsToCheckForGroups  
The parameters tells the script which domains should be used to search for mailbox and user group membership.  
The default value, '\*' tells the script to query all trusted domains in the Active Directory forest of the logged-on user.  
For a custom list of domains/forests, specify them as comma-separated list of strings: "domain-a.local", "dc=example,dc=com", "domain-b.internal".  
When a domain/forest in the custom list starts with a dash or minus ('-domain-a.local'), this domain is removed from the list.  
The '\*' entry in a custom list is only considered when it is the first entry of the list.  
The Active Directory forest of the currently logged-on user is always considered.  
Default value: '*'  
## 2.4. DeleteUserCreatedSignatures  
Shall the script delete signatures which were created by the user itself? The default value for this parameter is \$false.  
Remark: The script always deletes signatures which were deployed by the script earlier, but are no longer available in the central repository.  
Default value: \$false  
## 2.5. SetCurrentUserOutlookWebSignature  
Shall the script set the Outlook Web signature of the currently logged on user?  
Default value: \$true  
## 2.6. SetCurrentUserOOFMessage  
Shall the script set the Out of Office (OOF) auto reply message of the currently logged on user?  
Default value: \$true  
## 2.7. OOFTemplatePath  
Path to centrally managed Out of Office (OOF) auto reply templates.  
Local and remote paths are supported.  
Local paths can be absolute ('C:\OOF templates') or relative to the script path ('.\OOF templates').  
WebDAV paths are supported (https only): 'https:<area>//server.domain/SignatureSite/OOFTemplates' or '\\\\server.domain@SSL\SignatureSite\OOFTemplates'  
The currently logged-on user needs at least read access to the path.  
Default value: '.\OOF templates'  
## 2.8. AdditionalSignaturePath  
An additional path that the signatures shall be copied to.  
Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.  
This way, the user can easily copy-paste his preferred preconfigured signature for use in a mail app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.  
Local and remote paths are supported.  
Local paths can be absolute ('C:\Outlook signatures') or relative to the script path ('.\Outlook signatures').  
WebDAV paths are supported (https only): 'https:<area>//server.domain/User/Outlook signatures' or '\\\\server.domain@SSL\User\Outlook signatures'  
The currently logged-on user needs at least write access to the path.  
Default value: "\$(\[environment]::GetFolderPath(“MyDocuments”))\Outlook signatures"  
# 3. Outlook signature path  
The Outlook signature path is retrieved from the users registry, so the script is language independent.  
The registry setting does not allow for absolute paths, only for paths relative to '%APPDATA%\Microsoft'.  
If the relative path set in the registry would be a valid path but does not exist, the script creates it.  
# 4. Mailboxes  
The script only considers primary mailboxes (mailboxes added as separate accounts), no secondary mailboxes.  
This is the same way Outlook handles mailboxes from a signature perspective.  
The script is created for Exchange environments. Non-Exchange mailboxes can not have OOF messages or group signatures, but common and mailbox specific signatures.  
# 5. Group membership  
The script considers all groups the currently logged-on user belongs to, as well as all groups the currently processed mailbox belongs to.  
For both sets of groups, group membership is searched against the whole Active Directory forest of the currently logged-on user as well as all trusted domains the user can access.  
The script works fine with linked mailboxes in Exchange resource forest scenarios.  
Trusted domains can be modified with the DomainsToCheckForGroups parameter.  
Group membership is achieved by querying the tokenGroups attribute, which is not only very fast and resource saving on client and server, but also considers sIDHistory.  
# 6. Removing old signatures  
The script always deletes signatures which were deployed by the script earlier, but are no longer available in the central repository. The script marks each processed signature with a specific HTML tag, which enables this cleaning feature.  
Signatures created manually by the user are not deleted by default, this behavior can be changed with the DeleteUserCreatedSignatures parameter.  
# 7. Error handling  
Error handling is implemented rudimentarily.  
# 8. Run script while Outlook is running  
Outlook and the script can run simultaneously.  
New and changed signatures can be used instantly in Outlook.  
Changing which signature is to be used as default signature for new mails for for replies and forwards requires restarting Outlook.   
# 9. Signature and OOF file format  
Only Word files with the extension .DOCX are supported as signature and OOF template files.  
## 9.1. Signature and OOF file naming  
The script copies every signature and OOF file as-is, with one exception: When tags are defined in the file name, these tags are removed.  
Tags must be placed before the file extension and be separated from the base filename with a period.  
Examples:  
- 'Company external German.docx' -> 'Company external German.htm', no changes  
- 'Company external German.\[defaultNew].docx' -> 'Company external German.htm', tag(s) is/are removed  
- 'Company external \[English].docx' ' -> 'Company external \[English].htm', tag(s) is/are not removed, because there is no dot before  
- 'Company external \[English].\[defaultNew] \[Company-AD All Employees].docx' ' -> 'Company external \[English].htm', tag(s) is/are removed, because they are separated from base filename  
## 9.2. Allowed filename tags  
- \[defaultNew] (signature template files only)  
    - Set signature as default signature for new mails  
- \[defaultReplyFwd] (signature template files only)  
    - Set signature as default signature for replies and forwarded mails  
- \[internal] (OOF template files only)  
    - Set template as default OOF message for internal recipients  
    - If neither \[internal] nor \[external] is defined, the template is set as default OOF message for internal and external recipients  
- \[external] (OOF template files only)  
    - Set template as default OOF message for external recipients  
    - If neither \[internal] nor \[external] is defined, the template is set as default OOF message for internal and external recipients  
- \[NETBIOS-Domain Group-SamAccountName], e.g. \[EXAMPLE Domain Users]  
    - Make this template specific for an Outlook mailbox or the currently logged-on user being a member (direct or indirect) of this group  
    - Groups must be available in Active Directory. Groups like 'Everyone' and 'Authenticated Users' only exist locally, not in Active Directory  
- \[SMTP address], e.g. \[office<area>@example.com]  
    - Make this template specific for the assigned mail address (all SMTP addresses of a mailbox are considered, not only the primary one)  
- \[yyyyMMddHHmm-yyyyMMddHHmm], e.g. \[202112150000-202112262359] for the 2021 Christmas season  
    - Make this template valid only during the specific time range (yyyy = year, MM = month, dd = day, HH = hour, mm = minute)  
    - If the script does not run after a template has expired, the template is still available on the client and can be used.  
Filename tags can be combined: A template may be assigned to several groups, several mail addresses and several time ranges, be used as default signature for new e-mails and as default signature for replies and forwards at the same time.  
The number of possible tags is limited by Operating System file name and path length restrictions only. The script works with path names longer than the default Windows limit of 260 characters, even when "LongPathsEnabled" (https://docs.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation) is not active.  
# 10. Signature and OOF application order  
Templates are applied in a specific order: Common tempaltes first, group templates second, mail address specific templates last.  
Templates with a time range tag are only considered if the current system time is in range of at least one of these tags.  
Common templates are templates with either no tag or only \[defaultNew] and/or \[defaultReplyFwd] (\[internal] and \[external] for OOF templates).  
Within these groups, templates are applied alphabetically ascending.  
Every centrally stored signature template is applied only once, as there is only one signature path in Outlook, and subfolders are not allowed - so the file names have to be unique.  
The script always starts with the mailboxes in the default Outlook profile, preferrably with the current users personal mailbox.  
OOF templates are only applied if the Out of Office assistant is currently disabled. If it is currently active or scheduled to be activated in the future, OOF templates are not applied.  
# 11. Variable replacement  
Variables are case sensitive.  
Variables are replaced everywhere, including links, QuickTips and alternative text of images.  
With this feature, you can not only show mail addresses and telephone numbers in the signature and OOF message, but show them as links which open a new mail message ("mailto:") or dial the number ("tel:") via a locally installed softphone when clicked.  
Custom Active directory attributes are supported as well as custom replacement variables, see './config/default replacement variables.txt' for details.  
Per default, './config/default replacement variables.txt' contains the following replacement variables:  
- Currently logged-on user  
    - \$CURRENTUSERGIVENNAME\$: Given name  
    - \$CURRENTUSERSURNAME\$: Surname  
    - \$CURRENTUSERDEPARTMENT\$: Department  
    - \$CURRENTUSERTITLE\$: Title  
    - \$CURRENTUSERSTREETADDRESS\$: Street address  
    - \$CURRENTUSERPOSTALCODE\$: Postal code  
    - \$CURRENTUSERLOCATION\$: Location  
    - \$CURRENTUSERCOUNTRY\$: Country  
    - \$CURRENTUSERTELEPHONE\$: Telephone number  
    - \$CURRENTUSERFAX\$: Facsimile number  
    - \$CURRENTUSERMOBILE\$: Mobile phone  
    - \$CURRENTUSERMAIL\$: Mail address  
    - \$CURRENTUSERPHOTO\$: Photo from Active Directory, see [Photos from Active Directory](#photos-from-active-directory) for details  
    - \$CURRENTUSERPHOTODELETEEMPTY\$: Photo from Active Directory, see [Photos from Active Directory](#photos-from-active-directory) for details  
    - \$CURRENTUSEREXTATTR1\$ to \$CURRENTUSEREXTATTR15\$: Exchange Extension Attributes 1 to 15  
- Manager of currently logged-on user  
    - Same variables as logged-on user, \$CURRENTUSERMANAGER\[...]\$ instead of \$CURRENTUSER\[...]\$  
- Current mailbox  
    - Same variables as logged-on user, \$CURRENTMAILBOX\[...]\$ instead of \$CURRENTUSER\[...]\$  
- Manager of current mailbox  
    - Same variables as logged-on user, \$CURRENTMAILBOXMANAGER\[...]\$ instead of \$CURRENTMAILBOX\[...]\$  
## 11.1. Photos from Active Directory  
The script supports replacing images in signature templates with photos stored in Active Directory.  
When using images in OOF templates, please be aware that Exchange and Outlook do not yet support images in OOF messages.  
  
As with other variables, photos can be obtained from the currently logged-on user, it's manager, the currently processed mailbox and it's manager.  
  
To be able to apply Word image features such as sizing, cropping, frames, 3D effects etc, you have to exactly follow these steps:  
1. Create a sample image file which will later be used as placeholder.  
2. Optionally: If the sample image file name contains one of the following variable names, the script recognizes it and you do not need to add the value to the alternative text of the image in step 4:  
 - \$CURRENTUSERPHOTO\$  
 - \$CURRENTUSERPHOTODELETEEMPTY\$  
 - \$CURRENTUSERMANAGERPHOTO\$  
 - \$CURRENTUSERMANAGERPHOTODELETEEMPTY\$  
 - \$CURRENTMAILBOXPHOTO\$  
 - \$CURRENTMAILBOXPHOTODELETEEMPTY\$  
 - \$CURRENTMAILBOXMANAGERPHOTO\$  
 - \$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY\$  
3. Insert the image into the signature template. Make sure to use "Insert | Pictures | This device" (Word 2019, other versions have the same feature in different menus) and to select the option "Insert and Link" - if you forget this step, a specific Word property is not set and the script will not be able to replace the image.  
4. If you did not follow optional step 2, please add one of the following variable names to the alternative text of the image in Word (these variables are removed from the alternative text in the final signature):  
 - \$CURRENTUSERPHOTO\$  
 - \$CURRENTUSERPHOTODELETEEMPTY\$  
 - \$CURRENTUSERMANAGERPHOTO\$  
 - \$CURRENTUSERMANAGERPHOTODELETEEMPTY\$  
 - \$CURRENTMAILBOXPHOTO\$  
 - \$CURRENTMAILBOXPHOTODELETEEMPTY\$  
 - \$CURRENTMAILBOXMANAGERPHOTO\$  
 - \$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY\$  
5. Format the image as wanted.  
  
For the script to recognize images to replace, you need to follow at least one of the steps 2 and 4. If you follow both, the script first checks for step 2 first. If you provide multiple image replacement variables, \$CURRENTUSER\[..]\$ has the highest priority, followed by \$CURRENTUSERMANAGER\[..]\$, \$CURRENTMAILBOX\[..]\$ and \$CURRENTMAILBOXMANAGER\[..]\$. It is recommended to use only one image replacement variable per image.  
  
The script will replace all images meeting the conditions described in the steps above and replace them with Active Directory photos in the background. This keeps Work image formatting option alive, just as if you would use the "Change picture" function.  
  
If there is no photo available in Active Directory, there are two options:  
- You used the \$CURRENT\[...]PHOTO\$ variables: The sample image used as placeholder is shown in the signature.  
- You used the \$CURRENT\[...]PHOTODELETEEMPTY\$ variables: The sample image used as placeholder is deleted from the signature, which may affect the layout of the remaining signature depending on your formatting options.  
  

Attention: A signature with embedded images has the expected file size in DOCX, HTM and TXT formats, but the RTF format file will be much bigger.  
The signature template 'Test all signature replacement variables.docx' contains several embedded images and can be used for a file comparison:  
- DOCX: 23 KB  
- HTM: 87 KB  
- RTF without workaround: 27.5 MB  
- RTF with workaround: 1.4 MB  
  
The script uses a workaround, but the resulting RTF files are still huge compared to other file types and especially for use in emails. If this is a problem, please either do not use embedded images in the signature template (including photos from Active Directory), or switch to HTML formatted emails.  
  
If you ran into this problem outside this script, please consider modifying the ExportPictureWithMetafile setting as described in https://support.microsoft.com/kb/224663.  
# 12. Outlook Web  
If the currently logged-on user has configured his personal mailbox in Outlook, the default signature for new emails is configured in Outlook Web automatically.  
If the default signature for new mails matches the one used for replies and forwarded mail, this is also set in Outlook.  
If different signatures for new and reply/forward are set, only the new signature is copied to Outlook Web.  
If only a default signature for replies and forwards is set, only this new signature is copied to Outlook Web.  
If there is no default signature in Outlook, Outlook Web settings are not changed.  
All this happens with the credentials of the currently logged-on user, without any interaction neccessary.  
# 13. FAQ  
## 13.1. Why use legacyExchangeDN to find the user behind a mailbox, and not mail or proxyAddresses?  
The legacyExchangeDN attribute is used to find the user behind a mailbox, because mail and proxyAddresses are not unique in certain Exchange scenarios:  
- A separate Active Directory forest for users and Exchange mailboxes: In this case, the mail attribute is usually set in the user forest, although there are no mailboxes in this forest.  
- One common mail domain across multiple Exchange organizations: In this case, the address book is very like synchronized between Active Directory forests by using contacts or mail-enabled users, which both will have the SMTP address of the mailbox in the proxyAddresses attribute.  
The disadvantage of using legacyExchangeDN is that no group membership information can be retrieved for Exchange mailboxes configured as IMAP or POP accounts in Outlook. This scenario is very rare in Exchange/Outlook enterprise environments. These mailboxes can still receive common and mailbox specific signatures and OOF messages.  
## 13.2. Which ports are required?  
Ports 389 (LDAP) and 3268 (Global Catalog), both TCP and UDP, are required to communicate with Active Directory domains. 
The client needs the following ports to access a SMB file share on a Windows server: 137 UDP, 138 UDP, 139 TCP, 445 TCP (for details, see https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc731402(v=ws.11).  
The client needs port 443 to access a WebDAV share (a SharePoint document library, for example).  
## 13.3. Why is Out of Office abbreviated OOF and not OOO?  
Back in the 1980s, Microsoft had a UNIX OS named Xenix ... but read yourself: https://techcommunity.microsoft.com/t5/exchange-team-blog/why-is-oof-an-oof-and-not-an-ooo/ba-p/610191  
## 13.4. What about the new signature roaming feature Microsoft announced?  
Microsoft announced a change in how and where signatures are stored. Basically, signatures are no longer stored in the file system, but in the mailbox itself.  
This is a good idea, as it makes signatures available across devices and avoids file naming conflicts which may appear in current solutions.  
Based on currently available information, the disadvantage is that signatures for shared mailboxes can no longer be personalized, as the latest signature change would be propagated to all users accessing the shared mailbox (which is especially bad when personalized signatures for shared mailboxes are set as default signature).  
Microsoft has stated that only cloud mailboxes support the new feature and that Outlook for Windows will be the only client supporting the new feature for now. I am confident more mail clients will follow soon. Future will tell if the feature will be made available for mailboxes on premises, too.  
Currently, there is no detailed documentation and no API available to programatically access the new feature.  
Until the feature is fully rolled out and an API is available, you can disable the feature with a registry key. This forces Outlook for Windows to use the well-known file based approach and ensures full compatibility with this script.  
For details, please see https://support.microsoft.com/en-us/office/outlook-roaming-signatures-420c2995-1f57-4291-9004-8f6f97c54d15?ui=en-us&rs=en-us&ad=us.  
## 13.5. Why DOCX as template format and not HTML? Signatures in Outlook sometimes look different than my DOCX templates.  
The script uses DOCX as template format, as this seems to be the easiest way to delegate the creation and management of templates to departments such as Marketing or Corporate Communications.  
I am aware that not all Word formatting options are supported in HTML, which can lead to signatures looking a bit different than templates. For example, images may be placed at a different position in the signature compared to the template - this is because the Outlook HTML component only supports the "in line with text" text wrapping option, while Word offers more options.  
The script can be adopted to work with HTML files, but the requirements for these files are harder to fulfill as it is the case with DOCX files:  
- The template must be UTF8 encoded, or at least only contain UTF8 compatible characters  
- The template must be a single file, additional files and folders not be supported  
- Images must either reference a public URL or be part of the template as Base64 encoded string  

Possible approaches for fulfilling these requirements are:  
- Design the template in a HTML editor that supports all features required
- Design the template in Outlook, paste it into Word, save it as "Website, filtered" and run the resulting file through a script that converts the Word output to a single UTF8 encoded HTML file  

If you prefer HTML templates over Word templates, please leave feedback on the script homepage and I will consider it as an option for a future release.  