<#
.SYNOPSIS
Set-OutlookSignatures XXXVersionStringXXX
Email signatures and out-of-office replies for Exchange and all of Outlook: Classic and New, Windows, Web, Mac, Linux, Android, iOS

.DESCRIPTION
With Set-OutlookSignatures, signatures and out-of-office replies can be:
- Generated from **templates in DOCX or HTML** file format
- Customized with a **broad range of variables**, including **photos**, from Active Directory and other sources
  - Variables are available for the **currently logged-on user, this user's manager, each mailbox and each mailbox's manager**
  - Images in signatures can be **bound to the existence of certain variables** (useful for optional social network icons, for example)
- Designed for **barrier-free accessibility** with custom link and image descriptions for screen readers and comparable tools
- Applied to all **mailboxes (including shared mailboxes¹)**, specific **mailbox groups**, specific **email addresses** (including alias and secondary addresses), or specific **user or mailbox properties**, for **every mailbox across all Outlook profiles (Outlook, New Outlook¹, Outlook Web¹)**, including **automapped and additional mailboxes¹**
- Created with different names from the same template, **one template can be used for many mailboxes**
- Assigned **time ranges** within which they are valid¹
- Set as **default signature** for new emails, or for replies and forwards (signatures only)
- Set as **default OOF message** for internal or external recipients (OOF messages only)
- Set in **Outlook Web¹** for the currently logged-in user, including mirroring signatures to the cloud as **roaming signatures¹** (Linux/macOS/Windows, Classic and New Outlook¹)
- Centrally managed only¹, or **exist along user-created signatures** (signatures only)
- Automatically added to new emails, reply emails and appointments with the **Outlook add-in**¹
- Copied to an **additional path¹** for easy access to signatures on mobile devices or for use with email clients and apps besides Outlook: Apple Mail, Google Gmail, Samsung Mail, Mozilla Thunderbird, GNOME Evolution, KDE KMail, and others.
- Create an **email draft containing all available signatures** in HTML and plain text for easy access in mail clients that do not have a signatures API
- **Write protected** (Outlook for Windows signatures only)

Set-OutlookSignatures can be **run by users on Windows, Linux and macOS clients, including shared devices and terminal servers - or on a central system with a service account¹**.
On clients, it can run as part of the logon script, as scheduled task, or on user demand via a desktop icon, start menu entry, shortcut or any other way of starting a program - **whatever your operating system and software deployment mechanism allows**.
Signatures and OOF messages can also be created and pushed into mailboxes centrally, **without end user or client involvement¹**.

**Sample templates** for signatures and OOF messages demonstrate many features and are provided as .docx and .htm files.

**Simulation mode** allows content creators and admins to simulate the behavior of the software for a specific user at a specific point in time, and to inspect the resulting signature files before going live.

**SimulateAndDeploy¹** allows to deploy signatures to Outlook Web¹/New Outlook¹ without any client deployment or end user interaction, making it ideal for users that only log on to web services but never to a client (users with a Microsoft 365 F-license, for example).

The software is **designed to work in big and complex environments** (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). It works **on premises, in hybrid and in cloud-only environments**.
All **national clouds are supported**: Public (AzurePublic), US Government L4 (AzureUSGovernment), US Government L5 (AzureUSGovernment DoD), China (AzureChinaCloud operated by 21Vianet).

It is **multi-client capable** by using different template paths, configuration files and script parameters.

Set-OutlookSignatures requires **no installation on servers or clients**. You only need a standard SMB file share on a central system, and optionally Office on your clients.
There is also **no telemetry** or "calling home", emails are **not routed through a 3rd party data center or cloud service**, and there is **no need to change DNS records (MX, SPF) or mail flow**.

A **documented implementation approach**, based on real life experiences implementing the software in multi-client environments with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and email and client administrators.
The implementation approach is **suited for service providers as well as for clients**, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The software core is **Free and Open-Source Software (FOSS)**. It is published under a license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) and other popular licenses. Please see `.\LICENSE.txt` for copyright and license details.

Footnote 1 (¹): **Some features are exclusive to the Benefactor Circle add-on.**
ExplicIT Consulting's commercial Benefactor Circle add-on enhances Set-OutlookSignatures with additional features and commercial support, ensuring that the core of Set-OutlookSignatures can remain Free and Open-Source Software (FOSS) and continues to evolve. See <a href="./Benefactor%20Circle.md" target="_blank">'.\docs\Benefactor Circle'</a> for details.

.LINK
Github: https://github.com/Set-OutlookSignatures/Set-OutlookSignatures
Benefactor Circle add-on: https://explicitconsulting.at/Set-OutlookSignatures

.PARAMETER SignatureTemplatePath
Path to centrally managed signature templates.

Local and remote paths are supported.

Local paths can be absolute ('C:\Signature templates') or relative to the software path ('.\sample templates\Signatures DOCX').

SharePoint document libraries are supported (https only): 'https://server.domain/sites/SignatureSite/SignatureDocLib/SignatureFolder' or '\\server.domain@SSL\sites\SignatureSite\SignatureDocLib\SignatureFolder'

Parameters and SharePoint sharing hints ('/:u:/r', etc.) are removed: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' -> 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'

On Linux and macOS, only already existing mount points and SharePoint Online paths can be accessed. Set-OutlookSignatures cannot create mount points itself, and access to SharePoint on-prem paths is a Windows-only feature.

For access to SharePoint Online, the Entra ID app needs the Files.Read.All or Files.SelectedOperations.Selected permission, and you need to pass the 'GraphClientID' parameter to Set-OutlookSignatures.

Default value: '.\sample templates\Signatures DOCX' on Windows, '.\sample templates\Signatures HTML' on Linux and macOS

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '.\sample templates\Signatures DOCX'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '.\sample templates\Signatures DOCX'"

.PARAMETER SignatureIniFile
Path to ini file containing signature template tags.

The file must be UTF8 encoded.

See '.\sample templates\Signatures DOCX\_Signatures.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the software path ('.\sample templates\Signatures DOCX')

SharePoint document libraries are supported (https only): 'https://server.domain/sites/SignatureSite/SignatureDocLib/SignatureFolder' or '\\server.domain@SSL\sites\SignatureSite\SignatureDocLib\SignatureFolder'

Parameters and SharePoint sharing hints ('/:u:/r', etc.) are removed: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' -> 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'

On Linux and macOS, only already existing mount points and SharePoint Online paths can be accessed. Set-OutlookSignatures cannot create mount points itself, and access to SharePoint on-prem paths is a Windows-only feature.

For access to SharePoint Online, the Entra ID app needs the Files.Read.All or Files.SelectedOperations.Selected permission, and you need to pass the 'GraphClientID' parameter to Set-OutlookSignatures.

Default value: '.\sample templates\Signatures DOCX\_Signatures.ini' on Windows, '.\sample templates\Signatures HTML\_Signatures.ini' on Linux and macOS

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SignatureIniFile '.\templates\Signatures DOCX\_Signatures.ini'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SignatureIniFile '.\templates\Signatures DOCX\_Signatures.ini'"

.PARAMETER ReplacementVariableConfigFile
Path to a replacement variable config file.

The file must be UTF8 encoded.

Local and remote paths are supported.

Local paths can be absolute ('C:\Signature templates') or relative to the software path ('.\sample templates\Signatures DOCX').

SharePoint document libraries are supported (https only): 'https://server.domain/sites/SignatureSite/SignatureDocLib/SignatureFolder' or '\\server.domain@SSL\sites\SignatureSite\SignatureDocLib\SignatureFolder'

Parameters and SharePoint sharing hints ('/:u:/r', etc.) are removed: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' -> 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'

On Linux and macOS, only already existing mount points and SharePoint Online paths can be accessed. Set-OutlookSignatures cannot create mount points itself, and access to SharePoint on-prem paths is a Windows-only feature.

For access to SharePoint Online, the Entra ID app needs the Files.Read.All or Files.SelectedOperations.Selected permission, and you need to pass the 'GraphClientID' parameter to Set-OutlookSignatures.

Default value: '.\config\default replacement variables.ps1'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -ReplacementVariableConfigFile '.\config\default replacement variables.ps1'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -ReplacementVariableConfigFile '.\config\default replacement variables.ps1'"

.PARAMETER GraphClientID
ID of the Entra ID app to use for Graph authentication.

This parameter must be used when the parameter GraphConfigFile points to a SharePoint Online location.

Per default, GraphClientID is not overwritten by the configuration defined in GraphConfigFile, but you can change this in the Graph config file itself.

Default value: $null

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 GraphClientID '3dc5f201-6c36-4b94-98ca-c66156a686a8'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 GraphClientID '3dc5f201-6c36-4b94-98ca-c66156a686a8'"

.PARAMETER GraphConfigFile
Path to a Graph variable config file.

The file must be UTF8 encoded.

Local and remote paths are supported.

Local paths can be absolute ('C:\config\default graph config.ps1') or relative to the software path ('.\config\default graph config.ps1')

SharePoint document libraries are supported (https only): 'https://server.domain/SignatureSite/config/default graph config.ps1' or '\\server.domain@SSL\SignatureSite\config\default graph config.ps1'

Parameters and SharePoint sharing hints ('/:u:/r', etc.) are removed: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' -> 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'

When GraphConfigFile is hosted on SharePoint Online, it is highly recommended to set the `GraphClientID` parameter. Else, access to GraphConfigFile will fail on Linux and macOS, and fall back to WebDAV with a required Internet Explorer authentication cookie on Windows.

On Linux and macOS, only already existing mount points and SharePoint Online paths can be accessed. Set-OutlookSignatures cannot create mount points itself, and access to SharePoint on-prem paths is a Windows-only feature.

For access to SharePoint Online, the Entra ID app needs the Files.Read.All or Files.SelectedOperations.Selected permission, and you need to pass the 'GraphClientID' parameter to Set-OutlookSignatures.

Default value: '.\config\default graph config.ps1'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -GraphConfigFile '.\config\default graph config.ps1'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 GraphConfigFile '.\config\default graph config.ps1'"

.PARAMETER TrustsToCheckForGroups
List of domains to check for group membership.

If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered.

If a string starts with a minus or dash ('-domain-a.local'), the domain after the dash or minus is removed from the list (no wildcards allowed).

All domains belonging to the Active Directory forest of the currently logged-in user are always considered, but specific domains can be removed ('*', '-childA1.childA.user.forest').

When a cross-forest trust is detected by the '*' option, all domains belonging to the trusted forest are considered but specific domains can be removed ('*', '-childX.trusted.forest').

On Linux and macOS, this parameter is ignored because on-prem Active Directories are not supported (only Graph is supported).

Default value: '*'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -TrustsToCheckForGroups 'corp.example.com', 'corp.example.net'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -TrustsToCheckForGroups 'corp.example.com', 'corp.example.net'"

.PARAMETER IncludeMailboxForestDomainLocalGroups
Shall the software consider group membership in domain local groups in the mailbox's AD forest?

Per default, membership in domain local groups in the mailbox's forest is not considered as the required LDAP queries are slow and domain local groups are usually not used in Exchange.

Domain local groups across trusts behave differently, they are always considered as soon as the trusted domain/forest is included in TrustsToCheckForGroups.

On Linux and macOS, this parameter is ignored because on-prem Active Directories are not supported (only Graph is supported).

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -IncludeMailboxForestDomainLocalGroups $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -IncludeMailboxForestDomainLocalGroups false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -IncludeMailboxForestDomainLocalGroups $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -IncludeMailboxForestDomainLocalGroups false"

.PARAMETER DeleteUserCreatedSignatures
Shall the software delete signatures which were created by the user itself?

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DeleteUserCreatedSignatures $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DeleteUserCreatedSignatures false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DeleteUserCreatedSignatures $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DeleteUserCreatedSignatures false"

.PARAMETER DeleteScriptCreatedSignaturesWithoutTemplate
Shall the software delete signatures which were created by the software before but are no longer available as template?

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DeleteScriptCreatedSignaturesWithoutTemplate $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DeleteScriptCreatedSignaturesWithoutTemplate false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DeleteScriptCreatedSignaturesWithoutTemplate $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DeleteScriptCreatedSignaturesWithoutTemplate false"

.PARAMETER SetCurrentUserOutlookWebSignature
Shall the software set the Outlook Web signature of the currently logged-in user?

If the parameter is set to '$true' and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. If no Outlook mailboxes are configured at all, additional mailbox configured in Outlook Web are used. This way, the software can be used in environments where only Outlook Web is used.

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SetCurrentUserOutlookWebSignature $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SetCurrentUserOutlookWebSignature true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SetCurrentUserOutlookWebSignature $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SetCurrentUserOutlookWebSignature true"

.PARAMETER SetCurrentUserOOFMessage
Shall the software set the out-of-office (OOF) message of the currently logged-in user?

If the parameter is set to '$true' and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. If no Outlook mailboxes are configured at all, additional mailbox configured in Outlook Web are used. This way, the software can be used in environments where only Outlook Web is used.

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SetCurrentUserOOFMessage $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SetCurrentUserOOFMessage true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SetCurrentUserOOFMessage $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SetCurrentUserOOFMessage true"

.PARAMETER OOFTemplatePath
Path to centrally managed signature templates.

Local and remote paths are supported.

Local paths can be absolute ('C:\OOF templates') or relative to the software path ('.\sample templates\ Out-of-office ').

SharePoint document libraries are supported (https only): 'https://server.domain/SignatureSite/OOFTemplates' or '\\server.domain@SSL\SignatureSite\OOFTemplates'

Parameters and SharePoint sharing hints ('/:u:/r', etc.) are removed: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' -> 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'

On Linux and macOS, only already existing mount points and SharePoint Online paths can be accessed. Set-OutlookSignatures cannot create mount points itself, and access to SharePoint on-prem paths is a Windows-only feature.

For access to SharePoint Online, the Entra ID app needs the Files.Read.All or Files.SelectedOperations.Selected permission, and you need to pass the 'GraphClientID' parameter to Set-OutlookSignatures.

Default value: '.\sample templates\Out-of-office DOCX' on Windows, '.\sample templates\Out-of-office HTML' on Linux and macOS

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -OOFTemplatePath '.\templates\Out-of-office DOCX'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -OOFTemplatePath '.\templates\Out-of-office DOCX'"

.PARAMETER OOFIniFile
Path to ini file containing signature template tags.

The file must be UTF8 encoded.

See '.\sample templates\Out-of-office DOCX\_OOF.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the software path ('.\sample templates\Signatures')

SharePoint document libraries are supported (https only): 'https://server.domain/sites/SignatureSite/SignatureDocLib/SignatureFolder' or '\\server.domain@SSL\sites\SignatureSite\SignatureDocLib\SignatureFolder'

Parameters and SharePoint sharing hints ('/:u:/r', etc.) are removed: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' -> 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'

On Linux and macOS, only already existing mount points and SharePoint Online paths can be accessed. Set-OutlookSignatures cannot create mount points itself, and access to SharePoint on-prem paths is a Windows-only feature.

For access to SharePoint Online, the Entra ID app needs the Files.Read.All or Files.SelectedOperations.Selected permission, and you need to pass the 'GraphClientID' parameter to Set-OutlookSignatures.

Default value: '.\sample templates\Out-of-office DOCX\_OOF.ini' on Windows, '.\sample templates\Out-of-office HTML\_OOF.ini' on Linux and macOS

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -OOFIniFile '.\templates\Out-of-office DOCX\_OOF.ini'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -OOFIniFile '.\templates\Out-of-office DOCX\_OOF.ini'"

.PARAMETER AdditionalSignaturePath
An additional path that the signatures shall be copied to.

Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.

This way, the user can easily copy-paste the preferred preconfigured signature for use in an email app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.

Local and remote paths are supported.

Local paths can be absolute ('C:\Outlook signatures') or relative to the software path ('.\Outlook signatures').

SharePoint document libraries are supported (https only, no SharePoint Online): 'https://server.domain/User' or '\\server.domain@SSL\User'

Parameters and SharePoint sharing hints ('/:u:/r', etc.) are removed: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' -> 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'

On Linux and macOS, only already existing mount points and SharePoint Online paths can be accessed. Set-OutlookSignatures cannot create mount points itself, and access to SharePoint on-prem paths is a Windows-only feature.

For access to SharePoint Online, the Entra ID app needs the Files.Read.All or Files.SelectedOperations.Selected permission, and you need to pass the 'GraphClientID' parameter to Set-OutlookSignatures.

If the folder or folder structure does not exist, it is created.

Also see related parameter 'EmbedImagesInHtmlAdditionalSignaturePath'.

This feature requires a Benefactor Circle license (when used outside of simulation mode).

Default value: "$(try { $([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Outlook Signatures')) } catch {})"

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -AdditionalSignaturePath "$(try { $([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Outlook Signatures')) } catch {})"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -AdditionalSignaturePath ""$(try { $([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Outlook Signatures')) } catch {})

.PARAMETER UseHtmTemplates
With this parameter, the software searches for templates with the extension .htm instead of .docx.

Each format has advantages and disadvantages, please see "Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates." for a quick overview.

Templates in .htm format must be UTF8 encoded.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false on Windows, $true on Linux and macOS

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -UseHtmTemplates $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -UseHtmTemplates false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -UseHtmTemplates $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -UseHtmTemplates false"


.PARAMETER SimulateUser
SimulateUser is a mandatory parameter for simulation mode. This value replaces the currently logged-in user.

Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an email-address, but is not necessarily one).

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SimulateUser "EXAMPLEDOMAIN\UserA"
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SimulateUser "user.a@example.com"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SimulateUser ""EXAMPLEDOMAIN\UserA"""
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SimulateUser ""user.a@example.com"""

.PARAMETER SimulateMailboxes
SimulateMailboxes is optional for simulation mode, although highly recommended.

It is a comma separated list of email addresses replacing the list of mailboxes otherwise gathered from the simulated user's Outlook Web.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $null

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SimulateMailboxes 'user.b@example.com', 'user.b@example.net'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SimulateMailboxes 'user.a@example.com', 'user.b@example.net'"

.PARAMETER SimulateTime
Use a certain timestamp for simulation mode. This allows you to simulate time-based templates.

Format: yyyyMMddHHmm (yyyy = year, MM = two-digit month, dd = two-digit day, HH = two-digit hour (0..24), mm = two-digit minute), local time

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $null

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SimulateTime "202312311859"
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SimulateUser "202312311859"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SimulateUser ""202312311859"""
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SimulateUser ""202312311859"""

.PARAMETER SimulateAndDeploy
Not only simulate, but deploy signatures while simulating

Makes only sense in combination with '.\sample code\SimulateAndDeploy.ps1', do not use this parameter for other scenarios

See '.\sample code\SimulateAndDeploy.ps1' for an example how to use this parameter

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false

.PARAMETER SimulateAndDeployGraphCredentialFile
Path to file containing Graph credential which should be used as alternative to other token acquisition methods

Makes only sense in combination with '.\sample code\SimulateAndDeploy.ps1', do not use this parameter for other scenarios

See '.\sample code\SimulateAndDeploy.ps1' for an example how to create and use this file

Default value: $null

.PARAMETER GraphOnly
Try to connect to Microsoft Graph only, ignoring any local Active Directory. On Linux and macOS, only Graph is supported.

The default behavior is to try Active Directory first and fall back to Graph.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false on Windows, $true on Linux and macOS

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -GraphOnly $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -GraphOnly false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -GraphOnly $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -GraphOnly false"

.PARAMETER CloudEnvironment
The cloud environment to connect to.

Allowed values:
- 'Public' (or: 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC')
- 'AzureUSGovernment' (or: 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4')
- 'AzureUSGovernmentDOD' (or: 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5')
- 'China' (or: 'AzureChina', 'ChinaCloud', 'AzureChinaCloud')

Default value: 'Public'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -CloudEnvironment "Public"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -CloudEnvironment ""Public"""

.PARAMETER CreateRtfSignatures
Should signatures be created in RTF format?

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -CreateRtfSignatures $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -CreateRtfSignatures false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -CreateRtfSignatures $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -CreateRtfSignatures false"

.PARAMETER CreateTxtSignatures
Should signatures be created in TXT format?

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -CreateTxtSignatures $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -CreateTxtSignatures true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -CreateTxtSignatures $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -CreateTxtSignatures true"

.PARAMETER MoveCSSInline
Move CSS to inline style attributes, for maximum email client compatibility.

This parameter is enabled per default, as a workaround to Microsoft's problem with formatting in Outlook Web (M365 roaming signatures and font sizes, especially).

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -MoveCSSInline $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -MoveCSSInline true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -MoveCSSInline $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -MoveCSSInline true"

.PARAMETER EmbedImagesInHtml
Should images be embedded into HTML files?

Outlook 2016 and newer can handle images embedded directly into an HTML file as BASE64 string ('<img src="data:image/[…]"').

Outlook 2013 and earlier can't handle these embedded images when composing HTML emails (there is no problem receiving such emails, or when composing RTF or TXT emails).

When setting EmbedImagesInHtml to $false, consider setting the Outlook registry value "Send Pictures With Document" to 1 to ensure that images are sent to the recipient (see https://support.microsoft.com/en-us/topic/inline-images-may-display-as-a-red-x-in-outlook-704ae8b5-b9b6-d784-2bdf-ffd96050dfd6 for details).

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -EmbedImagesInHtml $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -EmbedImagesInHtml false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -EmbedImagesInHtml $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -EmbedImagesInHtml false"

.PARAMETER EmbedImagesInHtmlAdditionalSignaturePath
Some feature as 'EmbedImagesInHtml' parameter, but only valid for the path defined in AdditionalSignaturesPath when not in simulation mode.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -EmbedImagesInHtmlAdditionalSignaturePath $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -EmbedImagesInHtmlAdditionalSignaturePath true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -EmbedImagesInHtmlAdditionalSignaturePath $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -EmbedImagesInHtmlAdditionalSignaturePath true"

.PARAMETER DocxHighResImageConversion
Enables or disables high resolution images in HTML signatures.

When enabled, this parameter uses a workaround to overcome a Word limitation that results in low resolution images when converting to HTML. The price for high resolution images in HTML signatures are more time needed for document conversion and signature files requiring more storage space.

Disabling this feature speeds up DOCX to HTML conversion, and HTML signatures require less storage space - at the cost of lower resolution images.

Contrary to conversion to HTML, conversion to RTF always results in high resolution images.

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DocxHighResImageConversion $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DocxHighResImageConversion true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DocxHighResImageConversion $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DocxHighResImageConversion true"

.PARAMETER SignaturesForAutomappedAndAdditionalMailboxes
Deploy signatures for automapped mailboxes and additional mailboxes

Signatures can be deployed for these mailboxes, but not set as default signature due to technical restrictions in Outlook

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SignaturesForAutomappedAndAdditionalMailboxes $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SignaturesForAutomappedAndAdditionalMailboxes true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SignaturesForAutomappedAndAdditionalMailboxes $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SignaturesForAutomappedAndAdditionalMailboxes true"

.PARAMETER DisableRoamingSignatures
Disable signature roaming in Outlook. Onyl works on Windows. Has no effect on signature mirroring via the MirrorCloudSignatures parameter.

A value representing true disables roaming signatures, a value representing false enables roaming signatures, any other value leaves the setting as-is.

Only sets HKCU registry key, does not override configuration set by group policy.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no', $null, ''

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DisableRoamingSignatures $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DisableRoamingSignatures true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DisableRoamingSignatures $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DisableRoamingSignatures true"

.PARAMETER MirrorCloudSignatures
Should local signatures be mirrored with signatures in Exchange Online?

Possible for Exchange Online mailboxes:
- Download for every mailbox where the current user has full access
- Upload and set default signatures for the mailbox of the current user

Prerequisites:
- Download
  - Current user has full access to the mailbox
- Upload, set default signatures
  - Script parameter `SetCurrentUserOutlookWebSignature` is set to `true`
  - Mailbox is the mailbox of the currently logged-in user and is hosted in Exchange Online

Please note:
- As there is no Microsoft official API yet, this feature is to be used at your own risk.
- This feature does not work in simulation mode, because the user running the simulation does not have access to the signatures stored in another mailbox

The process is very simple and straight forward. Set-OutlookSignatures goes through the following steps for each mailbox:
1. Check if all prerequisites are met
2. Download all signatures stored in the Exchange Online mailbox
  - This mimics Outlook's behavior: Roaming signatures are only manipulated in the cloud and then downloaded from there.
  -   - An existing local signature is only overwritten when the cloud signature is newer and when it has not been processed before for a mailbox with higher priority
3. Go through standard template and signature processing
  - Loop through the templates and their configuration, and convert them to signatures
  - Set default signatures for replies and forwards
  - If configured, delete signatures created by the user
  - If configured, delete signatures created earlier by Set-OutlookSignatures but now no longer have a corresponding central configuration
4. Delete all signatures in the cloud and upload all locally stored signatures to the user's personal mailbox as roaming signatures

There may be too many signatures available in the cloud - in this case, having too many signatures is better than missing some.

Another advantage of this solution is that it makes roaming signatures available in Outlook versions that do not officially support them.

What will not work:
- Download from mailboxes where the current user does not have full access rights. This is not possible because of Microsoft API restrictions.
- Download from and upload to shared mailboxes. This is not possible because of Microsoft API restrictions.
- Uploading signatures other than device specific signatures and such assigned to the mailbox of the current user. Uploading is not implemented, because until now no way could be found that does not massively impact the user experience as soon as the Outlook integrated download process starts (signatures available multiple times, etc.)

Attention: When Outlook v16 and higher is allowed to sync signatures itself, it may overwrite signatures created by this software with their cloud versions. To avoid this, it is recommended to set the parameters DisableRoamingSignatures and MirrorCloudSignatures to true instead.

Consider combining MirrorCloudSignatures with MailboxSpecificSignatureNames.

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -MirrorCloudSignatures $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -MirrorCloudSignatures false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -MirrorCloudSignatures $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -MirrorCloudSignatures false"

.PARAMETER MailboxSpecificSignatureNames
Should signature names be mailbox specific by adding the email address?

For compatibility with Outlook storing signatures in the file system, Set-OutlookSignatures converts templates to signatures according to the following logic:
1. Get all mailboxes and sort them: Mailbox of logged-on/simulated user, other mailboxes in default Outlook profile or Outlook Web, mailboxes from other Outlook profiles
2. Get all template files, sort them by category (common, group specific, mailbox specific, replacement variable specific), and within each category by SortOrder and SortCulture defined in the INI file
3. Loop through the mailbox list, and for each mailbox loop through the template list.
4. If a template's conditions apply and if the template has not been used before, convert the template to a signature.

The step 4 condition `if the template has not been used before` makes sure that a lower priority mailbox does not replace a signature with the same name which has already been created for a higher priority mailbox.

With roaming signatures (signatures being stored in the Exchange Online mailbox itself) being used more and more, the step 4 condition `if the template has not been used before` makes less sense. By setting the `MailboxSpecificSignatureNames` parameter to `true`, this restriction no longer applies. To avoid naming collisions, the email address of the current mailbox is added to the name of the signature - instead of a single `Signature A` file, Set-OutlookSignatures can create a separate signature file for each mailbox: `Signature A (user.a@example.com)`, `Signature A (mailbox.b@example.net)`, etc.

This naming convention intentionally matches Outlook's convention for naming roaming signatures. Before setting `MailboxSpecificSignatureNames` to `true`, consider the impact on the `DisableRoamingSignatures` and `MirrorCloudSignatures` parameters - it is recommended to set both parameters to `true` to achieve the best user experience and to avoid problems with Outlook's own roaming signature synchronisation.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -MailboxSpecificSignatureNames $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -MailboxSpecificSignatureNames false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -MailboxSpecificSignatureNames $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -MailboxSpecificSignatureNames false"

.PARAMETER WordProcessPriority
Define the Word process priority. With lower values, Set-OutlookSignatures runs longer but minimizes possible performance impact

Allowed values (ascending priority): Idle, 64, BelowNormal, 16384, Normal, 32, AboveNormal, 32768, High, 128, RealTime, 256

Default value: 'Normal' ('32')

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -WordProcessPriority Normal
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -WordProcessPriority 32
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -WordProcessPriority Normal"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -WordProcessPriority 32"

.PARAMETER ScriptProcessPriority
Define the script process priority. With lower values, Set-OutlookSignatures runs longer but minimizes possible performance impact

Allowed values (ascending priority): Idle, 64, BelowNormal, 16384, Normal, 32, AboveNormal, 32768, High, 128, RealTime, 256

Default value: 'Normal' ('32')

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -ScriptProcessPriority Normal
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -ScriptProcessPriority 32
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -ScriptProcessPriority Normal"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -ScriptProcessPriority 32"

.PARAMETER SignatureCollectionInDrafts
When enabled, this creates and updates an email message with the subject 'My signatures, powered by Set-OutlookSignatures Benefactor Circle' in the drafts folder of the current user, containing all available signatures in HTML and plain text for easy access in mail clients that do not have a signatures API.

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SignatureCollectionInDrafts $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SignatureCollectionInDrafts false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SignatureCollectionInDrafts $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SignatureCollectionInDrafts false"

.PARAMETER BenefactorCircleID
The Benefactor Circle member ID matching your license file, which unlocks exclusive features.

Default value: ''

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -BenefactorCircleID 00000000-0000-0000-0000-000000000000
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -BenefactorCircleID 00000000-0000-0000-0000-000000000000"

.PARAMETER BenefactorCircleLicenseFile
The Benefactor Circle license file matching your member ID, which unlocks exclusive features.

Default value: ''

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -BenefactorCircleLicenseFile ".\license.dll"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -BenefactorCircleLicenseFile "".\license.dll"""

.PARAMETER VirtualMailboxConfigFile
Path a PowerShell file containing the logic to define virtual mailboxes. You can also use the VirtualMailboxConfigFile to dynamically define signature INI file entries.

Virtual mailboxes are mailboxes that are not available in Outlook but are treated by Set-OutlookSignatures as if they were.

This is an option for scenarios where you want to deploy signatures with not only the '`$CurrentUser...$`' but also '`$CurrentMailbox...$`' replacement variables for mailboxes that have not been added to Outlook, such as in Send As or Send On Behalf scenarios, where users often only change the from address but do not add the mailbox to Outlook.

See '`.\sample code\VirtualMailboxConfigFile.ps1`' for sample code showing the most relevant use cases.

For maximum automation, use VirtualMailboxConfigFile together with [Export-RecipientPermissions](https://github.com/Export-RecipientPermissions).

This feature requires a Benefactor Circle license.

Local and remote paths are supported. Local paths can be absolute ('C:\VirtualMailboxConfigFile.ps1') or relative to the software path ('.\sample code\VirtualMailboxConfigFile')

SharePoint document libraries are supported (https only): 'https://server.domain/SignatureSite/config/VirtualMailboxConfigFile.ps1' or '\\server.domain@SSL\SignatureSite\config\VirtualMailboxConfigFile.ps1'

Parameters and SharePoint sharing hints ('/:u:/r', etc.) are removed: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' -> 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'

On Linux and macOS, only already existing mount points and SharePoint Online paths can be accessed. Set-OutlookSignatures cannot create mount points itself, and access to SharePoint on-prem paths is a Windows-only feature.

For access to SharePoint Online, the Entra ID app needs the Files.Read.All or Files.SelectedOperations.Selected permission, and you need to pass the 'GraphClientID' parameter to Set-OutlookSignatures.

Default value: ''

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -VirtualMailboxConfigFile '.\sample code\VirtualMailboxConfigFile.ps1'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -VirtualMailboxConfigFile '.\sample code\VirtualMailboxConfigFile.ps1'"

.INPUTS
None. You cannot pipe objects to Set-OutlookSignatures.ps1.

.OUTPUTS
Set-OutlookSignatures.ps1 writes the current activities, warnings and error messages to the standard output stream.

.EXAMPLE
Run Set-OutlookSignatures with default values and sample templates
PS> .\Set-OutlookSignatures.ps1

.EXAMPLE
Use custom signature templates and custom ini file
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureIniFile '\\internal.example.com\share\Signature Templates\_Signatures.ini'

.EXAMPLE
Use custom signature templates, ignore trust to internal-test.example.com
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -TrustsToCheckForGroups '*', '-internal-test.example.com'

.EXAMPLE
Use custom signature templates, only check domains/trusts internal-test.example.com and company.b.com
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -TrustsToCheckForGroups 'internal-test.example.com', 'company.b.com'

.EXAMPLE
Passing arguments to PowerShell.exe from the command line or task scheduler can be very tricky when spaces are involved. See '.\docs\README' for details.
PowerShell.exe -Command "& '\\server\share\directory\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\server\share\directory\templates\Signatures DOCX' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -OOFTemplatePath '\\server\share\directory\templates\Out-of-office DOCX' -OOFTemplatePath '\\internal.example.com\share\Signature Templates\_OOF.ini' -ReplacementVariableConfigFile '\\server\share\directory\config\default replacement variables.ps1' "

.EXAMPLE
Please see '.\docs\README' and https://github.com/Set-OutlookSignatures/Set-OutlookSignatures for more details.

.NOTES
Software: Set-OutlookSignatures
Version : XXXVersionStringXXX
Web     : https://github.com/Set-OutlookSignatures/Set-OutlookSignatures
License : See '.\LICENSE.txt' for details and copyright
#>


# Suppress specific PSScriptAnalyzerRules for specific variables
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'SimulateAndDeployGraphCredentialFile')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentAutodiscoverSecureName')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentAzureADEndpoint')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentEnvironmentName')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentGraphApiEndpoint')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentSharePointOnlineDomains')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'ConnectedFilesFolderNames')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CurrentTemplateisForAliasSmtp')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'data')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'HTMLMarkerTag')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'OOFExternalValueBasename')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'OOFFilesExternal')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'OOFFilesInternal')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'OOFInternalValueBasename')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'pathHtmlFolderSuffix')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'PrimaryMailboxAddress')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'ScriptInvocation')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'ScriptVersion')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'SignatureFilesDefaultNew')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'SignatureFilesDefaultReplyFwd')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'SignatureFilesWriteProtect')]


[CmdletBinding(PositionalBinding = $false, DefaultParameterSetName = 'Z: All parameters')]


Param(
    # Path to a Benefactor Circle license file
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$BenefactorCircleLicenseFile = '',

    # The Benefactor Circle Member ID matching the Benefactor Circle license file
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$BenefactorCircleID = '',

    # Use templates in .HTM file format instead of .DOCX
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $UseHtmTemplates = $(if ($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) { $false } else { $true }),

    # Path to centrally managed signature templates
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$SignatureTemplatePath = $(if (($UseHtmTemplates -inotin @(1, 'true', '$true', 'yes')) -or (-not $UseHtmTemplates)) { '.\sample templates\Signatures DOCX' } else { '.\sample templates\Signatures HTML' }),

    # Path to ini file containing signature template tags
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [Alias('SignatureIniPath')]
    [string]$SignatureIniFile = $(if (($UseHtmTemplates -inotin @(1, 'true', '$true', 'yes')) -or (-not $UseHtmTemplates)) { '.\sample templates\Signatures DOCX\_Signatures.ini' } else { '.\sample templates\Signatures HTML\_Signatures.ini' }),

    # Deploy signatures for automapped mailboxes and additional mailboxes
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SignaturesForAutomappedAndAdditionalMailboxes = $true,

    # Shall the software delete signatures which were created by the user itself?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $DeleteUserCreatedSignatures = $false,

    # Shall the software delete signatures which were created by the software before but are no longer available as template?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $DeleteScriptCreatedSignaturesWithoutTemplate = $true,

    # Shall the software set the Outlook Web signature of the currently logged-in user?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SetCurrentUserOutlookWebSignature = $true,

    # An additional path that the signatures shall be copied to
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'F: Simulation mode')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [string]$AdditionalSignaturePath = $(try { $([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Outlook Signatures')) } catch {}),

    # Should HTML signatures contain high resolution images?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $DocxHighResImageConversion = $true,

    # Create RTF signatures
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $CreateRtfSignatures = $false,

    # Create TXT signatures
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $CreateTxtSignatures = $true,

    # Move CSS to inline style attributes
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $MoveCSSInline = $true,

    # Embed images in HTML
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $EmbedImagesInHtml = $false,

    # Embed images in HTML for AdditionalSignaturePath
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $EmbedImagesInHtmlAdditionalSignaturePath = $true,

    # Should signature names be mailbox specific by adding the email address?
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    $MailboxSpecificSignatureNames = $false,

    # Shall the software set the out-of-office (OOF) message(s) of the currently logged-in user?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SetCurrentUserOOFMessage = $true,

    # Path to centrally managed out-of-office (OOF, automatic reply) templates
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$OOFTemplatePath = $(if (($UseHtmTemplates -inotin @(1, 'true', '$true', 'yes')) -or (-not $UseHtmTemplates)) { '.\sample templates\Out-of-office DOCX' } else { '.\sample templates\Out-of-office HTML' }),

    # Path to ini file containing OOF template tags
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [Alias('OOFIniPath')]
    [string]$OOFIniFile = $(if (($UseHtmTemplates -inotin @(1, 'true', '$true', 'yes')) -or (-not $UseHtmTemplates)) { '.\sample templates\Out-of-office DOCX\_OOF.ini' } else { '.\sample templates\Out-of-office HTML\_OOF.ini' }),

    # Path to a replacement variable config file.
    [Parameter(Mandatory = $false, ParameterSetName = 'D: Replacement variables')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$ReplacementVariableConfigFile = '.\config\default replacement variables.ps1',

    # Path to a virtual mailbox config file.
    [Parameter(Mandatory = $false, ParameterSetName = 'D: Replacement variables')]
    [Parameter(Mandatory = $false, ParameterSetName = 'G: Outlook')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$VirtualMailboxConfigFile = '',

    # Try to connect to Microsoft Graph only, ignoring any local Active Directory.
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $GraphOnly = $(if ($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) { $false } else { $true }),

    # GraphClientID, later overwritten by $GraphConfigFile
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    $GraphClientID = $null,

    # Cloud environment to use
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet('Public', 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC', 'AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4', 'AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5', 'China', 'AzureChina', 'ChinaCloud', 'AzureChinaCloud')]
    [string]$CloudEnvironment = 'Public',

    # Path to a Graph variable config file.
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$GraphConfigFile = '.\config\default graph config.ps1',

    # List of domains/forests to check for group membership across trusts
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [Alias('DomainsToCheckForGroups')]
    [string[]]$TrustsToCheckForGroups = @('*'),

    # Shall the software consider group membership in domain local groups in the mailbox's AD forest?
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $IncludeMailboxForestDomainLocalGroups = $false,

    # Deploy while simulating
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SimulateAndDeploy = $false,

    # Path to file containing Graph credential which should be used as alternative to other token acquisition methods
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$SimulateAndDeployGraphCredentialFile = '',

    # Simulate another user as currently logged-in user
    [Parameter(Mandatory = $false, ParameterSetName = 'F: Simulation mode')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [Alias('SimulationUser', 'WhatIf')]
    [validatescript({
            $tempSimulateUser = $_
            if ($tempSimulateUser -imatch '^\S+@\S+$|^\S+\\\S+$') {
                $true
            } else {
                throw "'$tempSimulateUser' does not match the required format 'User@Domain' (UPN) or 'Domain\User'."
            }
        }
    )]
    [string]$SimulateUser = $null,

    # Simulate list of mailboxes instead of mailboxes configured in Outlook
    [Parameter(Mandatory = $false, ParameterSetName = 'F: Simulation mode')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [Alias('SimulationMailboxes')]
    [mailaddress[]]$SimulateMailboxes = $null,

    # Use a specific time for simulation mode
    [Parameter(Mandatory = $false, ParameterSetName = 'F: Simulation mode')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [Alias('SimulationTime')]
    [validatescript({
            $tempSimulateTime = $_
            if ($tempSimulateTime -imatch '\d{12}') {
                [DateTime]::ParseExact($tempSimulateTime, 'yyyyMMddHHmm', $null)
                $true
            } else {
                throw "'$tempSimulateTime' does not match the required format 'yyyyMMddHHmm'."
            }
        }
    )]
    [string]$SimulateTime = $null,

    # Should roaming signatures be disabled in Outlook?
    [Parameter(Mandatory = $false, ParameterSetName = 'G: Outlook')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    $DisableRoamingSignatures = $true,

    # Should local signatures be uploaded as roaming signature for the current user?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'G: Outlook')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    [Alias('MirrorLocalSignaturesToCloud')]
    $MirrorCloudSignatures = $true,

    # Word process priority
    [Parameter(Mandatory = $false, ParameterSetName = 'H: Word')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet('Idle', 64, 'BelowNormal', 16384, 'Normal', 32, 'AboveNormal', 32768, 'High', 128, 'RealTime', 256)]
    $WordProcessPriority = 'Normal',

    # Script process priority
    [Parameter(Mandatory = $false, ParameterSetName = 'I: Script')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet('', 'Idle', 64, 'BelowNormal', 16384, 'Normal', 32, 'AboveNormal', 32768, 'High', 128, 'RealTime', 256)]
    $ScriptProcessPriority = '',

    # Should the 'SignatureCollectionInDrafts' email draft be created and updated?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'G: Outlook')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SignatureCollectionInDrafts = $true
)


function ToSemVer($version) {
    $major = 0
    $minor = 0
    $patch = 0
    $pre = @()

    if (($version -ireplace '^v') -match '^(?<major>\d+)(\.(?<minor>\d+))?(\.(?<patch>\d+))?(\-(?<pre>[0-9A-Za-z\-\.]+))?(\+(?<build>[0-9A-Za-z\-\.]+))?$') {
        $major = [int]$matches['major']
        $minor = [int]$matches['minor']
        $patch = [int]$matches['patch']

        if ($null -eq $matches['pre']) {
            $pre = @()
        } else {
            $pre = $matches['pre'].Split('.')
        }
    }

    New-Object PSObject -Property @{
        Major         = $major
        Minor         = $minor
        Patch         = $patch
        Pre           = $pre
        VersionString = $version
    } | Select-Object -Property Major, Minor, Patch, Pre, VersionString
}


function CompareSemVer($a, $b) {
    $result = 0
    $result = $a.Major.CompareTo($b.Major)
    if ($result -ne 0) { return $result }

    $result = $a.Minor.CompareTo($b.Minor)
    if ($result -ne 0) { return $result }

    $result = $a.Patch.CompareTo($b.Patch)
    if ($result -ne 0) { return $result }

    $ap = $a.Pre
    $bp = $b.Pre

    if ($ap.Length -eq 0 -and $bp.Length -eq 0) { return 0 }
    if ($ap.Length -eq 0) { return 1 }
    if ($bp.Length -eq 0) { return -1 }

    $minLength = [Math]::Min($ap.Length, $bp.Length)

    for ($i = 0; $i -lt $minLength; $i++) {
        $ac = $ap[$i]
        $bc = $bp[$i]

        $anum = 0
        $bnum = 0
        $aIsNum = [Int]::TryParse($ac, [ref] $anum)
        $bIsNum = [Int]::TryParse($bc, [ref] $bnum)

        if ($aIsNum -and $bIsNum) {
            $result = $anum.CompareTo($bnum)
            if ($result -ne 0) {
                return $result
            }
        }
        if ($aIsNum) {
            return -1
        }
        if ($bIsNum) {
            return 1
        }

        $result = [string]::CompareOrdinal($ac, $bc)

        if ($result -ne 0) { return $result }
    }

    return $ap.Length.CompareTo($bp.Length)
}


function rankedSemVer($versions) {
    for ($i = 0; $i -lt $versions.Length; $i++) {
        $rank = 0

        for ($j = 0; $j -lt $versions.Length; $j++) {
            $diff = 0
            $diff = compareSemVer $versions[$i] $versions[$j]

            if ($diff -gt 0) {
                $rank++
            }
        }

        $current = [PsObject]$versions[$i]
        Add-Member -InputObject $current -MemberType NoteProperty -Name Rank -Value $rank -Force
    }

    return $versions
}


function CheckFilenamePossiblyInvalid ([string] $Filename = '', [bool] $CheckOutlook = $true, [bool] $CheckDeviceNames = $false) {
    $InvalidCharacters = @()

    # [System.Io.Path]::GetInvalidFileNameChars()
    $InvalidCharacters += @(($Filename | Select-String -Pattern "[$([regex]::escape(([System.Io.Path]::GetInvalidFileNameChars() -join '')))]" -AllMatches).Matches.Value) | Where-Object { $_ }

    # Outlook GUI
    if ($CheckOutlook) {
        $InvalidCharacters += @(($Filename | Select-String -Pattern "[$([regex]::escape('\/:"*?><,|'))]" -AllMatches).Matches.Value) | Where-Object { $_ }
    }

    # Windows reserved file names and device names (CON, PRN, AUX, COMx, LPTx, …)
    if ($CheckDeviceNames) {
        if (([System.Io.Path]::GetFullPath($Filename)).StartsWith('\\.\')) {
            $InvalidCharacters += $Filename
        }
    }

    $InvalidCharacters = @(@($InvalidCharacters | Select-Object -Unique | Where-Object { $_ } | Sort-Object -Culture $TemplateFilesSortCulture) | ForEach-Object { "'$($_)'" })

    if ($InvalidCharacters) {
        return $InvalidCharacters -join ', '
    }
}


### ▼▼▼ BlockSleep initiation code below ▼▼▼
##
#
# Place this code in your main script, as early in the code as possible
#
# Call BlockSleep wherever you want the current process to block sleep
#   BlockSleep
#
# On Windows, you can set three parameters:
#   -RequireAwaymode: Allows Away mode (defaults to true when not set)
#   -RequireDisplay: Requires the display to be on (defaults to false when not set)
#   -RequireSystem: Requires the system to be on (default to true when not set)
# On Linux, systemd-inhibit is required (should be available on most distributions)
# On macOS, caffeinate is required (should be available built-in)
#
# To allow sleep again, call BlockSleep with the AllowSleep parameter:
#   BlockSleep -AllowSleep
#
function BlockSleep {
    param (
        [switch]$AllowSleep,
        [switch]$RequireAwayMode,
        [switch]$RequireDisplay,
        [switch]$RequireSystem
    )

    if ($AllowSleep) {
        $RequireAwayMode = $false
        $RequireDisplay = $false
        $RequireSystem = $false
    } else {
        if (-not $PSBoundParameters.ContainsKey('RequireAwayMode')) {
            $RequireAwayMode = $true
        }

        if (-not $PSBoundParameters.ContainsKey('RequireDisplay')) {
            $RequireDisplay = $false
        }

        if (-not $PSBoundParameters.ContainsKey('RequireSystem')) {
            $RequireSystem = $true
        }

        if (
            ($RequireAwayMode -eq $false) -and
            ($RequireDisplay -eq $false) -and
            ($RequireSystem -eq $false)
        ) {
            $AllowSleep = $true
        }
    }

    if ($isWindows -or (-not (Test-Path 'variable:IsWindows'))) {
        $code = @'
[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern void SetThreadExecutionState(uint esFlags);
'@

        $ste = Add-Type -MemberDefinition $code -Name System -Namespace Win32 -PassThru
        $ES_CONTINUOUS = [uint32]'0x80000000'
        $ES_AWAYMODE_REQUIRED = [uint32]'0x00000040'
        $ES_DISPLAY_REQUIRED = [uint32]'0x00000002'
        $ES_SYSTEM_REQUIRED = [uint32]'0x00000001'

        $flags = $ES_CONTINUOUS

        if (-not $AllowSleep) {
            if ($RequireAwayMode) { $flags = $flags -bor $ES_AWAYMODE_REQUIRED }
            if ($RequireDisplay) { $flags = $flags -bor $ES_DISPLAY_REQUIRED }
            if ($RequireSystem) { $flags = $flags -bor $ES_SYSTEM_REQUIRED }
        }

        $ste::SetThreadExecutionState($flags)
    } elseif ($isLinux) {
        if (Get-Command systemd-inhibit -ErrorAction SilentlyContinue) {
            if ($script:BlockSleepInhibitPID) {
                Stop-Process -Id $script:BlockSleepInhibitPID -Force
                Remove-Variable -Name BlockSleepInhibitPID -Scope script
            }

            if (-not $AllowSleep) {
                $script:BlockSleepInhibitPID = Start-Process systemd-inhibit -ArgumentList "--what=idle --why=""Set-OutlookSignatures"" --who=""Set-OutlookSignatures"" tail --pid=$($PID) --follow /dev/null" -PassThru | Select-Object -ExpandProperty Id
            }
        } else {
            Write-Host "  'systemd-inhibit' is not available."
        }
    } elseif ($isMacOS) {
        if (Get-Command caffeinate -ErrorAction SilentlyContinue) {
            if ($script:BlockSleepInhibitPID) {
                Stop-Process -Id $script:BlockSleepInhibitPID -Force
                Remove-Variable -Name BlockSleepInhibitPID -Scope script
            }

            if (-not $AllowSleep) {
                $script:BlockSleepInhibitPID = Start-Process caffeinate -ArgumentList "-ims -w $($PID)" -PassThru | Select-Object -ExpandProperty Id
            }
        } else {
            Write-Host "  'caffeinate' is not available."
        }
    }
}
#
##
### ▲▲▲ BlockSleep initiation code above ▲▲▲


function main {
    $ScriptVersion = 'XXXVersionStringXXX'

    try { WatchCatchableExitSignal } catch { }

    # Init default values
    if ($null -ne [SetOutlookSignatures.Common].GetMethod('Init')) {
        [SetOutlookSignatures.Common]::Init()

        if (-not $SetOutlookSignaturesCommonInitDone) {
            $script:ExitCode = 5
            $script:ExitCodeDescription = 'Common initialization routine failed.'
            exit
        }
    } else {
        Write-Host 'Error initializing Set-OutlookSignatures. Exiting.' -ForegroundColor Red
        $script:ExitCode = 6
        $script:ExitCodeDescription = 'Common initialization routine not available.'
        exit
    }

    try { WatchCatchableExitSignal } catch { }

    # Import AngleSharp.CSS
    $script:AngleSharpCssNetModulePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))

    if ($($PSVersionTable.PSEdition) -ieq 'Core') {
        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\AngleSharp.Css\netstandard2.0')) -Destination $script:AngleSharpCssNetModulePath -Recurse
        if (-not $IsLinux) { Get-ChildItem $script:AngleSharpCssNetModulePath -Recurse | Unblock-File }
        Import-Module (Join-Path -Path $script:AngleSharpCssNetModulePath -ChildPath 'AngleSharp.Css.dll')
        Import-Module (Join-Path -Path $script:AngleSharpCssNetModulePath -ChildPath 'AngleSharp.dll')
    } else {
        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\PreMailer.Net\net462')) -Destination $script:AngleSharpCssNetModulePath -Recurse
        if (-not $IsLinux) { Get-ChildItem $script:AngleSharpCssNetModulePath -Recurse | Unblock-File }
        Import-Module (Join-Path -Path $script:AngleSharpCssNetModulePath -ChildPath 'AngleSharp.dll')
    }

    try { WatchCatchableExitSignal } catch { }

    # Import QRCoder
    $script:QRCoderModulePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))

    Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\QRCoder\netstandard2.0')) -Destination $script:QRCoderModulePath -Recurse
    if (-not $IsLinux) { Get-ChildItem $script:QRCoderModulePath -Recurse | Unblock-File }
    Import-Module (Join-Path -Path $script:QRCoderModulePath -ChildPath 'QRCoder.dll')


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Get basic Outlook and Word information @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $OutlookProfiles = @()
    $OutlookUseNewOutlook = $null

    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled, skip Outlook checks'
    } else {
        if ($IsWindows) {
            Write-Host '  Outlook'

            if ($(Get-Command -Name 'Get-AppPackage' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                $NewOutlook = Get-AppPackage -Name 'Microsoft.OutlookForWindows' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            } else {
                $NewOutlook = $null
            }

            $OutlookRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Outlook.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))

            if ($OutlookRegistryVersion -eq [System.Version]::Parse('0.0.0.0')) {
                $OutlookRegistryVersion = $null
            }

            try {
                # [Microsoft.Win32.RegistryView]::Registry32 makes sure view the registry as a 32 bit application would
                # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
                # Covers:
                #   Office x86 on Windows x86
                #   Office x86 on Windows x64
                #   Any PowerShell process bitness
                $OutlookFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry32)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '').trim('"').trim('''') -ErrorAction Stop
            } catch {
                try {
                    # [Microsoft.Win32.RegistryView]::Registry64 makes sure we view the registry as a 64 bit application would
                    # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
                    # Covers:
                    #   Office x64 on Windows x64
                    #   Any PowerShell process bitness
                    $OutlookFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry64)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '').trim('"').trim('''') -ErrorAction Stop
                } catch {
                    $OutlookFilePath = $null
                }
            }

            if ($OutlookFilePath) {
                try {
                    $OutlookBitnessInfo = GetBitness -fullname $OutlookFilePath
                    $OutlookFileVersion = [System.Version]::Parse((((($OutlookBitnessInfo.'File Version'.ToString() + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
                    $OutlookBitness = $OutlookBitnessInfo.Architecture
                    Remove-Variable -Name 'OutlookBitnessInfo'
                } catch {
                    $OutlookBitness = $null
                    $OutlookFilePath = $null
                    $OutlookFileVersion = $null
                }
            } else {
                $OutlookBitness = $null
                $OutlookFilePath = $null
                $OutlookFileVersion = $null
            }

            if ($OutlookRegistryVersion.Major -ne $OutlookFileVersion.Major) {
                Write-Host "    Major parts of Outlook version from registry ('$OutlookRegistryVersion') and from outlook.exe ('$OutlookFileVersion') do not match." -ForegroundColor Yellow
                Write-Host '    Assuming that Outlook is not installed.' -ForegroundColor Yellow
                Write-Host '    To resolve this, repair the Outlook installation and/or the registry information about Outlook.' -ForegroundColor Yellow

                $OutlookRegistryVersion = $null
                $OutlookFilePath = $null
                $OutlookFileVersion = $null
                $OutlookBitness = $null
            }

            if ($null -ne $OutlookRegistryVersion) {
                if ($OutlookRegistryVersion.major -gt 16) {
                    Write-Host "    Outlook version $OutlookRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
                    $script:ExitCode = 7
                    $script:ExitCodeDescription = 'Outlook version newer than 16 is not yet known.'
                    exit
                } elseif ($OutlookRegistryVersion.major -eq 16) {
                    $OutlookRegistryVersion = '16.0'
                } elseif ($OutlookRegistryVersion.major -eq 15) {
                    $OutlookRegistryVersion = '15.0'
                } elseif ($OutlookRegistryVersion.major -eq 14) {
                    $OutlookRegistryVersion = '14.0'
                } elseif ($OutlookRegistryVersion.major -lt 14) {
                    Write-Host "    Outlook version $OutlookRegistryVersion is older than Outlook 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
                    $script:ExitCode = 8
                    $script:ExitCodeDescription = 'Outlook version older than 2010 is not supported.'
                    exit
                }
            }

            if ($null -ne $OutlookRegistryVersion) {
                Write-Host "    Set 'Send Pictures With Document' registry value to '1'"
                $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Options\Mail" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'Send Pictures With Document' -Type DWORD -Value 1 -Force
            }

            if (($DisableRoamingSignatures -in @($true, $false)) -and $OutlookRegistryVersion -and ($OutlookFileVersion -ge '16.0.0.0')) {
                Write-Host "    Set 'DisableRoamingSignatures' registry value to '$([int]$DisableRoamingSignatures)'"
                $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableRoamingSignaturesTemporaryToggle' -Type DWORD -Value $([int]$DisableRoamingSignatures) -Force
                $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableRoamingSignatures' -Type DWORD -Value $([int]$DisableRoamingSignatures) -Force
            }

            if ($null -ne $OutlookRegistryVersion) {
                try {
                    $OutlookDefaultProfile = (Get-ItemProperty "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\Outlook" -ErrorAction Stop -WarningAction SilentlyContinue).DefaultProfile

                    $OutlookProfiles = @(@((Get-ChildItem "hkcu:\SOFTWARE\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles" -ErrorAction Stop -WarningAction SilentlyContinue).PSChildName) | Where-Object { $_ })

                    if ($OutlookDefaultProfile -and ($OutlookDefaultProfile -iin $OutlookProfiles)) {
                        $OutlookProfiles = @(@($OutlookDefaultProfile) + @($OutlookProfiles | Where-Object { $_ -ine $OutlookDefaultProfile }))
                    }
                } catch {
                    $OutlookDefaultProfile = $null
                    $OutlookProfiles = @()
                }

                $OutlookIsBetaversion = $false

                if (
                    ((Get-Item 'registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Property -contains 'UpdateChannel') -and
                    ($OutlookFileVersion -ge '16.0.0.0')
                ) {
                    $x = (Get-ItemProperty 'registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop -WarningAction SilentlyContinue).'UpdateChannel'

                    if ($x -ieq 'http://officecdn.microsoft.com/pr/5440FD1F-7ECB-4221-8110-145EFAA6372F') {
                        $OutlookIsBetaversion = $true
                    }

                    if ((Get-Item "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Common\OfficeUpdate" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Property -contains 'UpdateBranch') {
                        $x = (Get-ItemProperty "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Common\OfficeUpdate" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).'UpdateBranch'

                        if ($x -ieq 'InsiderFast') {
                            $OutlookIsBetaversion = $true
                        }
                    }
                }

                $OutlookDisableRoamingSignatures = 0

                foreach ($RegistryFolder in @(
                        "registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                        "registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                        "registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                        "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup"
                    )
                ) {
                    try { WatchCatchableExitSignal } catch { }

                    $x = (Get-ItemProperty $RegistryFolder -ErrorAction SilentlyContinue).'DisableRoamingSignaturesTemporaryToggle'

                    if (($x -in (0, 1)) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                        $OutlookDisableRoamingSignatures = $x
                    }

                    $x = (Get-ItemProperty $RegistryFolder -ErrorAction SilentlyContinue).'DisableRoamingSignatures'

                    if (($x -in (0, 1)) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                        $OutlookDisableRoamingSignatures = $x
                    }
                }

                if ($NewOutlook -and ($((Get-ItemProperty "registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Preferences" -ErrorAction SilentlyContinue).'UseNewOutlook') -eq 1)) {
                    $OutlookUseNewOutlook = $true
                    $OutlookDisableRoamingSignatures = 1
                } else {
                    $OutlookUseNewOutlook = $false
                }
            } else {
                $OutlookDefaultProfile = $null
                $OutlookDisableRoamingSignatures = 1
                $OutlookIsBetaVersion = $null

                if ($NewOutlook) {
                    $OutlookUseNewOutlook = $true
                } else {
                    $OutlookUseNewOutlook = $false
                }
            }

            Write-Host "    Registry version: $OutlookRegistryVersion"
            Write-Host "    File version: $OutlookFileVersion"
            if (($OutlookFileVersion -lt '16.0.0.0') -and ($EmbedImagesInHtml -eq $true)) {
                Write-Host '      Outlook 2013 or earlier detected.' -ForegroundColor Yellow
                Write-Host '      Consider parameter ''-EmbedImagesInHtml false'' to avoid problems with images in templates.' -ForegroundColor Yellow
                Write-Host '      Microsoft supports Outlook 2013 until April 2023, older versions are already out of support.' -ForegroundColor Yellow
            }
            Write-Host "    Bitness: $OutlookBitness"
            Write-Host "    Default profile: $OutlookDefaultProfile"
            Write-Host "    Is C2R Beta: $OutlookIsBetaversion"
            Write-Host "    DisableRoamingSignatures: $OutlookDisableRoamingSignatures"
            if (($OutlookDisableRoamingSignatures -eq 0) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                Write-Host '      Outlook syncs signatures itself, so it may overwrite signatures created by this software.' -ForegroundColor Yellow
                Write-Host '      Consider setting parameters DisableRoamingSignatures and MirrorCloudSignatures to true instead.' -ForegroundColor Yellow
                Write-Host '      Also consider using the MailboxSpecificSignaturesNames parameter.' -ForegroundColor Yellow
            }

            Write-Host "    UseNewOutlook: $OutlookUseNewOutlook"
            Write-Host '  New Outlook'
            Write-Host "    Version: $($NewOutlook.Version)"
            Write-Host "    Status: $($NewOutlook.Status)"
            Write-Host "    UseNewOutlook: $OutlookUseNewOutlook"
        } elseif ($IsMacOS) {
            Write-Host '  Outlook'

            $macOsIsRunningNewOutlook = ($(defaults read com.microsoft.Outlook IsRunningNewOutlook *>&1).ToString() -eq 1)

            $OutlookFileVersion = @(@($(
                        @'
tell application "Microsoft Outlook"
            get version
end tell
'@ | osascript *>&1)) | ForEach-Object { $_.tostring() })[0]

            Write-Host "    Version: $($OutlookFileVersion)"

            try { WatchCatchableExitSignal } catch { }

            $macOSSignaturesScriptable = @(@($(
                        @'
tell application "Microsoft Outlook"
    set guid to do shell script "uuidgen"
    set newSignature to make new signature with properties {name:guid, content:"Set-OutlookSignatures test signature. Please delete."}

    if exists newSignature then
        delete newSignature
        return "Success"
    else
        return "Failure"
    end if
end tell
'@ | osascript *>&1)) | ForEach-Object { $_.tostring() })[0] -eq 'Success'

            try { WatchCatchableExitSignal } catch { }

            $macOSOutlookMailboxes = @(@($(
                        @'
tell application "Microsoft Outlook"
    try
        set exchangeAccounts to get exchange accounts
    on error
        set exchangeAccounts to {}
    end try

    try
        set popAccounts to get pop accounts
    on error
        set popAccounts to {}
    end try

    try
        set imapAccounts to get imap accounts
    on error
        set imapAccounts to {}
    end try

    try
        set ldapAccounts to get ldap accounts
    on error
        set ldapAccounts to {}
    end try

    try
        set delegatedAccounts to get delegated accounts
    on error
        set delegatedAccounts to {}
    end try

    try
        set otherAccounts to get other users folder accounts
    on error
        set otherAccounts to {}
    end try

    set allAccounts to exchangeAccounts & popAccounts & imapAccounts & ldapAccounts & delegatedAccounts & otherAccounts

    repeat with singleAccount in allAccounts
        set x to email address of singleAccount
        log x
    end repeat
end tell
'@ | osascript *>&1)) | ForEach-Object { $_.tostring() })

            try { WatchCatchableExitSignal } catch { }

            $OutlookFilePath = $null
            $OutlookRegistryVersion = $null
            $OutlookDefaultProfile = $null
            $OutlookProfiles = @()
            $OutlookIsBetaversion = $false
            $OutlookDisableRoamingSignatures = 1
            $OutlookUseNewOutlook = $false
            $script:WordRegistryVersion = $null
            $WordFilePath = $null

            if ($macOSSignaturesScriptable) {
                Write-Host '    Outlook for Mac with scriptable signatures detected.'

                $EmbedImagesInHtml = $true

                if ($macOSOutlookMailboxes.count -gt 0) {
                    Write-Host '    Outlook has accounts configured.'
                } else {
                    if ($macOsIsRunningNewOutlook) {
                        Write-Host '    No accounts detected via AppleScript, but New Outlook is enabled. Trying alternate detection method.'

                        if (Test-Path '~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/ProfilePreferences.plist') {
                            $macOSOutlookMailboxes = @($(Get-Content '~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/ProfilePreferences.plist' | Where-Object { $_ -match '.*actionsEndPointURLFor.*' } | ForEach-Object { $_.trim() -ireplace '<key>ActionsEndPointURLFor', '' -ireplace '</key>', '' }))

                            if ($macOSOutlookMailboxes.count -gt 0) {
                                Write-Host '      Accounts found. If too many accounts are found:'
                                Write-Host '        1. Quit Outlook'
                                Write-Host '        2. Delete ''~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/ProfilePreferences.plist'''
                                Write-Host '        3. Start Outlook and run Set-OutlookSignatures'
                            }
                        } else {
                            Write-Host "      Failed. '~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/ProfilePreferences.plist' not found." -ForegroundColor Yellow
                        }
                    }

                    if (-not ($macOSOutlookMailboxes.count -gt 0)) {
                        Write-Host '    Outlook does not have accounts configured, or accounts can not be scripted. Continuing with Outlook Web only.' -ForegroundColor Yellow
                        Write-Host "      Consider using 'sample code/SwitchTo-ClassicOutlookForMac.ps1' to temporarily switch from New Outlook to Classic Outlook." -ForegroundColor Yellow

                        $OutlookUseNewOutlook = $true
                        $macOSOutlookMailboxes = @()
                    }
                }
            } else {
                Write-Host '    Outlook for Mac not installed, or signatures can not be scripted. Continuing with Outlook Web only.' -ForegroundColor Yellow

                $OutlookUseNewOutlook = $true
                $macOSOutlookMailboxes = @()
            }
        } else {
            $OutlookFilePath = $null
            $OutlookRegistryVersion = $null
            $OutlookDefaultProfile = $null
            $OutlookProfiles = @()
            $OutlookIsBetaversion = $false
            $OutlookDisableRoamingSignatures = 1
            $OutlookUseNewOutlook = $true
            $script:WordRegistryVersion = $null
            $WordFilePath = $null
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ((($UseHtmTemplates -eq $true) -and (-not $CreateRtfSignatures)) -or (-not $IsWindows)) {
        Write-Host '  UseHtmTemplates set to true or not running on Windows, skip Word checks'
    } else {
        Write-Host '  Word'

        $script:WordRegistryVersion = $null

        $script:WordAlertIfNotDefaultOriginal = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -ErrorAction SilentlyContinue).AlertIfNotDefault

        $script:WordRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Word.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
        if ($script:WordRegistryVersion.major -gt 16) {
            Write-Host "    Word version $($script:WordRegistryVersion) is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
            $script:ExitCode = 9
            $script:ExitCodeDescription = 'Word version newer than 16 is not yet known.'
            exit
        } elseif ($script:WordRegistryVersion.major -eq 16) {
            $script:WordRegistryVersion = '16.0'
        } elseif ($script:WordRegistryVersion.major -eq 15) {
            $script:WordRegistryVersion = '15.0'
        } elseif ($script:WordRegistryVersion.major -eq 14) {
            $script:WordRegistryVersion = '14.0'
        } elseif ($script:WordRegistryVersion.major -lt 14) {
            Write-Host "    Word version $($script:WordRegistryVersion) is older than Word 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
            $script:ExitCode = 10
            $script:ExitCodeDescription = 'Word version older than 2010 is not supported.'
            exit
        }

        try {
            # [Microsoft.Win32.RegistryView]::Registry32 makes sure view the registry as a 32 bit application would
            # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
            # Covers:
            #   Office x86 on Windows x86
            #   Office x86 on Windows x64
            #   Any PowerShell process bitness
            $WordFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry32)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '').trim('"').trim('''') -ErrorAction Stop
        } catch {
            try {
                # [Microsoft.Win32.RegistryView]::Registry64 makes sure we view the registry as a 64 bit application would
                # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
                # Covers:
                #   Office x64 on Windows x64
                #   Any PowerShell process bitness
                $WordFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry64)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '').trim('"').trim('''') -ErrorAction Stop
            } catch {
                $WordFilePath = $null
            }
        }

        if ($WordFilePath) {
            Write-Host "    Set 'DontUseScreenDpiOnOpen' registry value to '1'"
            $null = "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DontUseScreenDpiOnOpen' -Type DWORD -Value 1 -Force

            try {
                $WordBitnessInfo = GetBitness -fullname $WordFilePath
                $WordFileVersion = [System.Version]::Parse((((($WordBitnessInfo.'File Version'.ToString() + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
                $WordBitness = $WordBitnessInfo.Architecture
                Remove-Variable -Name 'WordBitnessInfo'
            } catch {
                $WordBitness = 'Error'
                $WordFileVersion = $null
            }
        } else {
            $WordBitness = $null
            $WordFileVersion = $null
        }

        Write-Host "    Registry version: $script:WordRegistryVersion"
        Write-Host "    File version: $WordFileVersion"
        Write-Host "    Bitness: $WordBitness"
    }

    try { WatchCatchableExitSignal } catch { }

    Write-Host
    Write-Host "Get Outlook signature file path(s) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $SignaturePaths = @()

    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled. Skip task, use AdditionalSignaturePath instead'
        if ($AdditionalSignaturePath) {
            $SignaturePaths += $AdditionalSignaturePath
        }
    } elseif ($OutlookProfiles -and ($OutlookUseNewOutlook -ne $true)) {
        $x = (Get-ItemProperty "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\common\general" -ErrorAction SilentlyContinue).'Signatures'

        if ($x) {
            Push-Location ((Join-Path -Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Microsoft'))
            $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($x))

            if (Test-Path $x -IsValid) {
                if (-not (Test-Path $x -type container)) {
                    New-Item -Path $x -ItemType directory -Force | Out-Null
                }

                if ($x -inotin $SignaturePaths) {
                    $SignaturePaths += $x
                    Write-Host "  '$x'"
                }
            }

            Pop-Location
        }
    } else {
        $SignaturePaths = @(((New-Item -ItemType Directory (Join-Path -Path $script:tempDir -ChildPath ((New-Guid).guid))).fullname))

        if ($Iswindows) {
            Write-Host "  '$($SignaturePaths[-1])' (Outlook Web/New Outlook)"
        } elseif ($IsMacOS) {
            if ($macOSSignaturesScriptable) {
                Write-Host "  '$($SignaturePaths[-1])' (Outlook for Mac with scriptable signatures)"
            } else {
                Write-Host "  '$($SignaturePaths[-1])' (Outlook Web, because no Outlook, no accounts configured or signatures not scriptable)"
            }
        } elseif ($IsLinux) {
            Write-Host "  '$($SignaturePaths[-1])' (Outlook Web)"
        }
    }

    try { WatchCatchableExitSignal } catch { }

    # If Outlook is installed, synch profile folders anyway
    # Also makes sure that signatures are already there when starting Outlook for the first time
    if ((-not $SimulateUser) -and $OutlookFileVersion) {
        $x = (Get-ItemProperty "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\common\general" -ErrorAction SilentlyContinue).'Signatures'

        if ($x) {
            Push-Location ((Join-Path -Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Microsoft'))
            $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($x))

            if (Test-Path $x -IsValid) {
                if (-not (Test-Path $x -type container)) {
                    New-Item -Path $x -ItemType directory -Force | Out-Null
                }

                if ($x -inotin $SignaturePaths) {
                    $SignaturePaths += $x
                    Write-Host "  '$x'"
                }
            }

            Pop-Location
        }

        $SignaturePaths = @($SignaturePaths | Select-Object -Unique)
    }


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Enumerate domains @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $x = $TrustsToCheckForGroups
    [System.Collections.ArrayList]$TrustsToCheckForGroups = @()
    $LookupDomainsToTrusts = @{}

    if ($GraphOnly -eq $false) {
        # Users own domain/forest is always included
        try {
            $objTrans = New-Object -ComObject 'NameTranslate'
            $objNT = $objTrans.GetType()
            $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $Null)) # 3 = ADS_NAME_INITTYPE_GC
            $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (12, $(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value))) # 12 = ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME
            try { WatchCatchableExitSignal } catch { }
            $UserForest = (([ADSI]"LDAP://$(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1) -split ',DC=')[1..999] -join '.')/RootDSE").rootDomainNamingContext -ireplace [Regex]::Escape('DC='), '' -ireplace [Regex]::Escape(','), '.').tolower()
            try { WatchCatchableExitSignal } catch { }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
            Remove-Variable -Name 'objTrans'
            Remove-Variable -Name 'objNT'

            if ($UserForest -ne '') {
                Write-Host "  User forest: $UserForest"

                if ($TrustsToCheckForGroups -inotcontains $UserForest) {
                    $TrustsToCheckForGroups += $UserForest.tolower()
                }

                if (-not $LookupDomainsToTrusts.ContainsKey($UserForest.tolower())) {
                    $LookupDomainsToTrusts.add($UserForest.tolower(), $UserForest.tolower())
                }

                $Search.SearchRoot = "GC://$($UserForest)"
                $Search.Filter = '(ObjectClass=trustedDomain)'

                try { WatchCatchableExitSignal } catch { }
                $TrustedDomains = @($Search.FindAll())
                try { WatchCatchableExitSignal } catch { }

                if ($TrustedDomains) {
                    $TrustedDomains = @(
                        @($TrustedDomains) | Where-Object { $_ -ine $UserForest } | Sort-Object -Culture $TemplateFilesSortCulture -Property @{Expression = {
                                $TemporaryArray = @($_.properties.name.Split('.'))
                                [Array]::Reverse($TemporaryArray)
                                $TemporaryArray
                            }
                        }
                    )
                }

                try { WatchCatchableExitSignal } catch { }

                # Internal trusts
                foreach ($TrustedDomain in $TrustedDomains) {
                    if (($TrustedDomain.properties.trustattributes -eq 32) -and ($TrustedDomain.properties.name -ine $UserForest) -and (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower()))) {
                        Write-Host "    Child domain: $($TrustedDomain.properties.name.tolower())"

                        if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                            $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $UserForest.tolower())
                        }
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                # Other trusts
                if ($x[0] -eq '*') {
                    foreach ($TrustedDomain in $TrustedDomains) {
                        # No intra-forest trusts, only bidirectional trusts and outbound trusts
                        if (($($TrustedDomain.properties.trustattributes) -ne 32) -and (($($TrustedDomain.properties.trustdirection) -eq 2) -or ($($TrustedDomain.properties.trustdirection) -eq 3))) {
                            if ($TrustedDomain.properties.trustattributes -eq 8) {
                                # Cross-forest trust
                                Write-Host "  Trusted forest: $($TrustedDomain.properties.name.tolower())"
                                if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                    Write-Host "    Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name.tolower())'"
                                } else {
                                    if ($TrustsToCheckForGroups -inotcontains $TrustedDomain.properties.name) {
                                        $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                    }

                                    if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                        $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                    }
                                }

                                try { WatchCatchableExitSignal } catch { }

                                $temp = @(
                                    @(@(Resolve-DnsName -Name "_gc._tcp.$($TrustedDomain.properties.name)" -Type srv).nametarget) | ForEach-Object { ($_ -split '\.')[1..999] -join '.' } | Where-Object { $_ -ine $TrustedDomain.properties.name } | Select-Object -Unique | Sort-Object -Culture $TemplateFilesSortCulture -Property @{Expression = {
                                            $TemporaryArray = @($_.Split('.'))
                                            [Array]::Reverse($TemporaryArray)
                                            $TemporaryArray
                                        }
                                    }
                                )

                                $temp | ForEach-Object {
                                    Write-Host "    Child domain: $($_.tolower())"

                                    if (-not $LookupDomainsToTrusts.ContainsKey($_.tolower())) {
                                        $LookupDomainsToTrusts.add($_.tolower(), $TrustedDomain.properties.name.tolower())
                                    }
                                }
                            } else {
                                # No cross-forest trust
                                Write-Host "  Trusted domain: $($TrustedDomain.properties.name)"
                                if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                    Write-Host "    Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                                } else {
                                    if ($TrustsToCheckForGroups -inotcontains $TrustedDomain.properties.name) {
                                        $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                    }

                                    if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                        $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                    }
                                }
                            }
                        }
                    }
                }

                for ($a = 0; $a -lt $x.Count; $a++) {
                    if (($a -eq 0) -and ($x[$a] -ieq '*')) {
                        continue
                    }

                    $y = ($x[$a] -ireplace [Regex]::Escape('DC='), '' -ireplace ',', '.').tolower()

                    if ($y -eq $x[$a]) {
                        Write-Host "  User provided trusted domain/forest: $y"
                    } else {
                        Write-Host "  User provided trusted domain/forest: $($x[$a]) -> $y"
                    }

                    if (($a -ne 0) -and ($x[$a] -ieq '*')) {
                        Write-Host '    Entry * is only allowed at first position in list. Skip entry.' -ForegroundColor Red
                        continue
                    }

                    if ($y -imatch '[^a-zA-Z0-9.-]') {
                        Write-Host '    Allowed characters are a-z, A-Z, ., -. Skip entry.' -ForegroundColor Red
                        continue
                    }

                    if (-not ($y.StartsWith('-'))) {
                        if ($TrustsToCheckForGroups -icontains $y) {
                            Write-Host '    Trusted domain/forest already in list.'
                        } else {
                            if ($TrustedDomains.properties.name -icontains $y) {
                                foreach ($TrustedDomain in @($TrustedDomains | Where-Object { $_.properties.name -ieq $y })) {
                                    # No intra-forest trusts, only bidirectional trusts and outbound trusts
                                    if (($($TrustedDomain.properties.trustattributes) -ne 32) -and (($($TrustedDomain.properties.trustdirection) -eq 2) -or ($($TrustedDomain.properties.trustdirection) -eq 3))) {
                                        if ($TrustedDomain.properties.trustattributes -eq 8) {
                                            # Cross-forest trust
                                            Write-Host "    Trusted forest: $($TrustedDomain.properties.name)"
                                            if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                                Write-Host "      Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                                            } else {
                                                if ($TrustsToCheckForGroups -inotcontains $TrustedDomain.properties.name) {
                                                    $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                                }

                                                if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                                    $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                                }
                                            }

                                            try { WatchCatchableExitSignal } catch { }

                                            $temp = @(
                                                @(@(Resolve-DnsName -Name "_gc._tcp.$($TrustedDomain.properties.name)" -Type srv).nametarget) | ForEach-Object { ($_ -split '\.')[1..999] -join '.' } | Where-Object { $_ -ine $TrustedDomain.properties.name } | Select-Object -Unique | Sort-Object -Culture $TemplateFilesSortCulture -Property @{Expression = {
                                                        $TemporaryArray = @($_.Split('.'))
                                                        [Array]::Reverse($TemporaryArray)
                                                        $TemporaryArray
                                                    }
                                                }
                                            )

                                            $temp | ForEach-Object {
                                                Write-Host "      Child domain: $($_.tolower())"

                                                if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                                    $LookupDomainsToTrusts.add($_.tolower(), $TrustedDomain.properties.name.tolower())
                                                }
                                            }
                                        } else {
                                            # No cross-forest trust
                                            Write-Host "    Trusted domain: $($TrustedDomain.properties.name)"
                                            if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                                Write-Host "      Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                                            } else {
                                                if ($TrustsToCheckForGroups -inotcontains $TrustedDomain.properties.name) {
                                                    $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                                }

                                                if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                                    $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                Write-Host '    No trust to this domain/forest found.' -ForegroundColor Yellow
                            }
                        }
                    } else {
                        Write-Host '    Remove trusted domain/forest.'
                        for ($z = 0; $z -lt $TrustsToCheckForGroups.Count; $z++) {
                            if ($TrustsToCheckForGroups[$z] -ieq $y.substring(1)) {
                                $TrustsToCheckForGroups.RemoveAt($z)
                                $LookupDomainsToTrusts = $LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $y.substring(1) }
                            }
                        }
                    }
                }

                $TrustsToCheckForGroups = @($TrustsToCheckForGroups | Where-Object { $_ })


                try { WatchCatchableExitSignal } catch { }


                Write-Host
                Write-Host "Check trusts for open LDAP port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                CheckADConnectivity @(@(@($TrustsToCheckForGroups) + @($LookupDomainsToTrusts.GetEnumerator() | ForEach-Object { $_.Name })) | Select-Object -Unique) 'LDAP' '  ' | Out-Null


                try { WatchCatchableExitSignal } catch { }


                Write-Host
                Write-Host "Check trusts for open Global Catalog port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                CheckADConnectivity $TrustsToCheckForGroups 'GC' '  ' | Out-Null
            } else {
                Write-Host '  Problem connecting to logged-in user''s Active Directory (no error message, but forest root domain name is empty).' -ForegroundColor Yellow
                Write-Host '  Assuming Graph/Entra ID from now on.' -ForegroundColor Yellow
                $GraphOnly = $true
            }
        } catch {
            Write-Verbose "  $($error[0])"
            $y = ''
            Write-Host "  Problem connecting to logged-in user's Active Directory, use parameter '-verbose' to see error message." -ForegroundColor Yellow
            Write-Host '  Assuming Graph/Entra ID from now on.' -ForegroundColor Yellow
            $GraphOnly = $true
        }
    } else {
        Write-Host "  Parameter GraphOnly set to '$GraphOnly', ignore user's Active Directory in favor of Graph/Entra ID."
    }


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Get properties of currently logged-in user and assigned manager @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if (-not $SimulateUser) {
        Write-Host '  Currently logged-in user'
    } else {
        Write-Host "  Simulate '$SimulateUser' as currently logged-in user"
    }

    if ($GraphOnly -eq $false) {
        if ($null -ne $TrustsToCheckForGroups[0]) {
            try {
                if (-not $SimulateUser) {
                    $Search.SearchRoot = "GC://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$(([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName)))"
                    try { WatchCatchableExitSignal } catch { }
                    $ADPropsCurrentUser = $Search.FindOne().Properties
                    try { WatchCatchableExitSignal } catch { }
                    $ADPropsCurrentUser = [hashtable]::new($ADPropsCurrentUser, [StringComparer]::OrdinalIgnoreCase)

                    $Search.SearchRoot = "LDAP://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$(([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName)))"
                    try { WatchCatchableExitSignal } catch { }
                    $ADPropsCurrentUserLdap = $Search.FindOne().Properties
                    try { WatchCatchableExitSignal } catch { }
                    $ADPropsCurrentUserLdap = [hashtable]::new($ADPropsCurrentUserLdap, [StringComparer]::OrdinalIgnoreCase)

                    foreach ($keyName in @($ADPropsCurrentUserLdap.Keys)) {
                        if (
                            ($keyName -inotin $ADPropsCurrentUser.Keys) -or
                            (-not ($ADPropsCurrentUser[$keyName]) -and ($ADPropsCurrentUserLdap[$keyName]))
                        ) {
                            $ADPropsCurrentUser[$keyName] = $ADPropsCurrentUserLdap[$keyName]
                        }
                    }
                } else {
                    try {
                        $objTrans = New-Object -ComObject 'NameTranslate'
                        $objNT = $objTrans.GetType()
                        $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                        $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, $SimulateUser))
                        try { WatchCatchableExitSignal } catch { }
                        $SimulateUserDN = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1)
                        try { WatchCatchableExitSignal } catch { }
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
                        Remove-Variable -Name 'objTrans'
                        Remove-Variable -Name 'objNT'

                        $Search.SearchRoot = "GC://$(($SimulateUserDN -split ',DC=')[1..999] -join '.')"
                        $Search.Filter = "((distinguishedname=$SimulateUserDN))"
                        try { WatchCatchableExitSignal } catch { }
                        $ADPropsCurrentUser = $Search.FindOne().Properties
                        try { WatchCatchableExitSignal } catch { }
                        $ADPropsCurrentUser = [hashtable]::new($ADPropsCurrentUser, [StringComparer]::OrdinalIgnoreCase)

                        $Search.SearchRoot = "LDAP://$(($SimulateUserDN -split ',DC=')[1..999] -join '.')"
                        $Search.Filter = "((distinguishedname=$SimulateUserDN))"
                        try { WatchCatchableExitSignal } catch { }
                        $ADPropsCurrentUserLdap = $Search.FindOne().Properties
                        try { WatchCatchableExitSignal } catch { }
                        $ADPropsCurrentUserLdap = [hashtable]::new($ADPropsCurrentUserLdap, [StringComparer]::OrdinalIgnoreCase)

                        foreach ($keyName in @($ADPropsCurrentUserLdap.Keys)) {
                            if (
                                ($keyName -inotin $ADPropsCurrentUser.Keys) -or
                                (-not ($ADPropsCurrentUser[$keyName]) -and ($ADPropsCurrentUserLdap[$keyName]))
                            ) {
                                $ADPropsCurrentUser[$keyName] = $ADPropsCurrentUserLdap[$keyName]
                            }
                        }
                    } catch {
                        Write-Host "    $($error[0])"
                        Write-Host "    Simulation user '$($SimulateUser)' not found. Exit." -ForegroundColor REd
                        $script:ExitCode = 11
                        $script:ExitCodeDescription = 'Simulation user not found.'
                        exit
                    }
                }
            } catch {
                Write-Host $error[0]
                $ADPropsCurrentUser = $null
                Write-Host '    Problem connecting to Active Directory, or user is a local user. Exit.' -ForegroundColor Red
                $script:ExitCode = 12
                $script:ExitCodeDescription = 'Problem connecting to Active Directory, or user is a local user.'
                exit
            }
        }
    }

    if (
        ($GraphOnly -eq $true) -or
        (($GraphOnly -eq $false) -and ($ADPropsCurrentUser.msexchrecipienttypedetails -ge 2147483648) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true))) -or
        (($GraphOnly -eq $false) -and ($null -eq $ADPropsCurrentUser)) -or
        ($OutlookUseNewOutlook -eq $true) -or
        $(
            if (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('LicenseGroupRequiresGraph'))) {
                $result = [SetOutlookSignatures.BenefactorCircle]::LicenseGroupRequiresGraph()

                if ($result -ine 'false') {
                    $true
                } else {
                    $false
                }
            } else {
                $false
            }
        )
    ) {
        Write-Host '    Graph connection is required'
        Write-Verbose '      Required because at least one is true:'
        Write-Verbose "        GraphOnly is true: $($GraphOnly -eq $true)"
        Write-Verbose "        GraphOnly is false and mailbox is in cloud and SetCurrentUserOOFMessage or SetCurrentUserOutlookWebSignature is true: $(($GraphOnly -eq $false) -and ($ADPropsCurrentUser.msexchrecipienttypedetails -ge 2147483648) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true)))"
        Write-Verbose "        GraphOnly is false and on-prem AD properties of current user are empty: $(($GraphOnly -eq $false) -and ($null -eq $ADPropsCurrentUser))"
        Write-Verbose "        New Outlook is used: $($OutlookUseNewOutlook -eq $true)"
        Write-Verbose "        The only Benefactor Circle license group is in Entra ID: $(
            if (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('LicenseGroupRequiresGraph'))) {
                $result = [SetOutlookSignatures.BenefactorCircle]::LicenseGroupRequiresGraph()

                if ($result -ine 'false') {
                    $true
                } else {
                    $false
                }
            } else {
                $false
            }
        )"

        if (-not $GraphToken) {
            try {
                $GraphToken = GraphGetToken
            } catch {
                $GraphToken = $null
            }
        }

        if ($GraphToken -and (-not $SimulateAndDeployGraphCredentialFile)) {
            Write-Host "      Graph token cache: $($script:msalClientApp.cacheInfo)"
        }

        if ($GraphToken.error -eq $false) {
            Write-Verbose "      Graph Token metadata: $((ParseJwtToken $GraphToken.AccessToken) | ConvertTo-Json)"

            if (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true)) {
                Write-Verbose "      EXO Token metadata: $((ParseJwtToken $GraphToken.AccessTokenExo) | ConvertTo-Json)"

                if (-not $($GraphToken.AccessTokenExo)) {
                    Write-Host '        Problem connecting to Exchange Online with Graph token. Exit.' -ForegroundColor Red
                    $script:ExitCode = 13
                    $script:ExitCodeDescription = 'Problem connecting to Exchange Online with Graph token.'
                    exit
                }
            }

            if ($SimulateAndDeployGraphCredentialFile) {
                Write-Verbose "      App Graph Token metadata: $((ParseJwtToken $GraphToken.AppAccessToken) | ConvertTo-Json)"
                Write-Verbose "      App EXO Token metadata: $((ParseJwtToken $GraphToken.AppAccessTokenExo) | ConvertTo-Json)"
            }
        } else {
            Write-Host '      Problem connecting to Microsoft Graph. Exit.' -ForegroundColor Red
            Write-Host $GraphToken.error -ForegroundColor Red
            $script:ExitCode = 14
            $script:ExitCodeDescription = 'Problem connecting to Microsoft Graph.'
            exit
        }

        if ($SimulateUser) {
            $script:GraphUser = $SimulateUser
        }

        $x = (GraphGetUserProperties $script:GraphUser)

        if (($x.error -eq $false) -and ($x.properties.id)) {
            $AADProps = $x.properties
            $ADPropsCurrentUser = [PSCustomObject]@{}

            foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                $z = $AADProps

                foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                    $z = $z.$y
                }

                $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z -Force
            }

            $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $script:GraphUser).photo -Force
            $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'manager' -Value (GraphGetUserManager $script:GraphUser).properties.userprincipalname -Force
        } else {
            Write-Host "      Problem getting data for '$($script:GraphUser)' from Microsoft Graph. Exit." -ForegroundColor Red
            Write-Host $x.error -ForegroundColor Red
            $script:ExitCode = 15
            $script:ExitCodeDescription = "Problem getting data for '$($script:GraphUser)' from Microsoft Graph."
            exit
        }
    }

    if ($ADPropsCurrentUser) {
        Write-Host "    DistinguishedName: $($ADPropsCurrentUser.distinguishedname)"
        Write-Host "    UserPrincipalName: $($ADPropsCurrentUser.userprincipalname)"
        Write-Host "    Mail: $($ADPropsCurrentUser.mail)"
    } else {
        Write-Host '    User not found'
    }


    try { WatchCatchableExitSignal } catch { }

    $CurrentUserSIDs = @()

    if (($ADPropsCurrentUser.objectsid -ne '') -and ($null -ne $ADPropsCurrentUser.objectsid)) {
        if ($ADPropsCurrentUser.objectsid.tostring().startswith('S-', 'CurrentCultureIgnorecase')) {
            $CurrentUserSids += $ADPropsCurrentUser.objectsid.tostring()
        } else {
            $CurrentUserSids += (New-Object system.security.principal.securityidentifier $($ADPropsCurrentUser.objectsid), 0).value
        }
    }

    if (($ADPropsCurrentUser.onpremisessecurityidentifier -ne '') -and ($null -ne $ADPropsCurrentUser.onpremisessecurityidentifier)) {
        $CurrentUserSids += $ADPropsCurrentUser.onpremisessecurityidentifier.tostring()
    }

    foreach ($SidHistorySid in @($ADPropsCurrentUser.sidhistory | Where-Object { $_ })) {
        try { WatchCatchableExitSignal } catch { }

        if ($SidHistorySid.tostring().startswith('S-', 'CurrentCultureIgnorecase')) {
            $CurrentUserSids += $SidHistorySid.tostring()
        } else {
            $CurrentUserSids += (New-Object system.security.principal.securityidentifier $SidHistorySid, 0).value
        }
    }

    if (-not $SimulateUser) {
        Write-Host '  Manager of currently logged-in user'
    } else {
        Write-Host '  Manager of simulated currently logged-in user'
    }

    $ADPropsCurrentUserManager = $null

    if ($ADPropsCurrentUser -and ($ADPropsCurrentUser.manager)) {
        if ($ADPropsCurrentUser.manager -imatch '(\S+?)@(\S+?)\.(\S+?)') {
            # Manager is in UPN format, search via Graph
            # Graph connection must already be established, else the manager would not be in UPN format

            Write-Verbose "    Search manager '$($ADPropsCurrentUser.manager)' via Graph"

            try {
                $AADProps = (GraphGetUserProperties $ADPropsCurrentUser.manager).properties
                $ADPropsCurrentUserManager = [PSCustomObject]@{}

                foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                    $z = $AADProps

                    foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                        $z = $z.$y
                    }

                    $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z -Force
                }

                $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsCurrentUserManager.userprincipalname).photo -Force
                $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name 'manager' -Value $null -Force
            } catch {
                $ADPropsCurrentUserManager = $null
            }
        } else {
            # Manager is not in UPN format, try search on-prem
            # But only if ($GraphOnly -ne $true)

            Write-Verbose "    Search manager '$($ADPropsCurrentUser.manager)' on-prem"

            if ($GraphOnly -ne $true) {
                try {
                    $Search.SearchRoot = "GC://$(($ADPropsCurrentUser.manager -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$($ADPropsCurrentUser.manager)))"
                    try { WatchCatchableExitSignal } catch { }
                    $ADPropsCurrentUserManager = $Search.FindOne().Properties
                    try { WatchCatchableExitSignal } catch { }
                    $ADPropsCurrentUserManager = [hashtable]::new($ADPropsCurrentUserManager, [StringComparer]::OrdinalIgnoreCase)


                    $Search.SearchRoot = "LDAP://$(($ADPropsCurrentUser.manager -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$($ADPropsCurrentUser.manager)))"
                    try { WatchCatchableExitSignal } catch { }
                    $ADPropsCurrentUserManagerLdap = $Search.FindOne().Properties
                    try { WatchCatchableExitSignal } catch { }
                    $ADPropsCurrentUserManagerLdap = [hashtable]::new($ADPropsCurrentUserManagerLdap, [StringComparer]::OrdinalIgnoreCase)

                    foreach ($keyName in @($ADPropsCurrentUserManagerLdap.Keys)) {
                        if (
                        ($keyName -inotin $ADPropsCurrentUserManager.Keys) -or
                        (-not ($ADPropsCurrentUserManager[$keyName]) -and ($ADPropsCurrentUserManagerLdap[$keyName]))
                        ) {
                            $ADPropsCurrentUserManager[$keyName] = $ADPropsCurrentUserManagerLdap[$keyName]
                        }
                    }
                } catch {
                    $ADPropsCurrentUserManager = $null
                }
            } else {
                $ADPropsCurrentUserManager = $null

                Write-Verbose "    Undefined combination: GraphOnly is set to true, but manager '$($ADPropsCurrentUser.manager)' is not in UPN format."
            }
        }
    }

    if ($ADPropsCurrentUserManager) {
        Write-Host "    DistinguishedName: $($ADPropsCurrentUserManager.distinguishedname)"
        Write-Host "    UserPrincipalName: $($ADPropsCurrentUserManager.userprincipalname)"
        Write-Host "    Mail: $($ADPropsCurrentUserManager.mail)"
    } else {
        Write-Host '    No manager found'
    }


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Get email addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $MailAddresses = @()
    $RegistryPaths = @()
    $LegacyExchangeDNs = @()

    if ($SimulateUser -and $SimulateMailboxes) {
        Write-Host '  Simulation mode enabled and SimulateMailboxes defined, use SimulateMailboxes as mailbox list'
        for ($i = 0; $i -lt $SimulateMailboxes.count; $i++) {
            $MailAddresses += $SimulateMailboxes[$i].ToLower()
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''
        }
    } elseif ($IsWindows -and $OutlookProfiles -and ($OutlookUseNewOutlook -ne $true)) {
        Write-Host '  Get email addresses from Outlook'

        foreach ($OutlookProfile in $OutlookProfiles) {
            try { WatchCatchableExitSignal } catch { }

            Write-Host "    Profile '$($OutlookProfile)'"

            foreach ($RegistryFolder in @(Get-ItemProperty "hkcu:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$($OutlookProfile)\9375CFF0413111d3B88A00104B2A6676\*" -ErrorAction SilentlyContinue | Where-Object { if ($OutlookFileVersion -ge '16.0.0.0') { ($_.'Account Name' -like '*@*.*') } else { (($_.'Account Name' -join ',') -like '*,64,*,46,*') } })) {
                try { WatchCatchableExitSignal } catch { }

                if ($OutlookFileVersion -ge '16.0.0.0') {
                    $MailAddresses += ($RegistryFolder.'Account Name').ToLower()
                } else {
                    $MailAddresses += (@(ForEach ($char in @(($RegistryFolder.'Account Name' -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$($char)" }) -join '').ToLower()
                }

                $RegistryPaths += $RegistryFolder.PSPath

                if ($RegistryFolder.'Identity Eid') {
                    $LegacyExchangeDN = ('/O=' + ((@(foreach ($char in @(($RegistryFolder.'Identity Eid' -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$($char)" }) -join '') -split '/O=')[-1]).ToString().trim()
                    if ($LegacyExchangeDN.length -le 3) {
                        $LegacyExchangeDN = ''
                    }
                } else {
                    $LegacyExchangeDN = ''
                }

                $LegacyExchangeDNs += $LegacyExchangeDN

                Write-Host "      $($MailAddresses[-1])"
                Write-Verbose "        Registry: $($RegistryFolder.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $RegistryFolder.PSDrive)"
                Write-Verbose "        LegacyExchangeDN: $($LegacyExchangeDNs[-1])"
            }

            if ($SignaturesForAutomappedAndAdditionalMailboxes) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SignaturesForAutomappedAndAdditionalMailboxes')))) {
                    Write-Host '    Automapped and additional mailboxes will not be found.' -ForegroundColor Yellow
                    Write-Host "    The 'SignaturesForAutomappedAndAdditionalMailboxes' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                    Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SignaturesForAutomappedAndAdditionalMailboxes()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '      Error finding automapped and additional mailboxes.' -ForegroundColor Yellow
                        Write-Host "      $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "    Parameter 'SignaturesForAutomappedAndAdditionalMailboxes' is not enabled, skipping task."
            }
        }
    } elseif ($IsMacOS -and $macOSSignaturesScriptable -and ($macOSOutlookMailboxes.count -gt 0)) {
        Write-Host '  Get email addresses from Outlook'

        $macOSOutlookMailboxes | ForEach-Object {
            $MailAddresses += $_
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''

            Write-Host "    $($MailAddresses[-1])"
            Write-Verbose "      Registry: $($RegistryPaths[-1])"
            Write-Verbose "      LegacyExchangeDN: $($LegacyExchangeDNs[-1])"
        }
    } else {
        if ($IsWindows -and $OutlookUseNewOutlook) {
            Write-Host '  Get email addresses from New Outlook and Outlook Web, as New Outlook is set as default'
        } else {
            Write-Host '  Get email addresses from Outlook Web'
        }

        $OutlookProfiles = @()
        $OutlookDefaultProfile = $null

        $script:GraphUserDummyMailbox = $true

        if ($IsWindows -and $OutlookUseNewOutlook -eq $true) {
            $x = @(
                @((Get-Content -Path $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Microsoft\Olk\UserSettings.json') -Force -Encoding utf8 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ConvertFrom-Json).Identities.IdentityMap.PSObject.Properties | Select-Object -Unique | Where-Object { $_.name -match '(\S+?)@(\S+?)\.(\S+?)' }) | ForEach-Object {
                    if ((Get-Content -Path $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath "\Microsoft\OneAuth\accounts\$($_.Value)") -Force -Encoding utf8 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ConvertFrom-Json).association_status -ilike '*"com.microsoft.Olk":"associated"*') {
                        $_.name
                    }
                }
            )
        } else {
            $x = @()
        }

        if ($ADPropsCurrentUser.mail) {
            if ($x -icontains $ADPropsCurrentUser.mail) {
                $x = @($ADPropsCurrentUser.mail.tolower()) + @($x | Where-Object { $_ -ine $ADPropsCurrentUser.mail })
            } else {
                $x = @($ADPropsCurrentUser.mail.tolower()) + $x
            }
        } else {
            Write-Host '    User does not have mail attribute configured' -ForegroundColor Yellow
            $script:GraphUserDummyMailbox = $false
        }

        $x | ForEach-Object {
            $MailAddresses += $_.ToLower()
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''

            Write-Host "    $($MailAddresses[-1])"
            Write-Verbose "      Registry: $($RegistryFolder.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $RegistryFolder.PSDrive)"
            Write-Verbose "      LegacyExchangeDN: $($LegacyExchangeDNs[-1])"

            if ($ADPropsCurrentUser.mail -and ($_ -ieq $ADPropsCurrentUser.mail)) {
                $PrimaryMailboxAddress = $ADPropsCurrentUser.mail

                if (-not $script:WebServicesDllPath) {
                    Write-Host '    Set up environment for connection to Outlook Web'

                    try { WatchCatchableExitSignal } catch { }

                    $script:WebServicesDllPath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))

                    try {
                        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\netstandard2.0\Microsoft.Exchange.WebServices.Data.dll')) -Destination $script:WebServicesDllPath -Force
                        if (-not $IsLinux) {
                            Unblock-File -LiteralPath $script:WebServicesDllPath
                        }
                    } catch {
                        Write-Verbose "      $($_)"
                    }
                }

                ConnectEWS -MailAddress $MailAddresses[0] -Indent '    '

                if ($SignaturesForAutomappedAndAdditionalMailboxes) {
                    if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SignaturesForAutomappedAndAdditionalMailboxes')))) {
                        Write-Host '    Automapped and additional mailboxes will not be found.' -ForegroundColor Yellow
                        Write-Host "    The 'SignaturesForAutomappedAndAdditionalMailboxes' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                        Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                    } else {
                        try { WatchCatchableExitSignal } catch { }

                        $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SignaturesForAutomappedAndAdditionalMailboxes()

                        if ($FeatureResult -ne 'true') {
                            Write-Host '    Error finding automapped and additional mailboxes.' -ForegroundColor Yellow
                            Write-Host "    $FeatureResult" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "    Parameter 'SignaturesForAutomappedAndAdditionalMailboxes' is not enabled, skipping task."
                }
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ((($SetCurrentUserOutlookWebSignature -eq $true) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses -inotcontains $ADPropsCurrentUser.mail)) {
        # OOF and/or Outlook web signature must be set, but user does not seem to have a mailbox in Outlook
        # Maybe this is a pure Outlook Web user, so we will add a helper entry
        # This entry fakes the users mailbox in his default Outlook profile, so it gets the highest priority later
        Write-Host "  User's mailbox not found in email address list, but Outlook Web signature and/or OOF message should be set. Adding dummy mailbox entry." -ForegroundColor Yellow

        if ($ADPropsCurrentUser.mail) {
            $script:GraphUserDummyMailbox = $true

            $SignaturePaths = @(((New-Item -ItemType Directory (Join-Path -Path $script:tempDir -ChildPath ((New-Guid).guid))).fullname)) + $SignaturePaths

            $MailAddresses = @($ADPropsCurrentUser.mail.tolower()) + $MailAddresses
            $RegistryPaths = @('') + $RegistryPaths
            $LegacyExchangeDNs = @('') + $LegacyExchangeDNs
        } else {
            Write-Host '      User does not have mail attribute configured.' -ForegroundColor Yellow
            $script:GraphUserDummyMailbox = $false
        }
    } else {
        $script:GraphUserDummyMailbox = $false
    }

    try { WatchCatchableExitSignal } catch { }

    if ($MailAddresses.count -eq 0) {
        Write-Host
        Write-Host 'No email addresses found, exiting.'
        Write-Host '  In simulation mode, this might be a permission problem.'
        $script:ExitCode = 16
        $script:ExitCodeDescription = 'No email addresses found.'
        exit
    }


    try { WatchCatchableExitSignal } catch { }


    $ADPropsMailboxes = @()
    $ADPropsMailboxesUserDomain = @()
    $ADPropsMailboxManagers = @()

    Write-Host
    Write-Host "Get properties of each mailbox and its manager @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        Write-Host "  $($MailAddresses[$AccountNumberRunning])"

        $UserDomain = ''
        $ADPropsMailboxes += $null
        $ADPropsMailboxesUserDomain += $null
        $ADPropsMailboxManagers += $null
        $GroupsSIDs = @()

        $CurrentMailboxAlreadyFoundFirstIndex = $null

        for ($i = 0; $i -lt $ADPropsMailboxes.Count; $i++) {
            if ($ADPropsMailboxes[$i].proxyaddresses -icontains "smtp:$($MailAddresses[$AccountNumberRunning])") {
                $CurrentMailboxAlreadyFoundFirstIndex = $i
                break
            }
        }

        if (
            $null -eq $CurrentMailboxAlreadyFoundFirstIndex
        ) {
            if ((($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '') -or ($($MailAddresses[$AccountNumberRunning]) -ne ''))) {
                if ($null -ne $TrustsToCheckForGroups[0]) {
                    # Loop through domains until the first one knows the legacyExchangeDN or the proxy address
                    for ($DomainNumber = 0; (($DomainNumber -lt $TrustsToCheckForGroups.count) -and ($UserDomain -eq '')); $DomainNumber++) {
                        try { WatchCatchableExitSignal } catch { }

                        if (($TrustsToCheckForGroups[$DomainNumber] -ne '')) {
                            Write-Host "    Search for mailbox user object in domain/forest '$($TrustsToCheckForGroups[$DomainNumber])'"
                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustsToCheckForGroups[$DomainNumber])")
                            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                                $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(|(msexchrecipienttypedetails<=32)(msexchrecipienttypedetails>=2147483648))(msExchMailboxGuid=*)(|(legacyExchangeDN=$($LegacyExchangeDNs[$AccountNumberRunning]))(&(legacyExchangeDN=*)(proxyaddresses=x500:$($LegacyExchangeDNs[$AccountNumberRunning])))))"
                            } elseif (($($MailAddresses[$AccountNumberRunning]) -ne '')) {
                                $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(|(msexchrecipienttypedetails<=32)(msexchrecipienttypedetails>=2147483648))(msExchMailboxGuid=*)(legacyExchangeDN=*)(proxyaddresses=smtp:$($MailAddresses[$AccountNumberRunning])))"
                            }

                            try { WatchCatchableExitSignal } catch { }

                            $u = $Search.FindAll()

                            try { WatchCatchableExitSignal } catch { }

                            if ($u.count -eq 0) {
                                Write-Host '      Not found'
                            } elseif ($u.count -gt 1) {
                                Write-Host '      Multiple matches found' -ForegroundColor Yellow

                                foreach ($SingleU in $u) {
                                    Write-Host "      $($SingleU.path)" -ForegroundColor Yellow
                                }

                                Write-Host '        Check why your Active Directory returns multiple results for the following query:' -ForegroundColor Yellow
                                Write-Host "          $($Search.SearchRoot)" -ForegroundColor Yellow
                                Write-Host "          $($Search.Filter)" -ForegroundColor Yellow

                                $LegacyExchangeDNs[$AccountNumberRunning] = ''
                                $MailAddresses[$AccountNumberRunning] = ''
                                $UserDomain = $null
                            } else {
                                $Search.SearchRoot = "GC://$(($(([adsi]"$($u[0].path)").distinguishedname) -split ',DC=')[1..999] -join '.')"
                                $Search.Filter = "((distinguishedname=$(([adsi]"$($u[0].path)").distinguishedname)))"
                                try { WatchCatchableExitSignal } catch { }
                                $ADPropsMailboxes[$AccountNumberRunning] = $Search.FindOne().Properties
                                try { WatchCatchableExitSignal } catch { }
                                $ADPropsMailboxes[$AccountNumberRunning] = [hashtable]::new($ADPropsMailboxes[$AccountNumberRunning], [StringComparer]::OrdinalIgnoreCase)

                                $Search.SearchRoot = "LDAP://$(($(([adsi]"$($u[0].path)").distinguishedname) -split ',DC=')[1..999] -join '.')"
                                $Search.Filter = "((distinguishedname=$(([adsi]"$($u[0].path)").distinguishedname)))"
                                try { WatchCatchableExitSignal } catch { }
                                $tempLdap = $Search.FindOne().Properties
                                try { WatchCatchableExitSignal } catch { }
                                $tempLdap = [hashtable]::new($tempLdap, [StringComparer]::OrdinalIgnoreCase)

                                foreach ($keyName in @($tempLdap.Keys)) {
                                    if (
                                        ($keyName -inotin $ADPropsMailboxes[$AccountNumberRunning].Keys) -or
                                        (-not ($ADPropsMailboxes[$AccountNumberRunning][$keyName]) -and ($tempLdap[$keyName]))
                                    ) {
                                        $ADPropsMailboxes[$AccountNumberRunning][$keyName] = $tempLdap[$keyName]
                                    }
                                }

                                $UserDomain = $TrustsToCheckForGroups[$DomainNumber]
                                $ADPropsMailboxesUserDomain[$AccountNumberRunning] = $TrustsToCheckForGroups[$DomainNumber]
                                $LegacyExchangeDNs[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].legacyexchangedn
                                $MailAddresses[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].mail.tolower()
                                Write-Host "      distinguishedName: $($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                                Write-Host "      UserPrincipalName: $($ADPropsMailboxes[$AccountNumberRunning].userprincipalname)"
                                Write-Host "      Mail: $($ADPropsMailboxes[$AccountNumberRunning].mail)"
                                Write-Host "      Manager: $($ADPropsMailboxes[$AccountNumberRunning].manager)"
                            }
                        }
                    }

                    if ($u.count -eq 0) {
                        Write-Host "      No matching mailbox object found in any Active Directory. Use parameter '-verbose' to see details." -ForegroundColor Yellow
                        Write-Host '      This message can be ignored if the mailbox in question is not part of your environment.' -ForegroundColor Yellow
                        Write-Verbose "        You may have restricted the accessible environment with the 'TrustsToCheckForGroups' parameter."
                        Write-Verbose '        Else, check why the following Active Directory query did not return a result:'
                        Write-Verbose "          $($Search.Filter)"
                        Write-Verbose '        Usual root causes: Mailbox added in Outlook no longer exists or is not in your tenant, Exchange data in Active Directory is not complete, firewall rules, DNS.'
                        Write-Verbose "        Check if all required attributes documented in the 'README' file are available in your on-prem Active Directory and have values."
                        Write-Verbose "          Look for 'msExchMailboxGuid' in the 'README' file for details about the required attributes."
                        Write-Verbose '        For hybrid environments:'
                        Write-Verbose '          Add missing msExchMailboxGuid for cloud mailboxes to on-prem AD: https://learn.microsoft.com/en-US/exchange/troubleshoot/move-mailboxes/migrationpermanentexception-when-moving-mailboxes.'
                        Write-Verbose "          Consider using the '-GraphOnly true' parameter to not query on-prem Active Directory at all."
                    }

                    if (-not $ADPropsMailboxes[$AccountNumberRunning]) {
                        $LegacyExchangeDNs[$AccountNumberRunning] = ''
                        $UserDomain = $null
                    } elseif ($ADPropsMailboxManagers[$AccountNumberRunning].manager) {
                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")

                        try {
                            $Search.filter = "(distinguishedname=$($ADPropsMailboxes[$AccountNumberRunning].manager))"
                            try { WatchCatchableExitSignal } catch { }
                            $ADPropsMailboxManagers[$AccountNumberRunning] = ([ADSI]"$(($Search.FindOne()).path)").Properties
                            try { WatchCatchableExitSignal } catch { }

                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")
                            $Search.filter = "(distinguishedname=$($ADPropsMailboxes[$AccountNumberRunning].manager))"

                            try { WatchCatchableExitSignal } catch { }

                            $ADPropsCurrentMailboxManagerLdap = ([ADSI]"$(($Search.FindOne()).path)").Properties

                            try { WatchCatchableExitSignal } catch { }

                            foreach ($keyName in @($ADPropsCurrentMailboxManagerLdap.Keys)) {
                                if (
                                    ($keyName -inotin $ADPropsMailboxManagers[$AccountNumberRunning].Keys) -or
                                    (-not ($ADPropsMailboxManagers[$AccountNumberRunning][$keyName]) -and ($ADPropsCurrentMailboxManagerLdap[$keyName]))
                                ) {
                                    $ADPropsMailboxManagers[$AccountNumberRunning][$keyName] = $ADPropsCurrentMailboxManagerLdap[$keyName]
                                }
                            }

                            Write-Host "        distinguishedName: $($ADPropsMailboxManagers[$AccountNumberRunning].distinguishedname)"
                            Write-Host "        UserPrincipalName: $($ADPropsMailboxManagers[$AccountNumberRunning].userprincipalname)"
                            Write-Host "        Mail: $($ADPropsMailboxManagers[$AccountNumberRunning].mail)"
                        } catch {
                            $ADPropsMailboxManagers[$AccountNumberRunning] = @()
                        }
                    }
                } else {
                    Write-Host '    Search for mailbox user object in Graph'

                    $ADPropsMailboxes[$AccountNumberRunning] = [PSCustomObject]@{}

                    try { WatchCatchableExitSignal } catch { }

                    $AADProps = (GraphGetUserProperties $($MailAddresses[$AccountNumberRunning])).properties

                    try { WatchCatchableExitSignal } catch { }

                    if ($AADProps) {
                        foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                            $z = $AADProps

                            foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                                $z = $z.$y
                            }

                            $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z -Force
                        }

                        try { WatchCatchableExitSignal } catch { }

                        $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsMailboxes[$AccountNumberRunning].userprincipalname).photo -Force

                        try { WatchCatchableExitSignal } catch { }

                        $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'manager' -Value (GraphGetUserManager $ADPropsMailboxes[$AccountNumberRunning].userprincipalname).properties.userprincipalname -Force

                        try { WatchCatchableExitSignal } catch { }

                        if (-not $LegacyExchangeDNs[$AccountNumberRunning]) {
                            $LegacyExchangeDNs[$AccountNumberRunning] = 'dummy'
                        }

                        $MailAddresses[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].mail.tolower()

                        Write-Host "      DistinguishedName: $($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                        Write-Host "      UserPrincipalName: $($ADPropsMailboxes[$AccountNumberRunning].userprincipalname)"
                        Write-Host "      Mail: $($ADPropsMailboxes[$AccountNumberRunning].mail)"
                        Write-Host "      Manager: $($ADPropsMailboxes[$AccountNumberRunning].manager)"

                        if ($ADPropsMailboxes[$AccountNumberRunning].manager) {
                            # get properties of mailbox manager here

                            try {
                                $AADProps = $null

                                if ($ADPropsMailboxes[$AccountNumberRunning].manager) {
                                    try { WatchCatchableExitSignal } catch { }

                                    $AADProps = (GraphGetUserProperties $ADPropsMailboxes[$AccountNumberRunning].manager).properties

                                    try { WatchCatchableExitSignal } catch { }

                                    $ADPropsMailboxManagers[$AccountNumberRunning] = [PSCustomObject]@{}

                                    foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                                        $z = $AADProps

                                        foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                                            $z = $z.$y
                                        }

                                        $ADPropsMailboxManagers[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z -Force
                                    }

                                    try { WatchCatchableExitSignal } catch { }

                                    $ADPropsMailboxManagers[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsMailboxManagers[$AccountNumberRunning].userprincipalname).photo -Force

                                    try { WatchCatchableExitSignal } catch { }

                                    $ADPropsMailboxManagers[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'manager' -Value $null -Force

                                    try { WatchCatchableExitSignal } catch { }

                                    Write-Host "        DistinguishedName: $($ADPropsMailboxManagers[$AccountNumberRunning].distinguishedname)"
                                    Write-Host "        UserPrincipalName: $($ADPropsMailboxManagers[$AccountNumberRunning].userprincipalname)"
                                    Write-Host "        Mail: $($ADPropsMailboxManagers[$AccountNumberRunning].mail)"
                                }

                                try { WatchCatchableExitSignal } catch { }
                            } catch {
                                $ADPropsMailboxManagers[$AccountNumberRunning] = @()
                                Write-Host '        Skipping, mailbox manager not in Microsoft Graph.' -ForegroundColor yellow
                            }
                        }
                    } else {
                        Write-Host "      No matching mailbox object found via Graph/Entra ID. Use parameter '-verbose' to see details." -ForegroundColor Yellow
                        Write-Host '      This message can be ignored if the mailbox in question is not part of your environment.' -ForegroundColor Yellow
                        Write-Verbose '        Check why the following Graph queries return zero or more than 1 results, or do not contain any properties:'
                        Write-Verbose "          UserPrincipalName from: $("$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users?`$filter=proxyAddresses/any(x:x eq 'smtp:$($MailAddresses[$AccountNumberRunning])')")"
                        Write-Verbose "          Replace XXX with UPN from query above: $("$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/XXX?`$select=" + [System.Net.WebUtility]::UrlEncode($(@($GraphUserProperties | Select-Object -Unique) -join ',')))"
                        Write-Verbose '        Usual root causes: Mailbox added in Outlook no longer exists or is not in your tenant, firewall rules, DNS.'

                        $LegacyExchangeDNs[$AccountNumberRunning] = ''
                        $UserDomain = $null
                        $ADPropsMailboxManagers[$AccountNumberRunning] = ''
                    }
                }

                Write-Host '      Get group membership of mailbox'
                if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                    try { WatchCatchableExitSignal } catch { }

                    if ($null -ne $TrustsToCheckForGroups[0]) {
                        Write-Host "        $($ADPropsMailboxesUserDomain[$AccountNumberRunning]) (mailbox home domain/forest)"

                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")

                        $UserDomain = $ADPropsMailboxesUserDomain[$AccountNumberRunning]
                        $SIDsToCheckInTrusts = @()

                        if ($ADPropsMailboxes[$AccountNumberRunning].objectsid) {
                            $SIDsToCheckInTrusts += (New-Object System.Security.Principal.SecurityIdentifier $($ADPropsMailboxes[$AccountNumberRunning].objectsid), 0).value
                        }

                        foreach ($SidHistorySid in @($ADPropsMailboxes[$AccountNumberRunning].sidhistory | Where-Object { $_ })) {
                            $SIDsToCheckInTrusts += (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                        }

                        try { WatchCatchableExitSignal } catch { }

                        try {
                            # Security groups, global and universal, forest-wide
                            Write-Host '          LDAP query for security groups (global and universal, forest-wide, via tokengroupsglobalanduniversal)'
                            $UserAccount = [ADSI]"LDAP://$($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                            try { WatchCatchableExitSignal } catch { }
                            $UserAccount.GetInfoEx(@('tokengroupsglobalanduniversal'), 0)
                            try { WatchCatchableExitSignal } catch { }

                            foreach ($sidBytes in $UserAccount.Properties.tokengroupsglobalanduniversal) {
                                $sid = (New-Object System.Security.Principal.SecurityIdentifier($sidbytes, 0)).value
                                Write-Verbose "            $($sid) (global or universal group, incl. sIDHistory)"
                                $GroupsSIDs += $sid
                                $SIDsToCheckInTrusts += $sid
                            }

                            try { WatchCatchableExitSignal } catch { }

                            # Distribution groups (static only)
                            try { WatchCatchableExitSignal } catch { }
                            Write-Host '          GC query for static distribution groups (global and universal, forest-wide)'
                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$(($($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.')")
                            $Search.filter = "(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648))(member:1.2.840.113556.1.4.1941:=$($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)))"

                            try { WatchCatchableExitSignal } catch { }

                            foreach ($DistributionGroup in $search.findall()) {
                                try { WatchCatchableExitSignal } catch { }

                                if ($DistributionGroup.properties.objectsid) {
                                    $sid = (New-Object System.Security.Principal.SecurityIdentifier $($DistributionGroup.properties.objectsid), 0).value
                                    Write-Verbose "            $($sid) (static distribution group)"
                                    $GroupsSIDs += $sid
                                    $SIDsToCheckInTrusts += $sid
                                }

                                foreach ($SidHistorySid in @($DistributionGroup.properties.sidhistory | Where-Object { $_ })) {
                                    $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                                    Write-Verbose "            $($sid) (static distribution group sIDHistory)"
                                    $GroupsSIDs += $sid
                                    $SIDsToCheckInTrusts += $sid
                                }
                            }

                            try { WatchCatchableExitSignal } catch { }

                            # Domain local groups
                            if ($IncludeMailboxForestDomainLocalGroups -eq $true) {
                                Write-Host '        LDAP query for domain local groups (security and distribution, one query per domain)'

                                foreach ($DomainToCheckForDomainLocalGroups in @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ieq $LookupDomainsToTrusts[$(($($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.')] }).name)) {
                                    try { WatchCatchableExitSignal } catch { }
                                    Write-Host "          $($DomainToCheckForDomainLocalGroups)"
                                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DomainToCheckForDomainLocalGroups)")
                                    $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)))"

                                    try { WatchCatchableExitSignal } catch { }

                                    foreach ($LocalGroup in $search.findall()) {
                                        try { WatchCatchableExitSignal } catch { }

                                        if ($LocalGroup.properties.objectsid) {
                                            $sid = (New-Object System.Security.Principal.SecurityIdentifier $($LocalGroup.properties.objectsid), 0).value
                                            Write-Verbose "            $($sid) (domain local group)"
                                            $GroupsSIDs += $sid
                                            $SIDsToCheckInTrusts += $sid
                                        }

                                        foreach ($SidHistorySid in @($LocalGroup.properties.sidhistory | Where-Object { $_ })) {
                                            $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                                            Write-Verbose "            $($sid) (domain local group sIDHistory)"
                                            $GroupsSIDs += $sid
                                            $SIDsToCheckInTrusts += $sid
                                        }
                                    }
                                }
                            }
                        } catch {
                            Write-Host $error[0]
                            Write-Host "            Error getting group information from $((($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.'), check firewalls, DNS and AD trust" -ForegroundColor Red
                        }

                        try { WatchCatchableExitSignal } catch { }

                        $GroupsSIDs = @($GroupsSIDs | Select-Object -Unique | Sort-Object)

                        # Loop through all domains outside the mailbox account's home forest to check if the mailbox account has a group membership there
                        # Across a trust, a user can only be added to a domain local group.
                        # Domain local groups cannot be used outside their own domain, so we don't need to query recursively
                        # But when it's a cross-forest trust, we need to query every every domain on that other side of the trust
                        #   This is handled before by adding every single domain of a cross-forest trusted forest to $TrustsToCheckForGroups
                        if ($SIDsToCheckInTrusts.count -gt 0) {
                            $SIDsToCheckInTrusts = @($SIDsToCheckInTrusts | Select-Object -Unique)
                            $LdapFilterSIDs = '(|'

                            foreach ($SidToCheckInTrusts in $SIDsToCheckInTrusts) {
                                try { WatchCatchableExitSignal } catch { }

                                try {
                                    $SidHex = @()
                                    $ot = New-Object System.Security.Principal.SecurityIdentifier($SidToCheckInTrusts)
                                    $c = New-Object 'byte[]' $ot.BinaryLength
                                    $ot.GetBinaryForm($c, 0)
                                    foreach ($char in $c) {
                                        $SidHex += $('\{0:x2}' -f $char)
                                    }
                                    # Foreign Security Principals have an objectSID, but no sIDHistory
                                    # The sIDHistory of the current mailbox is part of $SIDsToCheckInTrusts and therefore also considered in $LdapFilterSIDs
                                    $LdapFilterSIDs += ('(objectsid=' + $($SidHex -join '') + ')')
                                } catch {
                                    Write-Host $error[0]
                                    Write-Host '        Error creating LDAP filter for search across trusts.' -ForegroundColor Red
                                }
                            }
                            $LdapFilterSIDs += ')'
                        } else {
                            $LdapFilterSIDs = ''
                        }

                        if ($LdapFilterSids -ilike '*(objectsid=*') {
                            # Across each trust, search for all Foreign Security Principals matching a SID from our list
                            foreach ($TrustToCheckForFSPs in @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $LookupDomainsToTrusts[$(($($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.')] }).value | Select-Object -Unique)) {
                                try { WatchCatchableExitSignal } catch { }

                                Write-Host "        $($TrustToCheckForFSPs) (trusted domain/forest of mailbox home forest) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustToCheckForFSPs)")
                                $Search.filter = "(&(objectclass=foreignsecurityprincipal)$LdapFilterSIDs)"

                                try { WatchCatchableExitSignal } catch { }
                                $fsps = $Search.FindAll()
                                try { WatchCatchableExitSignal } catch { }

                                if ($fsps.count -gt 0) {
                                    foreach ($fsp in $fsps) {
                                        try { WatchCatchableExitSignal } catch { }

                                        if (($fsp.path -ne '') -and ($null -ne $fsp.path)) {
                                            # A Foreign Security Principal (FSP) is created in each (sub)domain in which it is granted permissions
                                            # A FSP it can only be member of a domain local group - so we set the searchroot to the (sub)domain of the Foreign Security Principal
                                            # FSPs have no tokengroups or tokengroupsglobalanduniversal attribute, which would not contain domain local groups anyhow
                                            # member:1.2.840.113556.1.4.1941:= (LDAP_MATCHING_RULE_IN_CHAIN) returns groups containing a specific DN as member, incl. nesting
                                            Write-Verbose "          Found ForeignSecurityPrincipal $($fsp.properties.cn) in $((($fsp.path -split ',DC=')[1..999] -join '.'))"

                                            if ($((($fsp.path -split ',DC=')[1..999] -join '.')) -iin @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $LookupDomainsToTrusts[$(($($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.')] }).name)) {
                                                try {
                                                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$((($fsp.path -split ',DC=')[1..999] -join '.'))")
                                                    $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($fsp.Properties.distinguishedname)))"

                                                    try { WatchCatchableExitSignal } catch { }
                                                    $fspGroups = $Search.FindAll()
                                                    try { WatchCatchableExitSignal } catch { }

                                                    if ($fspGroups.count -gt 0) {
                                                        foreach ($group in $fspgroups) {
                                                            try { WatchCatchableExitSignal } catch { }

                                                            $sid = (New-Object System.Security.Principal.SecurityIdentifier $($group.properties.objectsid), 0).value
                                                            Write-Verbose "          $($sid) (domain local group across trust)"
                                                            $GroupsSIDs += $sid

                                                            foreach ($SidHistorySid in @($group.properties.sidhistory | Where-Object { $_ })) {
                                                                $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                                                                Write-Verbose "          $($sid) (domain local group sIDHistory across trust)"
                                                                $GroupsSIDs += $sid
                                                            }
                                                        }
                                                    } else {
                                                        Write-Verbose '          FSP is not member of any group'
                                                    }
                                                } catch {
                                                    Write-Host "          Error: $($error[0].exception)" -ForegroundColor red
                                                }
                                            } else {
                                                Write-Verbose "          Ignoring, because '$($fsp.path)' is not part of a trust in TrustsToCheckForGroups."
                                            }
                                        }
                                    }
                                } else {
                                    Write-Verbose '          No ForeignSecurityPrincipal(s) found'
                                }
                            }
                        }
                    } else {
                        try {
                            try { WatchCatchableExitSignal } catch { }

                            $tempX = GraphGetUserTransitiveMemberOf $ADPropsMailboxes[$AccountNumberRunning].userPrincipalName

                            try { WatchCatchableExitSignal } catch { }

                            foreach ($sid in @($tempX.memberof.value.securityidentifier | Where-Object { $_ })) {
                                $GroupsSIDs += $sid
                                Write-Verbose "        $($sid) (Entra ID group)"
                            }

                            try { WatchCatchableExitSignal } catch { }

                            foreach ($sid in @($tempX.memberof.value.onpremisessecurityidentifier | Where-Object { $_ })) {
                                $GroupsSIDs += $sid
                                Write-Verbose "        $($sid) (on-prem group)"
                            }

                            $tempX = $null
                        } catch {
                            Write-Host '        Skipping, mailbox not found in Microsoft Graph.' -ForegroundColor yellow
                        }
                    }
                } else {
                    Write-Host '        Skipping, as mailbox could not be found in your environment in an earlier step.' -ForegroundColor yellow
                }

                $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'GroupsSIDs' -Value $GroupsSIDs -Force

                if ($ADPropsCurrentUser.proxyaddresses -icontains "smtp:$($MailAddresses[$AccountNumberRunning])") {
                    $ADPropsCurrentUser = $ADPropsMailboxes[$AccountNumberRunning]
                }
            } else {
                $ADPropsMailboxes[$AccountNumberRunning] = $null
                $ADPropsMailboxManagers[$AccountNumberRunning] = $null
            }
        } else {
            Write-Host "    Mailbox user object already found before, using cached data of $($MailAddresses[$CurrentMailboxAlreadyFoundFirstIndex])"

            $ADPropsMailboxes[$AccountNumberRunning] = $ADPropsMailboxes[$CurrentMailboxAlreadyFoundFirstIndex]
            $ADPropsMailboxManagers[$AccountNumberRunning] = $ADPropsMailboxManagers[$CurrentMailboxAlreadyFoundFirstIndex]
        }

        if ($AccountNumberRunning -eq ($MailAddresses.count - 1)) {
            if ($VirtualMailboxConfigFile) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('DefineAndAddVirtualMailboxes')))) {
                    Write-Host '  Virtual mailboxes and dynamic signature INI entries can not be defined and added.' -ForegroundColor Yellow
                    Write-Host "  The 'VirtualMailboxConfigFile' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                    Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DefineAndAddVirtualMailboxes()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '  Error defining and adding virtual mailboxes.' -ForegroundColor Yellow
                        Write-Host "  $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "  Parameter 'VirtualMailboxConfigFile' is not enabled, skipping task."
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Sort mailbox list: User's primary mailbox, mailboxes in default Outlook profile, others @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    # Get users primary mailbox
    $p = $null

    # First, check if the user has a mail attribute set
    if ($ADPropsCurrentUser.mail) {
        Write-Host "  Mail attribute of currently logged-in or simulated user: '$($ADPropsCurrentUser.mail)'"

        for ($i = 0; $i -lt $LegacyExchangeDNs.count; $i++) {
            # if (($LegacyExchangeDNs[$i]) -and (($ADPropsMailboxes[$i].proxyaddresses) -icontains "smtp:$($ADPropsCurrentUser.mail)")) {
            if ((($ADPropsMailboxes[$i].proxyaddresses) -icontains "smtp:$($ADPropsCurrentUser.mail)")) {
                if (($SimulateUser) -or ((-not $SimulateUser) -and ($LegacyExchangeDNs[$i]))) {
                    $p = $i
                    break
                }
            }
        }

        if ($p -ge 0) {
            Write-Host '    Matching mailbox found'
        } else {
            Write-Host '    No matching mailbox found, see prior warning messages for details' -ForegroundColor Yellow
        }
    } else {
        Write-Host '  AD mail attribute of currently logged-in user is empty'

        if ($null -ne $TrustsToCheckForGroups[0]) {
            Write-Host '    Searching msExchMasterAccountSid'
            # No mail attribute set, check for match(es) of user's objectSID and mailbox's msExchMasterAccountSid
            for ($i = 0; $i -lt $MailAddresses.count; $i++) {
                if ($ADPropsMailboxes[$i].msexchmasteraccountsid) {
                    try { WatchCatchableExitSignal } catch { }

                    if ((New-Object System.Security.Principal.SecurityIdentifier $ADPropsMailboxes[$i].msexchmasteraccountsid[0], 0).value -iin $CurrentUserSIDs) {
                        if ($p -ge 0) {
                            # $p already set before, there must be at least two matches, so set it to -1
                            $p = -1
                        } elseif ((-not $p) -and ($RegistryPaths[$i] -ilike '*\9375CFF0413111d3B88A00104B2A6676\*')) {
                            $p = $i
                        }
                    }
                }
            }

            if ($p -ge 0) {
                Write-Host "    One matching primary mailbox found: $MailAddresses[$i]"
            } elseif ($null -eq $p) {
                Write-Host '    No matching primary mailbox found' -ForegroundColor Yellow
            } else {
                Write-Host '    Multiple matching primary mailboxes found, no prioritization possible' -ForegroundColor Yellow
            }
        } else {
            Write-Host
        }
    }

    Write-Host '  Mailbox priority (highest to lowest)'
    $MailboxNewOrder = @()
    $PrimaryMailboxAddress = $null

    if ($p -ge 0) {
        $MailboxNewOrder += $p
        $PrimaryMailboxAddress = $MailAddresses[$p]
    }

    if ((-not $SimulateUser) -and ($OutlookProfiles.count -gt 0)) {
        foreach ($OutlookProfile in $OutlookProfiles) {
            $MailAddressesToSearch = @()
            $MailAddressesToSearchLookup = @{}
            for ($count = 0; $count -lt $RegistryPaths.count; $count++) {
                if ($MailAddresses[$count] -and ($RegistryPaths[$count] -ilike "Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\*")) {
                    $MailAddressesToSearch += $MailAddresses[$count]
                    $MailAddressesToSearchLookup[$($MailAddresses[$count])] = $MailAddresses[$count]

                    foreach ($ProxyAddress in $ADPropsMailboxes[$count].proxyaddresses) {
                        if ([string]$ProxyAddress -ilike 'smtp:*') {
                            $MailAddressesToSearch += $([string]$ProxyAddress -ireplace 'smtp:', '')
                            $MailAddressesToSearchLookup[$([string]$ProxyAddress -ireplace 'smtp:', '')] = $MailAddresses[$count]
                        }
                    }
                }
            }

            $CurrentOutlookProfileMailboxSortOrder = @()

            foreach ($RegistryFolder in @(Get-ItemProperty "hkcu:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$($OutlookProfile)\0a0d020000000000c000000000000046" -ErrorAction SilentlyContinue | Where-Object { ($_.'11020458') })) {
                try { WatchCatchableExitSignal } catch { }

                try {
                    @(@(([regex]::Matches((@(ForEach ($char in @(($RegistryFolder.'11020458' -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$($char)" }) -join ''), (@(@($MailAddressesToSearch) | ForEach-Object { [Regex]::Escape($_) }) -join '|'), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).captures.value).tolower()) | Select-Object -Unique) | ForEach-Object {
                        $CurrentOutlookProfileMailboxSortOrder += $MailAddressesToSearchLookup[$_]
                    }
                } catch {
                }
            }

            if (($CurrentOutlookProfileMailboxSortOrder.count -gt 0) -and ($CurrentOutlookProfileMailboxSortOrder.count -eq (@($RegistryPaths | Where-Object { $_ -ilike "Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\*" }).count))) {
                Write-Verbose '  Outlook mailbox display sort order is defined and contains all found mail addresses.'
                foreach ($CurrentOutlookProfileMailboxSortOrderMailbox in $CurrentOutlookProfileMailboxSortOrder) {
                    for ($i = 0; $i -le $RegistryPaths.count - 1; $i++) {
                        try { WatchCatchableExitSignal } catch { }

                        if (($RegistryPaths[$i] -ilike "Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\*") -and ($i -ne $p)) {
                            if ($MailAddresses[$i] -ieq $CurrentOutlookProfileMailboxSortOrderMailbox) {
                                $MailboxNewOrder += $i
                                break
                            }
                        }
                    }
                }
            } else {
                for ($i = 0; $i -le $RegistryPaths.count - 1; $i++) {
                    try { WatchCatchableExitSignal } catch { }

                    if (($RegistryPaths[$i] -ilike "Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\*") -and ($i -ne $p)) {
                        $MailboxNewOrder += $i
                    }
                }
            }

        }
    }

    for ($i = 0; $i -lt $MailAddresses.Count; $i++) {
        if ($MailboxNewOrder -inotcontains $i ) {
            $MailboxNewOrder += $i
        }
    }

    foreach ($VariableName in ('RegistryPaths', 'MailAddresses', 'LegacyExchangeDNs', 'ADPropsMailboxesUserDomain', 'ADPropsMailboxes')) {
        (Get-Variable -Name $VariableName).value = (Get-Variable -Name $VariableName).value[$MailboxNewOrder]
    }

    for ($x = 0; $x -lt $MailAddresses.count; $x++) {
        if ($MailAddresses.IndexOf($MailAddresses[$x]) -eq $x) {
            Write-Host "    $($MailAddresses[$x])"

            $y = 0

            @(
                foreach ($MailAddress in $MailAddresses) {
                    if ($MailAddress -ieq $MailAddresses[$x]) {
                        $y
                    }
                    $y++
                }
            ) | ForEach-Object {
                Write-Verbose "      Outlook profile '$(($RegistryPaths[$_] -split '\\')[8])'"
                Write-Verbose "        Registry: $($RegistryPaths[$_] -ireplace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_CURRENT_USER', 'HKCU')"
                Write-Verbose "        LegacyExchangeDN: $($LegacyExchangeDNs[$_])"
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    $TemplateFilesGroupSIDsOverall = @{}

    foreach ($SigOrOOF in ('signature', 'OOF')) {
        if (($SigOrOOF -eq 'OOF') -and ($SetCurrentUserOOFMessage -eq $false)) {
            break
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host
        Write-Host "Get all $($SigOrOOF) template files and categorize them @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        $TemplateFilesCommon = @{}
        $TemplateFilesGroup = @{}
        $TemplateFilesGroupFilePart = @{}
        $TemplateFilesMailbox = @{}
        $TemplateFilesMailboxFilePart = @{}
        $TemplateFilesReplacementvariable = @{}
        $TemplateFilesReplacementvariableFilePart = @{}
        $TemplateFilesDefaultnewOrInternal = @{}
        $TemplateFilesDefaultreplyfwdOrExternal = @{}
        $TemplateFilesWriteProtect = @{}

        $TemplateTemplatePath = Get-Variable -Name "$($SigOrOOF)TemplatePath" -ValueOnly
        $TemplateIniFile = Get-Variable -Name "$($SigOrOOF)IniFile" -ValueOnly
        $TemplateIniSettings = Get-Variable -Name "$($SigOrOOF)IniSettings" -ValueOnly

        # Remove trailing null character from file names being enumerated in SharePoint folders. .Net or the WebDAV client sometimes add a null character, which is not allowed in file and path names.
        ## Original code:
        ## $TemplateFiles = @((Get-ChildItem -LiteralPath $TemplateTemplatePath -File -Filter $(if ($UseHtmTemplates) { '*.htm' } else { '*.docx' })) | Sort-Object -Culture $TemplateFilesSortCulture)
        $TemplateFiles = @(@(@(@(Get-ChildItem -LiteralPath $TemplateTemplatePath -File) | Where-Object { $_.Extension -iin $(if ($UseHtmTemplates) { @('.htm', ".htm$([char]0)") } else { @('*.docx', ".docx$([char]0)") }) }) | Select-Object -Property @{n = 'FullName'; e = { $_.FullName.ToString() -ireplace '\x00$', '' } }, @{n = 'Name'; Expression = { $_.Name.ToString() -ireplace '\x00$', '' } }) | Sort-Object -Property FullName, Name -Culture $TemplateFilesSortCulture)

        if ($TemplateIniFile -ne '') {
            Write-Host "  Compare $($SigOrOOF) ini entries and file system"
            foreach ($Enumerator in $TemplateIniSettings.GetEnumerator().name) {
                try { WatchCatchableExitSignal } catch { }

                if ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>']) {
                    if (($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -ine '<Set-OutlookSignatures configuration>') -and ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -inotin $TemplateFiles.name)) {
                        Write-Host "    '$($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'])' ($($SigOrOOF) ini index #$($Enumerator)) found in ini but not in signature template path." -ForegroundColor Yellow
                    }

                    if (($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -ine '<Set-OutlookSignatures configuration>') -and ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -inotlike "*.$(if($UseHtmTemplates){'htm'}else{'docx'})")) {
                        Write-Host "    '$($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'])' ($($SigOrOOF) ini index #$($Enumerator)) has the wrong file extension ('-UseHtmTemplates true' allows .htm, else .docx)" -ForegroundColor Yellow
                    }
                }
            }

            $x = @(foreach ($Enumerator in $TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)]) { $Enumerator['<Set-OutlookSignatures template>'] })

            foreach ($TemplateFile in $TemplateFiles) {
                if ($TemplateFile.name -inotin $x) {
                    Write-Host "    '$($TemplateFile.name)' found in $($SigOrOOF) template path but not in ini." -ForegroundColor Yellow
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host '  Sort template files according to configuration'
            $TemplateFilesSortCulture = (@($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1)['SortCulture']

            # Populate template files in the most complicated way first: SortOrder 'AsInThisFile'
            # This also considers that templates can be referenced multiple times in the INI file
            # If the setting in the ini file is different, we only need to sort $TemplateFiles
            $TemplateFilesExisting = @(foreach ($Enumerator in $TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)]) { $Enumerator['<Set-OutlookSignatures template>'] })
            $TemplateFiles = @($TemplateFiles | Where-Object { $_.name -iin $TemplateFilesExisting })
            $TemplateFiles | Add-Member -MemberType NoteProperty -Name TemplateIniSettingsIndex -Value $null -Force
            $TemplateFilesSortOrder = @()
            $TemplateFilesIniIndex = @()

            if ($TemplateFiles) {
                foreach ($Enumerator in $TemplateIniSettings.GetEnumerator().name) {
                    if (@($TemplateFiles.name) -icontains $TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>']) {
                        for ($x = 0; $x -lt $TemplateFiles.count; $x++) {
                            if ($TemplateFiles[$x].name -ieq $TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>']) {
                                $TemplateFilesSortOrder += $x
                                $TemplateFilesIniIndex += $Enumerator
                            }
                        }
                    }
                }

                $TemplateFiles = @($TemplateFiles[$TemplateFilesSortOrder] | Select-Object -Property fullname, name, TemplateIniSettingsIndex)

                if ($TemplateFiles.count -gt 0) {
                    foreach ($index In 0..($TemplateFiles.Count - 1)) {
                        $TemplateFiles[$index].TemplateIniSettingsIndex = $TemplateFilesIniIndex[$index]
                    }
                }

                if (($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' })) {
                    switch ((@($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1)['SortOrder']) {
                        { $_ -iin ('AsInThisFile', 'AsListed') } {
                            # nothing to do, $TemplateFiles is already correctly populated and sorted
                            break
                        }

                        { $_ -iin ('a', 'asc', 'ascending', 'az', 'a-z', 'a..z', 'up') } {
                            $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } })
                            break
                        }

                        { $_ -iin ('d', 'des', 'desc', 'descending', 'za', 'z-a', 'z..a', 'dn', 'down') } {
                            $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } } -Descending)
                            break
                        }

                        default {
                            # same as 'AsInThisFile'
                            # nothing to do, $TemplateFiles is already correctly populated and sorted
                        }
                    }
                } else {
                    $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } })
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        foreach ($TemplateFile in $TemplateFiles) {
            try { WatchCatchableExitSignal } catch { }

            $TemplateIniSettingsIndex = $TemplateFile.TemplateIniSettingsIndex
            $TemplateFileGroupSIDs = @{}
            Write-Host ("    '$($TemplateFile.Name)' ($($SigOrOOF) ini index #$($TemplateIniSettingsIndex))")

            if ($TemplateIniSettings[$TemplateIniSettingsIndex]['<Set-OutlookSignatures template>'] -ieq $TemplateFile.name) {
                $TemplateFilePart = (@(@($TemplateIniSettings[$TemplateIniSettingsIndex].GetEnumerator().Name) | Sort-Object -Culture $TemplateFilesSortCulture) -join '] [')
                if ($TemplateFilePart) {
                    $TemplateFilePart = ($TemplateFilePart -split '\] \[' | Where-Object { $_ -inotin ('OutlookSignatureName', '<Set-OutlookSignatures template>') }) -join '] ['
                    $TemplateFilePart = '[' + $TemplateFilePart + ']'
                    $TemplateFilePart = $TemplateFilePart -ireplace '\[\]', ''
                }

                if ($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName']) {
                    Write-Host "      Outlook signature name: '$($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'])'"

                    if ((CheckFilenamePossiblyInvalid -Filename $TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'])) {
                        Write-Host "        Ignore INI entry, signature name is invalid: $((CheckFilenamePossiblyInvalid -Filename $TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName']))" -ForegroundColor Yellow

                        Continue
                    }

                    $TemplateFileTargetName = ($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'] + $(if ($UseHtmTemplates) { '.htm' } else { '.docx' }))

                } else {
                    if ((CheckFilenamePossiblyInvalid -Filename $TemplateFile.Name)) {
                        Write-Host "      Ignore INI entry, signature name is invalid: $((CheckFilenamePossiblyInvalid -Filename $TemplateFile.Name))" -ForegroundColor Yellow

                        Continue
                    }

                    $TemplateFileTargetName = $TemplateFile.Name
                }
            } else {
                $TemplateFilePart = ''
                $TemplateFileTargetName = $TemplateFile.Name
            }

            $TemplateFilePartRegexTimeAllow = '\[(?!-:)\d{12}Z?-\d{12}Z?\]'
            $TemplateFilePartRegexTimeDeny = '\[-:\d{12}Z?-\d{12}Z?\]'
            $TemplateFilePartRegexGroupAllow = '(?i)\[(?!-:|-CURRENTUSER:)\S+?(?<!]) .+?\]'
            $TemplateFilePartRegexGroupDeny = '(?i)\[(-:|-CURRENTUSER:)\S+?(?<!]) .+?\]'
            $TemplateFilePartRegexMailaddressAllow = '(?i)\[(?!-:|-CURRENTUSER:)(\S+?)@(\S+?)\.(\S+?)\]'
            $TemplateFilePartRegexMailaddressDeny = '(?i)\[(-:|-CURRENTUSER:)(\S+?)@(\S+?)\.(\S+?)\]'
            $TemplateFilePartRegexReplacementvariableAllow = '(?i)\[(?!-:)\$.*\$\]'
            $TemplateFilePartRegexReplacementvariableDeny = '(?i)\[(-:)\$.*\$\]'

            if ($SigOrOOF -ieq 'signature') {
                $TemplateFilePartRegexDefaultneworinternal = '(?i)\[DefaultNew\]'
                $TemplateFilePartRegexDefaultreplyfwdorexternal = '(?i)\[DefaultReplyFwd\]'
                $TemplateFilePartRegexWriteprotect = '(?i)\[WriteProtect\]'
            } else {
                $TemplateFilePartRegexDefaultneworinternal = '(?i)\[internal\]'
                $TemplateFilePartRegexDefaultreplyfwdorexternal = '(?i)\[external\]'
                $TemplateFilePartRegexWriteprotect = ''
            }

            $TemplateFilePartRegexKnown = '(' + (($TemplateFilePartRegexTimeAllow, $TemplateFilePartRegexTimeDeny, $TemplateFilePartRegexGroupAllow, $TemplateFilePartRegexGroupDeny, $TemplateFilePartRegexMailaddressAllow, $TemplateFilePartRegexMailaddressDeny, $TemplateFilePartRegexReplacementvariableAllow, $TemplateFilePartRegexReplacementvariableDeny, $TemplateFilePartRegexDefaultneworinternal, $TemplateFilePartRegexDefaultreplyfwdorexternal, $TemplateFilePartRegexWriteprotect) -join '|') + ')'

            try { WatchCatchableExitSignal } catch { }

            # time based template
            $TemplateFileTimeActive = $true
            if (($TemplateFilePart -imatch $TemplateFilePartRegexTimeAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexTimeDeny)) {
                Write-Host '      Time based template'
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('TimeBasedTemplate')))) {
                    Write-Host '        Templates can not be activated or deactivated for specified time ranges.' -ForegroundColor Yellow
                    Write-Host "        The 'time based template' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                    Write-Host "        Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    try { WatchCatchableExitSignal } catch { }
                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::TimeBasedTemplate()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '        Error evaluating time based templates.' -ForegroundColor Yellow
                        Write-Host "        $FeatureResult" -ForegroundColor Yellow
                    }
                }
            }

            if ($TemplateFileTimeActive -ne $true) {
                continue
            }

            try { WatchCatchableExitSignal } catch { }

            # common template
            if (($TemplateFilePart -inotmatch $TemplateFilePartRegexGroupAllow) -and ($TemplateFilePart -inotmatch $TemplateFilePartRegexMailaddressAllow) -and ($TemplateFilePart -inotmatch $TemplateFilePartRegexReplacementvariableAllow)) {
                Write-Host '      Common template (no group, email address or replacement variable allow tags specified)'
                if (-not $TemplateFilesCommon.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesCommon.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesCommon[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                }

                $TemplateClassificationDisplayOrder = ('group', 'mail', 'replacementvariable')
            } elseif ($TemplateFilePart -imatch $TemplateFilePartRegexGroupAllow) {
                $TemplateClassificationDisplayOrder = ('group', 'mail', 'replacementvariable')
            } elseif ($TemplateFilePart -imatch $TemplateFilePartRegexMailaddressAllow) {
                $TemplateClassificationDisplayOrder = ('mail', 'group', 'replacementvariable')
            } elseif ($TemplateFilePart -imatch $TemplateFilePartRegexReplacementvariableAllow) {
                $TemplateClassificationDisplayOrder = ('replacementvariable', 'group', 'mail')
            }

            foreach ($TemplateClassificationDisplayOrderEntry in $TemplateClassificationDisplayOrder) {
                try { WatchCatchableExitSignal } catch { }

                # group specific template
                if ($TemplateClassificationDisplayOrderEntry -ieq 'group') {
                    if (($TemplateFilePart -imatch $TemplateFilePartRegexGroupAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexGroupDeny)) {
                        if (-not $TemplateFilesGroup.ContainsKey($TemplateIniSettingsIndex)) {
                            $TemplateFilesGroup.add($TemplateIniSettingsIndex, @{})
                            $TemplateFilesGroup[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                        }

                        $InclusionCount = $null
                        $ExclusionCount = $null

                        foreach ($TemplateFilePartTag in @(@(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexGroupAllow).captures.value) + @([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexGroupDeny).captures.value)) | Where-Object { $_ })) {
                            if (($TemplateFilePartTag -imatch $TemplateFilePartRegexGroupAllow) -and ($null -eq $InclusionCount)) {
                                Write-Host '      Group specific template'
                                $InclusionCount++
                            } elseif (($TemplateFilePartTag -imatch $TemplateFilePartRegexGroupDeny) -and ($null -eq $ExclusionCount)) {
                                Write-Host '      Group specific exclusions'
                                $ExclusionCount++
                            }

                            Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                            $NTName = $TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|)(.*?) (.*)(\])$', '$3\$4'

                            # Check cache
                            #   $TemplateFilesGroupSIDsOverall contains tags without prefix only: [xxx xxx]
                            #   $TemplateFilesGroupSIDsOverall contains tag with extracted prefix: -:[xxx xxx]

                            if ($TemplateFilesGroupSIDsOverall.ContainsKey($($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '$1$3'))) {
                                $TemplateFileGroupSIDs.add($TemplateFilePartTag, "$($TemplateFilePartTag -ireplace '(?i)(^\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '$2')$($TemplateFilesGroupSIDsOverall[$($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '$1$3')])")
                            }

                            if ((-not $TemplateFileGroupSIDs.ContainsKey($TemplateFilePartTag))) {
                                $tempSid = ResolveToSid($NTName)

                                if ($tempSid) {
                                    $TemplateFilesGroupSIDsOverall.add($($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '$1$3'), $tempSid)
                                    $TemplateFileGroupSIDs.add($TemplateFilePartTag, "$($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '$2')$($TemplateFilesGroupSIDsOverall[$($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '$1$3')])")
                                }
                            }

                            if ($TemplateFileGroupSIDs.containskey($TemplateFilePartTag)) {
                                if ($null -ne $TemplateFileGroupSIDs[$TemplateFilePartTag]) {
                                    Write-Host "          $($TemplateFileGroupSIDs[$TemplateFilePartTag] -ireplace '(?i)^(-:|-CURRENTUSER:|CURRENTUSER:|)', '')"
                                    $TemplateFilesGroupFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesGroupFilePart[$TemplateIniSettingsIndex] + '[' + $TemplateFileGroupSIDs[$TemplateFilePartTag] + ']')
                                } else {
                                    Write-Host '          Not found' -ForegroundColor Yellow
                                }
                            } else {
                                Write-Host '          Not found' -ForegroundColor Yellow
                                $TemplateFilesGroupSIDsOverall.add($($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '$1$3'), $null)
                            }
                        }
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                # mailbox specific template
                if ($TemplateClassificationDisplayOrderEntry -ieq 'mail') {
                    if (($TemplateFilePart -imatch $TemplateFilePartRegexMailaddressAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexMailaddressDeny)) {
                        if (-not $TemplateFilesMailbox.ContainsKey($TemplateIniSettingsIndex)) {
                            $TemplateFilesMailbox.add($TemplateIniSettingsIndex, @{})
                            $TemplateFilesMailbox[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                        }

                        $InclusionCount = $null
                        $ExclusionCount = $null

                        foreach ($TemplateFilePartTag in @(@(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexMailaddressAllow).captures.value) + @([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexMailaddressDeny).captures.value)) | Where-Object { $_ })) {
                            if (($TemplateFilePartTag -imatch $TemplateFilePartRegexMailaddressAllow) -and ($null -eq $InclusionCount)) {
                                Write-Host '      Mailbox specific template'
                                $InclusionCount++
                            } elseif (($TemplateFilePartTag -imatch $TemplateFilePartRegexMailaddressDeny) -and ($null -eq $ExclusionCount)) {
                                Write-Host '      Mailbox specific exclusions'
                                $ExclusionCount++
                            }

                            Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                            $TemplateFilesMailboxFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesMailboxFilePart[$TemplateIniSettingsIndex] + $TemplateFilePartTag)
                        }
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                # Replacement variable specific template
                if ($TemplateClassificationDisplayOrderEntry -ieq 'replacementvariable') {
                    if (($TemplateFilePart -imatch $TemplateFilePartRegexReplacementvariableAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexReplacementvariableDeny)) {
                        if (-not $TemplateFilesReplacementvariable.ContainsKey($TemplateIniSettingsIndex)) {
                            $TemplateFilesReplacementvariable.add($TemplateIniSettingsIndex, @{})
                            $TemplateFilesReplacementvariable[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                        }

                        $InclusionCount = $null
                        $ExclusionCount = $null

                        foreach ($TemplateFilePartTag in @(@(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexReplacementvariableAllow).captures.value) + @([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexReplacementvariableDeny).captures.value)) | Where-Object { $_ })) {
                            if (($TemplateFilePartTag -imatch $TemplateFilePartRegexReplacementvariableAllow) -and ($null -eq $InclusionCount)) {
                                Write-Host '      Replacement variable specific template'
                                $InclusionCount++
                            } elseif (($TemplateFilePartTag -imatch $TemplateFilePartRegexReplacementvariableDeny) -and ($null -eq $ExclusionCount)) {
                                Write-Host '      Replacement variable exclusions'
                                $ExclusionCount++
                            }

                            Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                            $TemplateFilesReplacementvariableFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesReplacementvariableFilePart[$TemplateIniSettingsIndex] + $TemplateFilePartTag)
                        }
                    }
                }
            }

            try { WatchCatchableExitSignal } catch { }

            # DefaultNew, DefaultReplyFwd, Internal, External
            if ($TemplateFilePart -imatch $TemplateFilePartRegexDefaultneworinternal) {
                foreach ($TemplateFilePartTag in @(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexDefaultneworinternal).captures.value) | Where-Object { $_ })) {
                    if ($SigOrOOF -ieq 'signature') {
                        Write-Host '      Default signature for new emails'
                    } else {
                        Write-Host '      Default internal OOF message'
                    }

                    Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                }

                if (-not $TemplateFilesDefaultnewOrInternal.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesDefaultnewOrInternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultnewOrInternal[$TemplateIniSettingsIndex].add($TemplateFile.fullname, $TemplateFileTargetName)
                }
            }

            if ($TemplateFilePart -imatch $TemplateFilePartRegexDefaultreplyfwdorexternal) {
                foreach ($TemplateFilePartTag in @(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexDefaultreplyfwdorexternal).captures.value) | Where-Object { $_ })) {
                    if ($SigOrOOF -ieq 'signature') {
                        Write-Host '      Default signature for replies and forwards'
                    } else {
                        Write-Host '      Default external OOF message'
                    }

                    Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                }

                if (-not $TemplateFilesDefaultreplyfwdOrExternal.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesDefaultreplyfwdOrExternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultreplyfwdOrExternal[$TemplateIniSettingsIndex].add($TemplateFile.fullname, $TemplateFileTargetName)
                }
            }

            if ($SigOrOOF -ieq 'OOF') {
                if (($TemplateFilePart -notmatch $TemplateFilePartRegexDefaultreplyfwdorexternal) -and ($TemplateFilePart -notmatch $TemplateFilePartRegexDefaultneworinternal)) {
                    $TemplateFilesDefaultnewOrInternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultnewOrInternal[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                    Write-Host '      Default internal OOF message (neither internal nor external tag specified)'
                    $TemplateFilesDefaultreplyfwdOrExternal.add($TemplateFile.FullName, $TemplateFileTargetName)
                    Write-Host '      Default external OOF message (neither internal nor external tag specified)'
                }
            }

            try { WatchCatchableExitSignal } catch { }

            # WriteProtect
            if ($TemplateFilePart -imatch $TemplateFilePartRegexWriteprotect) {
                foreach ($TemplateFilePartTag in @(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexWriteprotect).captures.value) | Where-Object { $_ })) {
                    if ($SigOrOOF -ieq 'signature') {
                        Write-Host '      Signature will be write protected'
                        if (-not $TemplateFilesWriteProtect.containskey($TemplateIniSettingsIndex)) {
                            $TemplateFilesWriteProtect.add($TemplateIniSettingsIndex, @{})
                            $TemplateFilesWriteProtect[$TemplateIniSettingsIndex].add($TemplateFile.fullname, $TemplateFileTargetName)
                        }
                    }
                }

            }

            try { WatchCatchableExitSignal } catch { }

            # unknown tags
            $x = ($TemplateFilePart -ireplace $TemplateFilePartRegexKnown, '').trim()
            if ($x) {
                Write-Host '      Unknown tags' -ForegroundColor yellow
                Write-Host "        $(($x -ireplace '^\[', '') -ireplace '\]$', '')"
            }

            Set-Variable -Name "$($SigOrOOF)Files" -Value $TemplateFiles
            Set-Variable -Name "$($SigOrOOF)FilesCommon" -Value $TemplateFilesCommon
            Set-Variable -Name "$($SigOrOOF)FilesGroup" -Value $TemplateFilesGroup
            Set-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -Value $TemplateFilesGroupFilePart
            Set-Variable -Name "$($SigOrOOF)FilesMailbox" -Value $TemplateFilesMailbox
            Set-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -Value $TemplateFilesMailboxFilePart
            Set-Variable -Name "$($SigOrOOF)FilesReplacementvariable" -Value $TemplateFilesReplacementvariable
            Set-Variable -Name "$($SigOrOOF)FilesReplacementvariableFilePart" -Value $TemplateFilesReplacementvariableFilePart

            if ($SigOrOOF -ieq 'signature') {
                $SignatureFilesDefaultNew = $TemplateFilesDefaultnewOrInternal
                $SignatureFilesDefaultReplyFwd = $TemplateFilesDefaultreplyfwdOrExternal
                $SignatureFilesWriteProtect = $TemplateFilesWriteProtect
            } else {
                $OOFFilesInternal = $TemplateFilesDefaultnewOrInternal
                $OOFFilesExternal = $TemplateFilesDefaultreplyfwdOrExternal
            }
        }
    }


    try { WatchCatchableExitSignal } catch { }


    if ($macOSSignaturesScriptable) {
        Write-Host
        Write-Host "Create copies of Outlook for Mac signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        $SignaturePaths | ForEach-Object {
            try { WatchCatchableExitSignal } catch { }

            @(@(@"
tell application "Microsoft Outlook"
    set allSignatures to every signature

    repeat with aSignature in allSignatures
        set sigName to name of aSignature

        set sigContent to content of aSignature
        set fileName to sigName & ".htm"
        set filePath to "$($_)/" & fileName
        try
            log "  '" & fileName & "'"
            set fileRef to open for access POSIX file filePath with write permission as «class utf8»
            set eof of fileRef to 0
            write sigContent to fileRef as «class utf8»
            close access fileRef
        on error errorMessage
            log "    Error copying to '" & filepath & "': " & errorMessage
        end try

        set sigContent to plain text content of aSignature
        set fileName to sigName & ".txt"
        set filePath to "$($_)/" & fileName
        try
            log "  '" & filename & "'"
            set fileRef to open for access POSIX file filePath with write permission as «class utf8»
            set eof of fileRef to 0
            write sigContent to fileRef as «class utf8»
            close access fileRef
        on error errorMessage
            log "    Error copying to '" & filePath & "': " & errorMessage
        end try
    end repeat
end tell
"@ | osascript *>&1)) | ForEach-Object { Write-Host $_.tostring() }
        }
    }


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Start Word background process @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if (($UseHtmTemplates -eq $true) -and (($CreateRtfSignatures -eq $false))) {
        Write-Host '  Not required: UseHtmTemplates = $true, CreateRtfSignatures = $false'
    } else {
        Write-Verbose "  WordProcessPriority: '$($WordProcessPriorityText)' ('$($WordProcessPriority)')"

        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class Win32Api
{
[System.Runtime.InteropServices.DllImportAttribute( "User32.dll", EntryPoint =  "GetWindowThreadProcessId" )]
public static extern int GetWindowThreadProcessId ( [System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, out int lpdwProcessId );

[DllImport("User32.dll", CharSet = CharSet.Auto)]
public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
}
'@

        # Start Word dummy object, set process priority, start real Word object, set process priority, close dummy object - this seems to avoid a rare problem where a manually started Word instance connects to the Word process created by the software
        try {
            try { WatchCatchableExitSignal } catch { }

            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value 0 -ErrorAction SilentlyContinue

            $tempVerbosePreference = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'
            $script:COMWordDummy = New-Object -ComObject Word.Application
            $VerbosePreference = $tempVerbosePreference
            $script:COMWordDummy.Visible = $false

            # Restore original Word AlertIfNotDefault setting
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value $script:WordAlertIfNotDefaultOriginal -ErrorAction SilentlyContinue | Out-Null


            if ($script:COMWordDummy) {
                try { WatchCatchableExitSignal } catch { }

                # Set Word process priority
                $script:COMWordDummyCaption = $script:COMWordDummy.Caption
                $script:COMWordDummy.Caption = "Set-OutlookSignatures $([guid]::NewGuid())"
                $script:COMWordDummyHWND = [Win32Api]::FindWindow( 'OpusApp', $($script:COMWordDummy.Caption) )
                $script:COMWordDummyPid = [IntPtr]::Zero
                $null = [Win32Api]::GetWindowThreadProcessId( $script:COMWordDummyHWND, [ref] $script:COMWordDummyPid );
                $script:COMWordDummy.Caption = $script:COMWordDummyCaption
                $null = Get-CimInstance Win32_process -Filter "ProcessId = ""$script:COMWordDummyPid""" | Invoke-CimMethod -Name SetPriority -Arguments @{Priority = $WordProcessPriority }
            }

            try { WatchCatchableExitSignal } catch { }

            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value 0 -ErrorAction SilentlyContinue

            $tempVerbosePreference = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'
            $script:COMWord = New-Object -ComObject Word.Application
            $VerbosePreference = $tempVerbosePreference
            $script:COMWord.Visible = $false

            # Restore original Word AlertIfNotDefault setting
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value $script:WordAlertIfNotDefaultOriginal -ErrorAction SilentlyContinue | Out-Null


            if ($script:COMWord) {
                try { WatchCatchableExitSignal } catch { }

                # Set Word process priority
                $script:COMWordCaption = $script:COMWord.Caption
                $script:COMWord.Caption = "Set-OutlookSignatures $([guid]::NewGuid())"
                $script:COMWordHWND = [Win32Api]::FindWindow( 'OpusApp', $($script:COMWord.Caption) )
                $script:COMWordPid = [IntPtr]::Zero
                $null = [Win32Api]::GetWindowThreadProcessId( $script:COMWordHWND, [ref] $script:COMWordPid );
                $script:COMWord.Caption = $script:COMWordCaption
                $null = Get-CimInstance Win32_process -Filter "ProcessId = ""$script:COMWordPid""" | Invoke-CimMethod -Name SetPriority -Arguments @{Priority = $WordProcessPriority }
            }

            if ($script:COMWordDummy) {
                $script:COMWordDummy.Quit([ref]$false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWordDummy) | Out-Null
                Remove-Variable COMWordDummy -Scope 'script'
            }

            try { WatchCatchableExitSignal } catch { }

            Add-Type -Path (Get-ChildItem -LiteralPath ((Join-Path -Path ($env:SystemRoot) -ChildPath 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word')) -Filter 'Microsoft.Office.Interop.Word.dll' -Recurse | Select-Object -ExpandProperty FullName -Last 1)
        } catch {
            Write-Host $error[0]
            Write-Host '  Word not installed or not working correctly. Install or repair Word and the registry information about Word, or consider using HTM templates instead of DOCX tempates. Exit.' -ForegroundColor Red

            # Restore original Word AlertIfNotDefault setting
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value $script:WordAlertIfNotDefaultOriginal -ErrorAction SilentlyContinue | Out-Null

            $script:ExitCode = 17
            $script:ExitCodeDescription = 'Word not installed or not working correctly.'
            exit
        }
    }


    # Process each email address only once
    $script:SignatureFilesDone = @()

    if ($SimulateUser) {
        try { WatchCatchableExitSignal } catch { }

        Get-ChildItem (Join-Path -Path ($SignaturePaths[0]) -ChildPath '___Mailbox *') -Attributes Directory -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ForEach-Object {
            try { WatchCatchableExitSignal } catch { }

            RemoveItemAlternativeRecurse $($_.FullName)
        }
    }

    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        try { WatchCatchableExitSignal } catch { }

        if (($AccountNumberRunning -eq $MailAddresses.IndexOf($MailAddresses[$AccountNumberRunning])) -and ($($MailAddresses[$AccountNumberRunning]) -like '*@*')) {
            Write-Host
            Write-Host "Mailbox $($MailAddresses[$AccountNumberRunning]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

            $UserDomain = ''
            $GroupsSIDs = @()
            $ADPropsCurrentMailbox = @()
            $ADPropsCurrentMailboxManager = @()

            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                $ADPropsCurrentMailbox = $ADPropsMailboxes[$AccountNumberRunning]
                $ADPropsCurrentMailboxManager = $ADPropsMailboxManagers[$AccountNumberRunning]
                $GroupsSIDs = $ADPropsMailboxes[$AccountNumberRunning].GroupsSIDs
            }


            if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('CLCGM')))) {
                Write-Host '  Mailbox is member of license group: False (no valid Benefactor Circle license file found)'
            } else {
                try { WatchCatchableExitSignal } catch { }

                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::CLCGM()

                if ($FeatureResult -ine 'true') {
                    Write-Host "  Mailbox is member of license group: False ($($FeatureResult))"
                } else {
                    Write-Host '  Mailbox is member of license group: True'
                }
            }


            try { WatchCatchableExitSignal } catch { }


            Write-Host "  Extract SMTP addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            $CurrentMailboxSMTPAddresses = @()
            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                foreach ($ProxyAddress in $ADPropsCurrentMailbox.proxyaddresses) {
                    if ([string]$ProxyAddress -ilike 'smtp:*') {
                        $CurrentMailboxSmtpaddresses += [string]$ProxyAddress -ireplace 'smtp:', ''
                        Write-Verbose "    $($CurrentMailboxSMTPAddresses[-1])"
                    }
                }
            } else {
                $CurrentMailboxSmtpaddresses += $($MailAddresses[$AccountNumberRunning])
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox.' -ForegroundColor Yellow
                Write-Host "    Using '$($MailAddresses[$AccountNumberRunning])' as single known SMTP address." -ForegroundColor Yellow
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host "  Calculate replacement variables @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            $ReplaceHash = @{}

            if (Test-Path -Path $ReplacementVariableConfigFile -PathType Leaf) {
                try {
                    Write-Host "    '$ReplacementVariableConfigFile'"
                    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $ReplacementVariableConfigFile -Encoding UTF8 -Raw)))
                } catch {
                    Write-Host $error[0]
                    Write-Host "    Problem executing content of '$ReplacementVariableConfigFile'. Exit." -ForegroundColor Red
                    $script:ExitCode = 18
                    $script:ExitCodeDescription = 'Problem executing content of ReplacementVariableConfigFile.'
                    exit
                }
            } else {
                Write-Host "    Problem connecting or reading '$ReplacementVariableConfigFile'. Exit." -ForegroundColor Red
                $script:ExitCode = 19
                $script:ExitCodeDescription = 'Problem connecting or reading ReplacementVariableConfigFile.'
                exit
            }

            try { WatchCatchableExitSignal } catch { }

            $PictureVariablesArray = @()

            foreach ($VariableName in @(foreach ($VariableName in @(
                            @(
                                foreach ($ReplacementVariableScope in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
                                    @(1..10) | ForEach-Object { "`$$($ReplacementVariableScope)CustomImage$($_)`$" }
                                }
                            ) +
                            @('$CurrentMailboxManagerPhoto$', '$CurrentMailboxPhoto$', '$CurrentUserManagerPhoto$', '$CurrentUserPhoto$')
                        )
                    ) {
                        $VariableName
                    }
                )
            ) {
                try { WatchCatchableExitSignal } catch { }

                New-Variable -Name $($($VariableName).Trim('$') + 'Guid') -Value (New-Guid).Guid -Force

                $PictureVariablesArray += , @($VariableName, $(Get-Variable -Name $($VariableName.Trim('$') + 'Guid') -ValueOnly))
            }

            foreach ($replaceKey in @($replaceHash.Keys | Sort-Object -Culture $TemplateFilesSortCulture)) {
                try { WatchCatchableExitSignal } catch { }

                if ($replaceKey -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' })) {
                    if ($($replaceHash[$replaceKey])) {
                        Write-Verbose "    $($replaceKey): $($replaceHash[$replaceKey])"
                    }
                } else {
                    if ($null -ne $($replaceHash[$replaceKey])) {
                        Write-Verbose "    $($replaceKey): Photo available, $([math]::ceiling($($replaceHash[$replaceKey]).Length / 1KB)) KiB"
                    }
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host '    Export available images'
            foreach ($VariableName in $PictureVariablesArray) {
                try { WatchCatchableExitSignal } catch { }

                Write-Verbose "    $($VariableName[0]), $([math]::ceiling(($ReplaceHash[$VariableName[0]]).Length / 1KB)) KiB @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                if ($null -ne $($ReplaceHash[$VariableName[0]])) {
                    [System.IO.File]::WriteAllBytes($(((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))), $($ReplaceHash[$VariableName[0]]))
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host "  Download roaming signatures from Exchange Online @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

            if ($MirrorCloudSignatures -eq $true) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('RoamingSignaturesDownload')))) {
                    Write-Host '    Roaming signatures can not be downloaded from Exchange Online.' -ForegroundColor Yellow
                    Write-Host "    The 'MirrorCloudSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                    Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesDownload()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '    Error downloading roaming signatures from the cloud.' -ForegroundColor Yellow
                        Write-Host "    $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "    Parameter 'MirrorCloudSignatures' is not enabled, skipping task."
            }

            try { WatchCatchableExitSignal } catch { }


            $CurrentTemplateIsForAliasSmtp = $null

            EvaluateAndSetSignatures


            # Delete photos from file system
            foreach ($VariableName in $PictureVariablesArray) {
                try { WatchCatchableExitSignal } catch { }

                Remove-Item -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))) -Force -ErrorAction SilentlyContinue
                $ReplaceHash.Remove($VariableName[0])
                $ReplaceHash.Remove(($VariableName[0][-999..-2] -join '') + 'DELETEEMPTY$')
            }


            # Set OOF message and Outlook Web signature
            if (((($SetCurrentUserOutlookWebSignature -eq $true)) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses[$AccountNumberRunning] -ieq $PrimaryMailboxAddress)) {
                if ((-not $SimulateUser)) {
                    if (-not $script:WebServicesDllPath) {
                        try { WatchCatchableExitSignal } catch { }

                        Write-Host "  Set up environment for connection to Outlook Web @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                        $script:WebServicesDllPath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))
                        try {
                            Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\netstandard2.0\Microsoft.Exchange.WebServices.Data.dll')) -Destination $script:WebServicesDllPath -Force
                            if (-not $IsLinux) {
                                Unblock-File -LiteralPath $script:WebServicesDllPath
                            }
                        } catch {
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }

                    ConnectEWS -MailAddress $PrimaryMailboxAddress -Indent '  '

                    if (-not $script:exchService) {
                        if ($SetCurrentUserOutlookWebSignature) {
                            Write-Host '    Outlook Web signature cannot be set' -ForegroundColor Red
                            $SetCurrentUserOutlookWebSignature = $false
                        }

                        if ($SetCurrentUserOOFMessage -and (($null -ne $TrustsToCheckForGroups[0]) -and ($ADPropsCurrentMailbox.msexchrecipienttypedetails -lt 2147483648))) {
                            Write-Host '   out-of-office (OOF) message(s) cannot be set' -ForegroundColor Red
                            $SetCurrentUserOOFMessage = $false
                        }
                    }
                }

                Write-Host "  Set default signature(s) in Outlook Web @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                if ($SetCurrentUserOutlookWebSignature) {
                    if ($SimulateUser -and (-not $SimulateAndDeploy)) {
                        Write-Host '      Simulation mode enabled, skipping task.' -ForegroundColor Yellow
                    } else {
                        Write-Host "    Set default classic Outlook Web signature @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SetCurrentUserOutlookWebSignature')))) {
                            Write-Host '      Default classic Outlook Web signature can not be set.' -ForegroundColor Yellow
                            Write-Host "      The 'SetCurrentUserOutlookWebSignature' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                            Write-Host "      Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                        } else {
                            try { WatchCatchableExitSignal } catch { }

                            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SetCurrentUserOutlookWebSignature()

                            if ($FeatureResult -ne 'true') {
                                Write-Host '      Error setting current user Outlook web signature.' -ForegroundColor Yellow
                                Write-Host "      $FeatureResult" -ForegroundColor Yellow
                            }
                        }

                        Write-Host "    Set default roaming Outlook Web signature(s) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                        if ($MirrorCloudSignatures -eq $true) {
                            if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('RoamingSignaturesSetDefaults')))) {
                                Write-Host '      Default roaming Outlook Web signature(s) can not be set. This also affects New Outlook on Windows.' -ForegroundColor Yellow
                                Write-Host "      The 'MirrorCloudSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                                Write-Host "      Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                            } else {
                                try { WatchCatchableExitSignal } catch { }

                                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesSetDefaults()

                                if ($FeatureResult -ne 'true') {
                                    Write-Host '      Error setting default roaming signatures in the cloud.' -ForegroundColor Yellow
                                    Write-Host "      $FeatureResult" -ForegroundColor Yellow
                                }
                            }
                        } else {
                            Write-Host "      Parameter 'MirrorCloudSignatures' is not enabled, skipping task."
                        }
                    }
                } else {
                    Write-Host "    Parameter 'SetCurrentUserOutlookWebSignature' is not enabled, skipping task."
                }

                Write-Host "  Process out-of-office (OOF) auto replies @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                if ($SetCurrentUserOOFMessage) {
                    if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SetCurrentUserOOFMessage')))) {
                        Write-Host '    The out-of-office replies can not be set.' -ForegroundColor Yellow
                        Write-Host "    The 'SetCurrentUserOOFMessage' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                        Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                    } else {
                        try { WatchCatchableExitSignal } catch { }

                        $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SetCurrentUserOOFMessage()

                        if ($FeatureResult -ne 'true') {
                            Write-Host '    Error setting current user out-of-office message.' -ForegroundColor Yellow
                            Write-Host "    $FeatureResult" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "    Parameter 'SetCurrentUserOOFMessage' is not enabled, skipping task."
                }
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    # Close Word, as it is no longer needed
    if ($script:COMWord) {
        try {
            $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal
        } catch {
        }

        $script:COMWord.Quit([ref]$false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWord) | Out-Null
        Remove-Variable -Name 'COMWord' -Scope 'script'
    }


    try { WatchCatchableExitSignal } catch { }


    # Delete old signatures created by this script, which are no longer available in $SignatureTemplatePath
    # We check all local signatures for a specific marker in HTML code, so we don't touch user-created signatures
    Write-Host
    Write-Host "Remove old signatures created by this script, which are no longer centrally available @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($DeleteScriptCreatedSignaturesWithoutTemplate -eq $true) {
        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('DeleteScriptCreatedSignaturesWithoutTemplate')))) {
            Write-Host '  Can not delete old signatures created by Set-OutlookSignatures, which are no longer centrally available.' -ForegroundColor Yellow
            Write-Host "  The 'DeleteScriptCreatedSignaturesWithoutTemplate' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
            Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
        } else {
            try { WatchCatchableExitSignal } catch { }
            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DeleteScriptCreatedSignaturesWithoutTemplate()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error deleting script created signature which no longer have a corresponding template.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "  Parameter 'DeleteScriptCreatedSignaturesWithoutTemplate' is not enabled, skipping task."
    }


    try { WatchCatchableExitSignal } catch { }


    # Delete user-created signatures if $DeleteUserCreatedSignatures -eq $true
    Write-Host
    Write-Host "Remove user-created signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($DeleteUserCreatedSignatures -eq $true) {
        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('DeleteUserCreatedSignatures')))) {
            Write-Host '  Can not remove user-created signatures.' -ForegroundColor Yellow
            Write-Host "  The 'DeleteUserCreatedSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
            Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
        } else {
            try { WatchCatchableExitSignal } catch { }

            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DeleteUserCreatedSignatures()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error removing user-created signatures.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "  Parameter 'DeleteUserCreatedSignatures' is not enabled, skipping task."
    }

    try { WatchCatchableExitSignal } catch { }

    # Upload local signatures to Exchange Online as roaming signatures
    Write-Host
    Write-Host "Upload local signatures to Exchange Online as roaming signatures for current user @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($MirrorCloudSignatures -eq $true) {
        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('RoamingSignaturesUpload')))) {
            Write-Host '  Signature(s) can not be uploaded to Exchange Online. This affects Outlook Web and New Outlook on Windows.' -ForegroundColor Yellow
            Write-Host "  The 'MirrorCloudSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
            Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
        } else {
            try { WatchCatchableExitSignal } catch { }

            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesUpload()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error uploading roaming signatures to the cloud.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "    Parameter 'MirrorCloudSignatures' is not enabled, skipping task."
    }


    try { WatchCatchableExitSignal } catch { }


    # Prepare data for Outlook add-in
    Write-Host
    Write-Host "Prepare data for Outlook add-in @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Write-Host '  Required because Microsoft actively blocks Outlook add-ins from using roaming signatures'

    [SetOutlookSignatures.Common]::PrepareOutlookAddinDataCommon()


    try { WatchCatchableExitSignal } catch { }


    # Create/update 'My signatures, powered by Set-OutlookSignatures Benefactor Circle' email draft
    Write-Host
    Write-Host "Create 'My signatures, powered by Set-OutlookSignatures Benefactor Circle' email draft for current user @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($SignatureCollectionInDrafts -eq $true) {
        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SignatureCollectionInDrafts')))) {
            Write-Host '  Can not create email draft containing all signatures.' -ForegroundColor Yellow
            Write-Host "  The 'SignatureCollectionInDrafts' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
            Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
        } else {
            try { WatchCatchableExitSignal } catch { }

            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SignatureCollectionInDrafts()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error creating ''My signatures, powered by Set-OutlookSignatures Benefactor Circle'' email draft.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "    Parameter 'SignatureCollectionInDrafts' is not enabled, skipping task."
    }


    try { WatchCatchableExitSignal } catch { }


    # Copy signatures to additional path if $AdditionalSignaturePath is set
    Write-Host
    Write-Host "Copy signatures to AdditionalSignaturePath @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($AdditionalSignaturePath) {
        Write-Host "  '$AdditionalSignaturePath'"

        if ($SimulateUser) {
            Write-Host '    Simulation mode enabled, AdditionalSignaturePath already used as output directory'
        } else {
            if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('AdditionalSignaturePath')))) {
                Write-Host '    Can not copy signatures to additional signature path.' -ForegroundColor Yellow
                Write-Host "    The 'AdditionalSignaturePath' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
            } else {
                try { WatchCatchableExitSignal } catch { }

                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::AdditionalSignaturePath()

                if ($FeatureResult -ne 'true') {
                    Write-Host '    Error copying signatures to additional signature path.' -ForegroundColor Yellow
                    Write-Host "    $FeatureResult" -ForegroundColor Yellow
                }
            }
        }
    } else {
        Write-Host "    Parameter 'AdditionalSignaturePath' is not enabled, skipping task."
    }

    try { WatchCatchableExitSignal } catch { }

    if (
        ($script:GraphUserDummyMailbox -eq $true) -or
        ($OutlookUseNewOutlook -eq $true)
    ) {
        RemoveItemAlternativeRecurse $SignaturePaths[0] -SkipFolder
    }
}


function ResolveToSid($string) {
    try { WatchCatchableExitSignal } catch { }

    # Find the last ':', use everything right from it and remove surrounding whitespace
    $string = (($string -split ':')[-1]).trim()

    if ($string.contains('\')) {
        # is already in pre-Windows 2000 format
        $local:NTName = $string
    } elseif ($string.contains(' ')) {
        # format it in pre-Windows 2000 format
        $local:NTName = ([regex]' ').replace($string, '\', 1)
    } else {
        # Invalid
        return $null
    }

    if (($null -ne $TrustsToCheckForGroups[0]) -and ($local:NTName -inotmatch '^(AzureAD\\|EntraID\\)')) {
        try {
            try { WatchCatchableExitSignal } catch { }
            $local:x = (New-Object System.Security.Principal.NTAccount($local:NTName)).Translate([System.Security.Principal.SecurityIdentifier]).value

            if ($local:x) {
                return $local:x
            }
        } catch {
            try { WatchCatchableExitSignal } catch { }

            try {
                # No group with this sAMAccountName found. Interpreting it as a display name.

                $objTrans = New-Object -ComObject 'NameTranslate'
                $objNT = $objTrans.GetType()
                $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($local:NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (4, ($local:NTName -split '\\')[1])) # 4 = ADS_NAME_TYPE_DISPLAY

                try { WatchCatchableExitSignal } catch { }
                $local:x = $(((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                try { WatchCatchableExitSignal } catch { }

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
                Remove-Variable -Name 'objTrans'
                Remove-Variable -Name 'objNT'

                if ($local:x) {
                    return $local:x
                }
            } catch {
                try { WatchCatchableExitSignal } catch { }

                try {
                    # Let the API guess what it is

                    $objTrans = New-Object -ComObject 'NameTranslate'
                    $objNT = $objTrans.GetType()
                    $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($local:NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                    $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, ($local:NTName -split '\\')[1])) # 8 = ADS_NAME_TYPE_UNKNOWN

                    try { WatchCatchableExitSignal } catch { }
                    $local:x = $(((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                    try { WatchCatchableExitSignal } catch { }

                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
                    Remove-Variable -Name 'objTrans'
                    Remove-Variable -Name 'objNT'

                    if ($local:x) {
                        return $local:x
                    }
                } catch {
                    # Nothing found
                    return $null
                }
            }
        }
    } else {
        $tempFilterOrder = @()

        # Object ID
        if ([guid]::TryParse($local:NTName.Split('\')[1], $([ref][guid]::Empty))) {
            $tempFilterOrder += "(id eq '$($local:NTName.Split('\')[1])')"
        }

        # securityIdentifier
        try {
            $null = [System.Security.Principal.SecurityIdentifier]$($local:NTName.Split('\')[1])
            $tempFilterOrder += "(securityIdentifier eq '$($local:NTName.Split('\')[1])')"
        } catch {
            # Do nothing
        }


        if ($local:NTName -inotmatch '^(AzureAD\\|EntraID\\)') {
            if ($local:NTName.Split('\')[0] -inotlike '*.*') {
                # NetBIOS domain name pattern
                $tempFilterOrder += "((onPremisesNetBiosName eq '$($local:NTName.Split('\')[0])') and (onPremisesSamAccountName eq '$($local:NTName.Split('\')[1])'))"
                $tempFilterOrder += "((onPremisesNetBiosName eq '$($local:NTName.Split('\')[0])') and (displayName eq '$($local:NTName.Split('\')[1])'))"
            } else {
                # DNS domain name pattern
                $tempFilterOrder += "((onPremisesDomainName eq '$($local:NTName.Split('\')[0])') and (onPremisesSamAccountName eq '$($local:NTName.Split('\')[1])'))"
                $tempFilterOrder += "((onPremisesDomainName eq '$($local:NTName.Split('\')[0])') and (displayName eq '$($local:NTName.Split('\')[1])'))"
            }
        }

        # Email address pattern
        if ($local:NTName.Split('\')[1] -ilike '*@*') {
            $tempFilterOrder += "(proxyAddresses/any(x:x eq 'smtp:$($local:NTName.Split('\')[1])'))"
        }

        $tempFilterOrder += "(mailNickname eq '$($local:NTName.Split('\')[1])')"
        $tempFilterOrder += "(displayName eq '$($local:NTName.Split('\')[1])')"

        # Search Graph for groups
        ForEach ($tempFilter in $tempFilterOrder) {
            try { WatchCatchableExitSignal } catch { }

            $tempResults = (GraphFilterGroups $tempFilter)

            if (($tempResults.error -eq $false) -and ($tempResults.groups.count -eq 1) -and $($tempResults.groups[0].value)) {
                if ($($tempResults.groups[0].value.securityidentifier)) {
                    return $($tempResults.groups[0].value.securityidentifier)
                }
            }
        }

        # Search Graph for users
        ForEach ($tempFilter in $tempFilterOrder) {
            try { WatchCatchableExitSignal } catch { }

            $tempResults = (GraphFilterUsers $tempFilter)

            if (($tempResults.error -eq $false) -and ($tempResults.users.count -eq 1) -and $($tempResults.users[0].value)) {
                if ($($tempResults.users[0].value.securityidentifier)) {
                    return $($tempResults.users[0].value.securityidentifier)
                }
            }
        }

        # Nothing found
        return $null
    }
}


function GetBitness {
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'files', HelpMessage = 'Comma separated list of files to process', ValueFromPipelineByPropertyName = $true)]
        [string[]]$fullname ,
        [Parameter(Mandatory = $true, ParameterSetName = 'folders', HelpMessage = 'Comma separated list of folders to process')]
        [string[]]$folders ,
        [Parameter(Mandatory = $false, ParameterSetName = 'folders')]
        [switch]$recurse ,
        [switch]$explain ,
        [switch]$quiet ,
        [switch]$dotnetOnly
    )

    Begin {
        try { WatchCatchableExitSignal } catch { }

        [int]$MACHINE_OFFSET = 4
        [int]$PE_POINTER_OFFSET = 60

        [hashtable]$machineTypes = @{
            # Source: https://learn.microsoft.com/en-us/windows/win32/debug/pe-format#machine-types
            0x0    = 'UNKNOWN' # IMAGE_FILE_MACHINE_UNKNOWN; The content of this field is assumed to be applicable to any machine type
            0x14c  = 'x86' # IMAGE_FILE_MACHINE_I386; Intel 386 or later processors and compatible processors
            0x166  = 'R4000' # IMAGE_FILE_MACHINE_R4000; MIPS little endian
            0x169  = 'WCEMIPSV2' # IMAGE_FILE_MACHINE_WCEMIPSV2; MIPS little-endian WCE v2
            0x1a2  = 'SH3' # IMAGE_FILE_MACHINE_SH3; Hitachi SH3
            0x1a3  = 'SH3DSP' # IMAGE_FILE_MACHINE_SH3DSP; Hitachi SH3 DSP
            0x1a6  = 'SH4' # IMAGE_FILE_MACHINE_SH4; Hitachi SH4
            0x1a8  = 'SH5' # IMAGE_FILE_MACHINE_SH5; Hitachi SH5
            0x1c0  = 'ARM' # IMAGE_FILE_MACHINE_ARM; ARM little endian
            0x1c2  = 'THUMB' # IMAGE_FILE_MACHINE_THUMB; Thumb
            0x1c4  = 'ARMNT' # IMAGE_FILE_MACHINE_ARMNT; ARM Thumb-2 little endian
            0x1d3  = 'AM33' # IMAGE_FILE_MACHINE_AM33; Matsushita AM33
            0x1f0  = 'POWERPC' # IMAGE_FILE_MACHINE_POWERPC; Power PC little endian
            0x1f1  = 'POWERPCFP' # IMAGE_FILE_MACHINE_POWERPCFP; Power PC with floating point support
            0x200  = 'IA64' # IMAGE_FILE_MACHINE_IA64; Intel Itanium processor family
            0x266  = 'MIPS16' # IMAGE_FILE_MACHINE_MIPS16; MIPS16
            0x366  = 'MIPSFPU' # IMAGE_FILE_MACHINE_MIPSFPU; MIPS with FPU
            0x466  = 'MIPSFPU16' # IMAGE_FILE_MACHINE_MIPSFPU16; MIPS16 with FPU
            0x5032 = 'RISCV32' # IMAGE_FILE_MACHINE_RISCV32; RISC-V 32-bit address space
            0x5064 = 'RISCV64' # IMAGE_FILE_MACHINE_RISCV64; RISC-V 64-bit address space
            0x5128 = 'RISCV128' # IMAGE_FILE_MACHINE_RISCV128; RISC-V 128-bit address space
            0x6232 = 'LOONGARCH32' # IMAGE_FILE_MACHINE_LOONGARCH32; LoongArch 32-bit processor family
            0x6264 = 'LOONGARCH64' # IMAGE_FILE_MACHINE_LOONGARCH64; LoongArch 64-bit processor family
            0x8664 = 'x64' # IMAGE_FILE_MACHINE_AMD64; x64
            0x9041 = 'M32R' # IMAGE_FILE_MACHINE_M32R; Mitsubishi M32R little endian
            0xaa64 = 'ARM64' # IMAGE_FILE_MACHINE_ARM64; ARM64 little endian
            0xebc  = 'EBC' # IMAGE_FILE_MACHINE_EBC; EFI byte code
        }

        [hashtable]$processorAchitectures = @{
            'None'  = 'None'
            'MSIL'  = 'AnyCPU'
            'X86'   = 'x86'
            'I386'  = 'x86'
            'IA64'  = 'Itanium'
            'Amd64' = 'x64'
            'Arm'   = 'ARM'
        }

        [hashtable]$pekindsExplanations = @{
            'ILOnly'                      = 'MSIL processor neutral'
            'NotAPortableExecutableImage' = 'Not in portable executable (PE) file format'
            'PE32Plus'                    = 'Requires a 64-bit platform'
            'Preferred32Bit'              = 'Platform-agnostic but should be run on 32-bit platform'
            'Required32Bit'               = 'Runs on a 32-bit platform or in the 32-bit WOW environment on a 64-bit platform'
            'Unmanaged32Bit'              = 'Contains pure unmanaged code'
        }

        If ($PSBoundParameters[ 'folders' ]) {
            $fullname = @(ForEach ($folder in $folders) {
                    Get-ChildItem -Path $folder -File -Recurse:$recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
                })
        }
    }

    Process {
        ForEach ($file in $fullname) {
            Try {
                try { WatchCatchableExitSignal } catch { }
                $runtimeAssembly = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($file)
            } Catch {
                $runtimeAssembly = $null
            }

            Try {
                try { WatchCatchableExitSignal } catch { }
                $assembly = [System.Reflection.AssemblyName]::GetAssemblyName($file)
            } Catch {
                $assembly = $null
            }

            if ((-not $dotnetOnly) -or ($assembly -and $runtimeAssembly)) {
                $data = New-Object System.Byte[] 4096

                Try {
                    $stream = New-Object System.IO.FileStream -ArgumentList $file, Open, Read
                } Catch {
                    $stream = $null

                    if (-not $quiet) {
                        Write-Verbose $_
                    }
                }

                If ($stream) {
                    try { WatchCatchableExitSignal } catch { }

                    [uint16]$machineUint = 0xffff
                    [int]$read = $stream.Read($data , 0 , $data.Count)

                    If ($read -gt $PE_POINTER_OFFSET) {
                        If (($data[0] -eq 0x4d) -and ($data[1] -eq 0x5a)) {
                            ## MZ
                            [int]$PE_HEADER_ADDR = [System.BitConverter]::ToInt32($data, $PE_POINTER_OFFSET)
                            [int]$typeOffset = $PE_HEADER_ADDR + $MACHINE_OFFSET
                            If ($data[$PE_HEADER_ADDR] -eq 0x50 -and $data[$PE_HEADER_ADDR + 1] -eq 0x45) {
                                ## PE
                                If ($read -gt $typeOffset + [System.Runtime.InteropServices.Marshal]::SizeOf($machineUint)) {
                                    [uint16]$machineUint = [System.BitConverter]::ToUInt16($data, $typeOffset)
                                    $versionInfo = Get-ItemProperty -Path $file -ErrorAction SilentlyContinue | Select-Object -ExpandProperty VersionInfo
                                    If ($runtimeAssembly -and ($module = ($runtimeAssembly.GetModules() | Select-Object -First 1))) {
                                        $pekinds = New-Object -TypeName System.Reflection.PortableExecutableKinds
                                        $imageFileMachine = New-Object -TypeName System.Reflection.ImageFileMachine
                                        $module.GetPEKind([ref]$pekinds, [ref]$imageFileMachine)
                                    } Else {
                                        $pekinds = $null
                                        $imageFileMachine = $null
                                    }

                                    try { WatchCatchableExitSignal } catch { }

                                    [pscustomobject][ordered]@{
                                        'File'                = $file
                                        'Architecture'        = $machineTypes[[int]$machineUint]
                                        'NET Architecture'    = $(If ($assembly) { $processorAchitectures[$assembly.ProcessorArchitecture.ToString()] } else { 'Not .NET' })
                                        'NET PE Kind'         = $(If ($pekinds) { if ($explain) { ($pekinds.ToString() -split ',\s?' | ForEach-Object { $pekindsExplanations[$_] }) -join ',' } else { $pekinds.ToString() } }  else { 'Not .NET' })
                                        'NET Platform'        = $(If ($imageFileMachine) { $processorAchitectures[ $imageFileMachine.ToString() ] } else { 'Not .NET' })
                                        'NET Runtime Version' = $(If ($runtimeAssembly) { $runtimeAssembly.ImageRuntimeVersion } else { 'Not .NET' })
                                        'Company'             = $versionInfo | Select-Object -ExpandProperty CompanyName
                                        'File Version'        = $versionInfo | Select-Object -ExpandProperty FileVersionRaw
                                        'Product Name'        = $versionInfo | Select-Object -ExpandProperty ProductName
                                    }
                                } Else {
                                    Write-Verbose "Only read $($data.Count) bytes from '$file' so can't read header at offset $typeOffset"
                                }
                            } ElseIf (-not $quiet) {
                                Write-Verbose "'$file' does not have a PE header signature"
                            }
                        } ElseIf (-not $quiet) {
                            Write-Verbose "'$file' is not an executable"
                        }
                    } ElseIf (-not $quiet) {
                        Write-Verbose "Only read $read bytes from '$file', not enough to get header at $PE_POINTER_OFFSET"
                    }
                    $stream.Close()
                    $stream = $null
                }
            }
        }
    }
}


function EvaluateAndSetSignatures {
    Param(
        [switch]$ProcessOOF = $false
    )

    if ($ProcessOOF -eq $true) {
        $SigOrOOF = 'OOF'
        $Indent = '  '
    } else {
        $SigOrOOF = 'Signature'
        $Indent = ''
    }

    foreach ($TemplateGroup in ('common', 'group', 'mailbox', 'replacementvariable')) {
        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent  Process $TemplateGroup $(if($TemplateGroup -iin ('group', 'mailbox', 'replacementvariable')){'specific '})templates @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        if (-not (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly -ErrorAction SilentlyContinue)) {
            continue
        }

        for ($TemplateFileIndex = 0; $TemplateFileIndex -lt (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly).count; $TemplateFileIndex++) {
            try { WatchCatchableExitSignal } catch { }

            $TemplateFile = (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly)[$TemplateFileIndex]
            $TemplateIniSettingsIndex = $TemplateFile.TemplateIniSettingsIndex

            if (-not $TemplateIniSettingsIndex) {
                continue
            }

            if (-not (Get-Variable -Name "$($SigOrOOF)Files$($TemplateGroup)" -ValueOnly).containskey($TemplateIniSettingsIndex)) {
                continue
            } else {
                $Template = (Get-Variable -Name "$($SigOrOOF)Files$($TemplateGroup)" -ValueOnly)[$TemplateIniSettingsIndex].GetEnumerator() | Select-Object -First 1
            }

            Write-Host "$Indent    '$([System.IO.Path]::GetFileName($Template.key))' ($($SigOrOOF) ini index #$($TemplateIniSettingsIndex)) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            Write-Host "$Indent      Check permissions"

            $TemplateAllowed = $false
            $CurrentTemplateIsForAliasSmtp = $null

            try { WatchCatchableExitSignal } catch { }

            # check for allow entries
            Write-Host "$Indent        Allows"
            if ($TemplateGroup -ieq 'common') {
                $TemplateAllowed = $true
                Write-Host "$Indent          Common: Template is classified as common template valid for all mailboxes"
            } elseif ($TemplateGroup -ieq 'group') {
                try { WatchCatchableExitSignal } catch { }

                $tempAllowCount = 0

                foreach ($GroupsSid in $GroupsSIDs) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($GroupsSid)``]*") {
                        $TemplateAllowed = $true
                        $tempAllowCount++
                        Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', '$1') -join '|') = $($GroupsSid) (current mailbox)"
                        break
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          No group match for current mailbox, checking current user specific allows"

                    foreach ($GroupsSid in $ADPropsCurrentUser.GroupsSIDs) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[CURRENTUSER:$($GroupsSid)``]*") {
                            $TemplateAllowed = $true
                            $tempAllowCount++
                            Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', 'CURRENTUSER:$1') -join '|') = $($GroupsSid) (current user)"
                            break
                        }
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          Group: Mailbox and current user are not member of any allowed group"
                }
            } elseif ($TemplateGroup -ieq 'mailbox') {
                try { WatchCatchableExitSignal } catch { }

                $tempAllowCount = 0

                foreach ($CurrentMailboxSmtpaddress in $CurrentMailboxSmtpAddresses) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($CurrentMailboxSmtpAddress)``]*") {
                        $TemplateAllowed = $true
                        $tempAllowCount++
                        $CurrentTemplateIsForAliasSmtp = $CurrentMailboxSmtpaddress
                        Write-Host "$Indent          First email address match: $($CurrentMailboxSmtpAddress) (current mailbox)"
                        break
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          No email address match for current mailbox, checking current user specific allows"

                    try { WatchCatchableExitSignal } catch { }

                    foreach ($CurrentUserSmtpaddress in $ADPropsCurrentUser.proxyaddresses) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[CURRENTUSER:$($CurrentUserSmtpAddress -ireplace '^smtp:', '')``]*") {
                            $TemplateAllowed = $true
                            $tempAllowCount++
                            $CurrentTemplateIsForAliasSmtp = $CurrentUserSmtpaddress
                            Write-Host "$Indent          First email address match: $($CurrentUserSmtpAddress -ireplace '^smtp:', '') (current user)"
                            break
                        }
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          Email address: Mailbox and current user do not have any allowed email address"
                }
            } elseif ($TemplateGroup -ieq 'replacementvariable') {
                try { WatchCatchableExitSignal } catch { }

                $tempAllowCount = 0

                foreach ($replaceKey in @($replaceHash.Keys | Sort-Object -Culture $TemplateFilesSortCulture)) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesReplacementvariableFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($replaceKey)``]*") {
                        if ([bool]($ReplaceHash[$replaceKey])) {
                            $TemplateAllowed = $true
                            $tempAllowCount++
                            Write-Host "$Indent          First replacement variable match: $($replaceKey) evaluates to true"
                            break
                        }
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          Replacement variable: No allowed replacement variable evaluates to true"
                }
            }

            try { WatchCatchableExitSignal } catch { }

            # check for deny entries
            if ($TemplateAllowed -eq $true) {
                Write-Host "$Indent        Denies"
                # check for group deny
                $tempDenyCount = 0

                foreach ($GroupsSid in $GroupsSIDs) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-:$($GroupsSid)``]*") {
                        $TemplateAllowed = $false
                        $tempDenyCount++
                        Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', '-:$1') -join '|') = $($GroupsSid) (current mailbox)"
                        break
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          No group match for current mailbox, checking current user specific denies"

                    foreach ($GroupsSid in $ADPropsCurrentUser.GroupsSIDs) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-CURRENTUSER:$($GroupsSid)``]*") {
                            $TemplateAllowed = $false
                            $tempDenyCount++
                            Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', '-CURRENTUSER:$1') -join '|') = $($GroupsSid) (current user)"
                            break
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Group: Mailbox and current user are not member of any denied group"
                }

                try { WatchCatchableExitSignal } catch { }

                # check for mail address deny
                $tempDenyCount = 0

                foreach ($CurrentMailboxSmtpaddress in $CurrentMailboxSmtpAddresses) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-:$($CurrentMailboxSmtpAddress)``]*") {
                        $TemplateAllowed = $false
                        $tempDenyCount++
                        Write-Host "$Indent          First email address match: $($CurrentMailboxSmtpAddress) (current mailbox)"
                        break
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          No email address match for current mailbox, checking current user specific denies"

                    foreach ($CurrentUserSmtpaddress in $ADPropsCurrentUser.proxyaddresses) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-CURRENTUSER:$($CurrentUserSmtpAddress -ireplace '^smtp:', '')``]*") {
                            $TemplateAllowed = $false
                            $tempDenyCount++
                            Write-Host "$Indent          First email address match: $($CurrentUserSmtpAddress -ireplace '^smtp:', '') (current user)"
                            break
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Email address: Mailbox and current user do not have any denied email address"
                }

                try { WatchCatchableExitSignal } catch { }

                # check for replacement variable deny
                $tempDenyCount = 0

                foreach ($replaceKey in @($replaceHash.Keys | Sort-Object -Culture $TemplateFilesSortCulture)) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesReplacementvariableFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-:$($replaceKey)``]*") {
                        if ([bool]($ReplaceHash[$replaceKey])) {
                            $TemplateAllowed = $false
                            $tempDenyCount++
                            Write-Host "$Indent          First replacement variable match: $($replaceKey) evaluates to true"
                            break
                        }
                    }
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Replacement variable: No deny replacement variable evaluates to true"
                }
            }

            try { WatchCatchableExitSignal } catch { }

            # result
            if ($Template -and ($TemplateAllowed -eq $true)) {
                Write-Host "$Indent        Use template as there is at least one allow and no deny"
                if ($ProcessOOF) {
                    if ($OOFFilesInternal.contains($TemplateIniSettingsIndex)) {
                        $OOFInternal = $Template
                        $script:OOFInternalValueBasename = $(($OOFInternal.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }

                    if ($OOFFilesExternal.contains($TemplateIniSettingsIndex)) {
                        $OOFExternal = $Template
                        $script:OOFExternalValueBasename = $(($OOFExternal.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }
                } else {
                    $Signature = $Template

                    try { WatchCatchableExitSignal } catch { }

                    SetSignatures -ProcessOOF:$ProcessOOF
                }
            } else {
                Write-Host "$Indent        Do not use template as there is no allow or at least one deny"
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ($ProcessOOF) {
        # Internal OOF message
        if ($OOFInternal -or $OOFExternal) {
            Write-Host "$Indent  Convert final OOF templates to HTM format @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        }

        if ($OOFInternal) {
            $Signature = $OOFInternal

            if ($OOFExternal -eq $OOFInternal) {
                Write-Host "$Indent    Common OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            } else {
                Write-Host "$Indent    Internal OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            }

            if ($UseHtmTemplates) {
                $Signature.value = "$OOFInternalGUID OOFInternal.htm"
            } else {
                $Signature.value = "$OOFInternalGUID OOFInternal.docx"
            }

            try { WatchCatchableExitSignal } catch { }

            SetSignatures -ProcessOOF:$ProcessOOF

            try { WatchCatchableExitSignal } catch { }

            if ($OOFExternal -eq $OOFInternal) {
                Copy-Item -Path (Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm") -Destination (Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    # External OOF message
    if ($OOFExternal -and ($OOFExternal -ne $OOFInternal)) {
        $Signature = $OOFExternal

        Write-Host "$Indent    External OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        if ($UseHtmTemplates) {
            $Signature.value = "$OOFExternalGUID OOFExternal.htm"
        } else {
            $Signature.value = "$OOFExternalGUID OOFExternal.docx"
        }

        try { WatchCatchableExitSignal } catch { }

        SetSignatures -ProcessOOF:$ProcessOOF
    }

    try { WatchCatchableExitSignal } catch { }
}


function SetSignatures {
    Param(
        [switch]$ProcessOOF = $false
    )

    try { WatchCatchableExitSignal } catch { }

    if ($ProcessOOF) {
        $Indent = '  '
    }

    if (-not $ProcessOOF) {
        Write-Host "      Outlook signature name: '$([System.IO.Path]::ChangeExtension($($Signature.value), $null) -ireplace '\.$')'"
    }

    if (-not $ProcessOOF) {
        if ($MailboxSpecificSignatureNames) {
            $SignatureFileAlreadyDone = $false
        } else {
            $SignatureFileAlreadyDone = ($script:SignatureFilesDone -contains $TemplateIniSettingsIndex)

            if ($SignatureFileAlreadyDone) {
                Write-Host "$Indent      $($SigOrOOF) ini index #$($TemplateIniSettingsIndex) already processed before with higher priority mailbox"
                Write-Host "$Indent        Not overwriting signature. Consider using parameter MailboxSpecificSignatureNames."
            } else {
                $script:SignatureFilesDone += $TemplateIniSettingsIndex
            }
        }
    }

    if (($SignatureFileAlreadyDone -eq $false) -or $ProcessOOF) {
        Write-Host "$Indent      Create temporary file copy"

        $pathGUID = (New-Guid).guid
        $path = Join-Path -Path $script:tempDir -ChildPath "$($pathGUID).htm"

        $pathConnectedFolderNames = @()

        foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
            $pathConnectedFolderNames += "$($pathGUID)$($ConnectedFilesFolderName)"
            $pathConnectedFolderNames += [uri]::EscapeDataString($pathConnectedFolderNames[-1])
            $pathConnectedFolderNames += [uri]::EscapeUriString($pathConnectedFolderNames[-2])

            $pathConnectedFolderNames += "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$($ConnectedFilesFolderName)"
            $pathConnectedFolderNames += [uri]::EscapeDataString($pathConnectedFolderNames[-1])
            $pathConnectedFolderNames += [uri]::EscapeUriString($pathConnectedFolderNames[-2])
        }

        $pathConnectedFolderNames = $pathConnectedFolderNames | Select-Object -Unique

        try { WatchCatchableExitSignal } catch { }

        if ($UseHtmTemplates) {
            try {
                if ($script:SpoDownloadUrls -and $script:SpoDownloadUrls["$($Signature.name)"]) {
                    $(New-Object Net.WebClient).DownloadFile(
                        $script:SpoDownloadUrls["$($Signature.name)"],
                        $path
                    )
                } else {
                    Copy-Item -LiteralPath $Signature.name -Destination $path -Force
                }

                try { WatchCatchableExitSignal } catch { }

                foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                    try { WatchCatchableExitSignal } catch { }

                    $pathTemp = (Join-Path -Path (Split-Path $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName")

                    if (Test-Path $pathTemp) {
                        if ($script:SpoDownloadUrls) {
                            # Work around a bug in WebDAV or .Net (https://github.com/dotnet/runtime/issues/49803)
                            #   Do not use 'Get-ChildItem'
                            $tempFiles = @()

                            [System.IO.Directory]::EnumerateFiles((Join-Path -Path (Split-Path $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName"), '*', [System.IO.SearchOption]::AllDirectories) | ForEach-Object {
                                $tempX = $_ -replace $([char]0)

                                if (
                                    $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar)."))$") -or
                                    $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar).$([IO.Path]::DirectorySeparatorChar)"))") -or
                                    $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar).."))$") -or
                                    $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar)..$([IO.Path]::DirectorySeparatorChar)"))")
                                ) {
                                    # do nothing
                                } else {
                                    $tempFiles += $tempX
                                }
                            }

                            $tempFiles = $tempFiles | Select-Object -Unique

                            foreach ($tempX in $tempFiles) {
                                if ($script:SpoDownloadUrls -and $script:SpoDownloadUrls["$($tempX)"]) {
                                    try { WatchCatchableExitSignal } catch { }

                                    $(New-Object Net.WebClient).DownloadFile(
                                        $script:SpoDownloadUrls["$($tempX)"],
                                        $tempX
                                    )
                                }
                            }
                        }

                        try { WatchCatchableExitSignal } catch { }


                        # Work around a bug in WebDAV or .Net (https://github.com/dotnet/runtime/issues/49803)
                        #   Do not use 'Get-ChildItem'
                        $tempFiles = @()

                        [System.IO.Directory]::EnumerateFiles((Join-Path -Path (Split-Path $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName"), '*', [System.IO.SearchOption]::AllDirectories) | ForEach-Object {
                            $tempX = $_ -replace $([char]0)

                            if (
                                $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar)."))$") -or
                                $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar).$([IO.Path]::DirectorySeparatorChar)"))") -or
                                $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar).."))$") -or
                                $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar)..$([IO.Path]::DirectorySeparatorChar)"))")
                            ) {
                                # do nothing
                            } else {
                                $tempFiles += $tempX
                            }
                        }

                        $tempFiles = $tempFiles | Select-Object -Unique

                        foreach ($tempX in $tempFiles) {
                            $tempY = (Join-Path -Path (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files") -ChildPath ($tempX -ireplace "^$([regex]::escape("$(Join-Path -Path (Split-Path $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName")$([IO.Path]::DirectorySeparatorChar)"))", ''))

                            $(Split-Path -LiteralPath $tempY) | ForEach-Object {
                                if (-not (Test-Path -LiteralPath $_ -PathType Container)) {
                                    $null = New-Item -ItemType Directory -Path $_
                                }
                            }

                            Copy-Item -LiteralPath $tempX -Destination $tempY -Force
                        }

                        break
                    }
                }
            } catch {
                Write-Host "$Indent        Error copying file. Skip template." -ForegroundColor Red
                Write-Host $error[0]
                continue
            }
        } else {
            $path = $([System.IO.Path]::ChangeExtension($($path), '.docx'))

            try {
                try { WatchCatchableExitSignal } catch { }

                if ($script:SpoDownloadUrls -and $script:SpoDownloadUrls["$($Signature.name)"]) {
                    $(New-Object Net.WebClient).DownloadFile(
                        $script:SpoDownloadUrls["$($Signature.name)"],
                        $path
                    )
                } else {
                    Copy-Item -LiteralPath $Signature.name -Destination $path -Force
                }
            } catch {
                Write-Host "$Indent        Error copying file. Skip template." -ForegroundColor Red
                continue
            }
        }

        try { WatchCatchableExitSignal } catch { }

        $Signature.value = $([System.IO.Path]::ChangeExtension($($Signature.value), '.htm'))

        if ($MailboxSpecificSignatureNames -and ($ProcessOOF -eq $false)) {
            if ($OutlookDisableRoamingSignatures -eq 0) {
                $Signature.value = ($Signature.Value -ireplace '\.htm$', " ($($MailAddresses[$AccountNumberRunning])).htm")
            } else {
                $Signature.value = ($Signature.Value -ireplace '\.htm$', " ($($MailAddresses[$AccountNumberRunning])).htm")
            }
        }

        if (-not $ProcessOOF) {
            $script:SignatureFilesDone += $Signature.Value
        }

        try { WatchCatchableExitSignal } catch { }

        if ($UseHtmTemplates) {
            Write-Host "$Indent      Replace picture variables"

            $AngleSharpConfig = [AngleSharp.Configuration]::Default
            $AngleSharpBrowsingContext = [AngleSharp.BrowsingContext]::New($AngleSharpConfig)
            $AngleSharpHtmlParser = $AngleSharpBrowsingContext.GetType().GetMethod('GetService').MakeGenericMethod([AngleSharp.Html.Parser.IHtmlParser]).Invoke($AngleSharpBrowsingContext, $null)
            $AngleSharpParsedDocument = $AngleSharpHtmlParser.ParseDocument("$(Get-Content -LiteralPath $path -Encoding UTF8 -Raw)")

            foreach ($image in @($AngleSharpParsedDocument.images)) {
                try { WatchCatchableExitSignal } catch { }

                $tempImageIsDeleted = $false

                if (($image.attributes['src'].value -ilike '*$*$*') -or ($image.attributes['alt'].value -ilike '*$*$*')) {
                    # Mailbox photos
                    foreach ($VariableName in $PictureVariablesArray) {
                        try { WatchCatchableExitSignal } catch { }

                        $tempImageVariableString = $Variablename[0] -ireplace '\$$', 'DELETEEMPTY$'

                        if (($image.attributes['src'].value -ilike "*$($VariableName[0])*") -or ($image.attributes['alt'].value -ilike "*$($VariableName[0])*")) {
                            if ($($ReplaceHash[$VariableName[0]])) {
                                if ($EmbedImagesInHtml -eq $false) {
                                    Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode(($image.attributes['src'].value -ireplace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                    Copy-Item (Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')) (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$($VariableName[0]).jpeg") -Force
                                    $image.attributes['src'].value = [System.Net.WebUtility]::UrlDecode("$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$($VariableName[0]).jpeg")

                                    if ($image.attributes['alt'].value) {
                                        $image.attributes['alt'].value = $($image.attributes['alt'].value) -ireplace [Regex]::Escape($VariableName[0]), ''
                                    }
                                } else {
                                    $image.attributes['src'].value = ('data:image/jpeg;base64,' + [Convert]::ToBase64String([System.IO.File]::ReadAllBytes(((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg'))))))
                                }
                            } else {
                                $image.attributes['src'].value = "$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode(($image.attributes['src'].value -ireplace '^about:', '')))))"
                            }
                        } elseif (($image.attributes['src'].value -ilike "*$($tempImageVariableString)*") -or ($image.attributes['alt'].value -ilike "*$($tempImageVariableString)*")) {
                            if ($($ReplaceHash[$VariableName[0]])) {
                                if ($EmbedImagesInHtml -eq $false) {
                                    Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode(($image.attributes['src'].value -ireplace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                    Copy-Item (Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')) (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$($VariableName[0]).jpeg") -Force
                                    $image.attributes['src'].value = [System.Net.WebUtility]::UrlDecode("$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$($VariableName[0]).jpeg")

                                    if ($image.attributes['alt'].value) {
                                        $image.attributes['alt'].value = $($image.attributes['alt'].value) -ireplace [Regex]::Escape($tempImageVariableString), ''
                                    }
                                } else {
                                    $image.attributes['src'].value = ('data:image/jpeg;base64,' + [Convert]::ToBase64String([System.IO.File]::ReadAllBytes(((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg'))))))
                                }
                            } else {
                                Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode(($image.attributes['src'].value -ireplace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                $image.Remove() | Out-Null
                                $tempImageIsDeleted = $true
                                break
                            }
                        }

                        if ((-not $tempImageIsDeleted) -and ($image.attributes['alt'].value)) {
                            $image.attributes['alt'].value = $($image.attributes['alt'].value) -ireplace [Regex]::Escape($VariableName[0]), ''
                            $image.attributes['alt'].value = $($image.attributes['alt'].value) -ireplace [Regex]::Escape($tempImageVariableString), ''
                        }
                    }

                    if ($tempImageIsDeleted) {
                        continue
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                # Other images
                if (($image.attributes['src'].value -ilike '*$*DELETEEMPTY$*') -or ($image.attributes['alt'].value -ilike '*$*DELETEEMPTY$*')) {
                    foreach ($VariableName in @(@($ReplaceHash.Keys) | Where-Object { $_ -inotin @('$CurrentMailboxPhoto$', '$CurrentMailboxManagerPhoto$', '$CurrentUserPhoto$', '$CurrentUserManagerPhoto$') })) {
                        try { WatchCatchableExitSignal } catch { }

                        $tempImageVariableString = $Variablename -ireplace '\$$', 'DELETEEMPTY$'

                        if (($image.attributes['src'].value -ilike "*$($tempImageVariableString)*") -or ($image.attributes['alt'].value -ilike "*$($tempImageVariableString)*")) {
                            if ($($ReplaceHash[$VariableName])) {
                                if ($image.attributes['alt'].value) {
                                    $image.attributes['alt'].value = $($image.attributes['alt'].value) -ireplace [Regex]::Escape($tempImageVariableString), ''
                                }
                            } else {
                                Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode(($image.attributes['src'].value -ireplace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                $image.remove() | Out-Null
                                $tempImageIsDeleted = $true
                                break
                            }
                        }
                    }

                    if ($tempImageIsDeleted) {
                        continue
                    }
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host "$Indent      Replace non-picture variables"
            $tempFileContent = $AngleSharpParsedDocument.documentelement.outerhtml

            foreach ($replaceKey in @($replaceHash.Keys | Where-Object { $_ -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' }) } | Sort-Object -Culture $TemplateFilesSortCulture)) {
                $tempFileContent = $tempFileContent -ireplace [Regex]::Escape($replacekey), $replaceHash.$replaceKey
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host "$Indent      Export to HTM format"
            [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $tempFileContent)
        } else {
            try { WatchCatchableExitSignal } catch { }

            $script:COMWord.Documents.Open($path, $false, $false, $false) | Out-Null

            try { WatchCatchableExitSignal } catch { }

            Write-Host "$Indent      Replace picture variables"
            if ($script:COMWord.ActiveDocument.Shapes.Count -gt 0) {
                Write-Host "$Indent        Warning: Template contains $($script:COMWord.ActiveDocument.Shapes.Count) image(s) configured as non-inline shapes." -ForegroundColor Yellow
                Write-Host "$Indent        Set the text wrapping to 'inline with text' to avoid incorrect positioning and other problems." -ForegroundColor Yellow
            }

            try {
                foreach ($image in @($script:COMWord.ActiveDocument.Shapes + $script:COMWord.ActiveDocument.InlineShapes)) {
                    try { WatchCatchableExitSignal } catch { }

                    # Setting the values in word is very slow, so we use temporay variables
                    $tempImageIsDeleted = $false
                    $tempImageSourceFullname = $image.linkformat.sourcefullname
                    $tempImageAlternativeText = $image.alternativetext
                    $tempImageHyperlinkAddress = $image.hyperlink.Address
                    $tempImageHyperlinkSubAddress = $image.hyperlink.SubAddress
                    $tempImageHyperlinkEmailSubject = $image.hyperlink.EmailSubject
                    $tempImageHyperlinkScreenTip = $image.hyperlink.ScreenTip

                    # Mailbox photos
                    if ($tempImageSourceFullname -or $tempImageAlternativeText) {
                        foreach ($VariableName in $PictureVariablesArray) {
                            try { WatchCatchableExitSignal } catch { }

                            if (
                                $(if ($tempImageSourceFullname) { ((Split-Path $tempImageSourceFullname -Leaf) -ilike "*$($Variablename[0])*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($Variablename[0])*") })
                            ) {
                                if ($null -ne $($ReplaceHash[$Variablename[0]])) {
                                    $tempImageSourceFullname = (Join-Path -Path $script:tempDir -ChildPath ($Variablename[0] + $Variablename[1] + '.jpeg'))
                                }
                            } elseif (
                                $(if ($tempImageSourceFullname) { ((Split-Path $tempImageSourceFullname -Leaf) -ilike "*$($Variablename[0] -ireplace '\$$', 'DELETEEMPTY$')*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($Variablename[0] -ireplace '\$$', 'DELETEEMPTY$')*") })
                            ) {
                                if ($null -ne $($ReplaceHash[$Variablename[0]])) {
                                    $tempImageSourceFullname = (Join-Path -Path $script:tempDir -ChildPath ($Variablename[0] + $Variablename[1] + '.jpeg'))
                                } else {
                                    $image.delete()
                                    $tempImageIsDeleted = $true
                                    break
                                }
                            }

                            if ((-not $tempImageIsDeleted) -and ($tempImageAlternativeText)) {
                                $tempImageAlternativeText = $($tempImageAlternativeText) -ireplace [Regex]::Escape($Variablename[0]), ''
                                $tempImageAlternativeText = $($tempImageAlternativeText) -ireplace [Regex]::Escape($($Variablename[0] -ireplace '\$$', 'DELETEEMPTY$')), ''
                            }
                        }

                        if ($tempImageIsDeleted) {
                            continue
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }

                    # Other images
                    if (
                        $(if ($tempImageSourceFullname) { ((Split-Path $tempImageSourceFullname -Leaf) -ilike '*$*DELETEEMPTY$*') }) -or
                        $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike '*$*DELETEEMPTY$*') })
                    ) {
                        foreach ($Variablename in @(@($ReplaceHash.Keys) | Where-Object { $_ -inotin @('$CurrentMailboxPhoto$', '$CurrentMailboxManagerPhoto$', '$CurrentUserPhoto$', '$CurrentUserManagerPhoto$') })) {
                            try { WatchCatchableExitSignal } catch { }

                            $tempImageVariableString = $Variablename -ireplace '\$$', 'DELETEEMPTY$'

                            if (
                                $(if ($tempImageSourceFullname) { ((Split-Path $tempImageSourceFullname -Leaf) -ilike "*$($tempImageVariableString)*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($tempImageVariableString)*") })
                            ) {
                                if ($($ReplaceHash[$Variablename])) {
                                    if ($tempImageAlternativeText) {
                                        $tempImageAlternativeText = $tempImageAlternativeText -ireplace [Regex]::Escape($tempImageVariableString), ''
                                    }
                                } else {
                                    $image.delete()
                                    $tempImageIsDeleted = $true
                                    break
                                }
                            }
                        }
                    }

                    if ($tempImageIsDeleted) {
                        continue
                    }

                    foreach ($replaceKey in @($replaceHash.Keys | Where-Object { $_ -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' }) } | Sort-Object -Culture $TemplateFilesSortCulture)) {
                        try { WatchCatchableExitSignal } catch { }

                        if ($replaceKey) {
                            if ($null -ne $tempImageAlternativeText) {
                                $tempImageAlternativeText = $tempImageAlternativeText -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }

                            if ($null -ne $tempimagehyperlinkAddress) {
                                $tempimagehyperlinkAddress = $tempimagehyperlinkAddress -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }

                            if ($null -ne $tempimagehyperlinkSubAddress) {
                                $tempimagehyperlinkSubAddress = $tempimagehyperlinkSubAddress -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }

                            if ($null -ne $tempimagehyperlinkEmailSubject) {
                                $tempimagehyperlinkEmailSubject = $tempimagehyperlinkEmailSubject -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }

                            if ($null -ne $tempimagehyperlinkScreenTip) {
                                $tempimagehyperlinkScreenTip = $tempimagehyperlinkScreenTip -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }
                        }
                    }

                    if (
                        $($tempImageSourceFullname) -and
                        $($image.linkformat.sourcefullname) -and
                        $($tempImageSourceFullname -ine $image.linkformat.sourcefullname)
                    ) {
                        $image.linkformat.sourcefullname = $tempImageSourceFullname
                    }

                    if ($null -ne $tempImageAlternativeText) {
                        $image.alternativeText = $tempImageAlternativeText
                    }

                    if ($null -ne $tempimagehyperlinkAddress) {
                        $image.hyperlink.Address = $tempImageHyperlinkAddress
                    }

                    if ($null -ne $tempimagehyperlinkSubAddress) {
                        $image.hyperlink.SubAddress = $tempImageHyperlinkSubAddress
                    }

                    if ($null -ne $tempimagehyperlinkEmailSubject) {
                        $image.hyperlink.EmailSubject = $tempImageHyperlinkEmailSubject
                    }

                    if ($null -ne $tempimagehyperlinkScreenTip) {
                        $image.hyperlink.ScreenTip = $tempImageHyperlinkScreenTip
                    }
                }
            } catch {
                Write-Host $error[0]
                Write-Host "$Indent        Error replacing picture variables in Word. Exit." -ForegroundColor Red
                Write-Host "$Indent        If the error says 'Access denied', your environment may require to assign a Microsoft Purview Information Protection sensitivity label to your DOCX templates." -ForegroundColor Red
                $script:ExitCode = 20
                $script:ExitCodeDescription = 'Error replacing picture variables in Word.'
                exit
            }


            try { WatchCatchableExitSignal } catch { }


            Write-Host "$Indent      Replace non-picture variables"
            $wdFindContinue = 1
            $MatchCase = $false
            $MatchWholeWord = $true
            $MatchWildcards = $False
            $MatchSoundsLike = $False
            $MatchAllWordForms = $False
            $Forward = $True
            $Wrap = $wdFindContinue
            $Format = $False
            $wdFindContinue = 1
            $ReplaceAll = 2

            $script:COMWordShowFieldCodesOriginal = $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes

            try {
                # Replace in view without field codes
                $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $false

                $script:COMWord.ActiveDocument.Select()
                $tempWordText = $script:COMWord.Selection.Text
                $script:COMWord.Selection.Collapse()

                foreach ($replaceKey in @($replaceHash.Keys | Where-Object { ($_ -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' })) -and ($tempWordText -imatch [regex]::escape($_)) } | Sort-Object -Culture $TemplateFilesSortCulture )) {
                    try { WatchCatchableExitSignal } catch { }

                    $script:COMWord.Selection.Find.Execute($replaceKey, $MatchCase, $MatchWholeWord, `
                            $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, `
                            $Wrap, $Format, $(($replaceHash.$replaceKey -ireplace "`r`n", '^p') -ireplace "`n", '^l'), $ReplaceAll) | Out-Null
                }

                # Restore original view
                $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal

                $tempWordText = $null

                try { WatchCatchableExitSignal } catch { }

                # Replace in field codes
                foreach ($field in $script:COMWord.ActiveDocument.Fields) {
                    try { WatchCatchableExitSignal } catch { }

                    $tempWordFieldCodeOriginal = $field.Code.Text
                    $tempWordFieldCodeNew = $tempWordFieldCodeOriginal

                    foreach ($replaceKey in @($replaceHash.Keys | Where-Object { ($_ -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' })) } | Sort-Object -Culture $TemplateFilesSortCulture )) {
                        $tempWordFieldCodeNew = $tempWordFieldCodeNew -ireplace [regex]::escape($replaceKey), $($replaceHash.$replaceKey)
                    }

                    if ($tempWordFieldCodeOriginal -ne $tempWordFieldCodeNew) {
                        $field.Code.Text = $tempWordFieldCodeNew
                    }
                }
            } catch {
                Write-Host $error[0]
                Write-Host "$Indent        Error replacing non-picture variables in Word. Exit." -ForegroundColor Red
                Write-Host "$Indent        If the error says 'Access denied', your environment may require to assign a Microsoft Purview Information Protection sensitivity label to your DOCX templates." -ForegroundColor Red
                $script:ExitCode = 21
                $script:ExitCodeDescription = 'Error replacing non-picture variables in Word.'
                exit
            }

            try { WatchCatchableExitSignal } catch { }

            # Save changed document, it's later used for export to .htm, .rtf and .txt
            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatDocumentDefault')

            try { WatchCatchableExitSignal } catch { }

            try {
                # Overcome Word security warning when export contains embedded pictures
                if ($null -eq (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                }

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                try { WatchCatchableExitSignal } catch { }

                # Save
                $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            } catch {
                # Restore original security setting after error
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null

                Start-Sleep -Seconds 2

                # Overcome Word security warning when export contains embedded pictures
                if ($null -eq (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                }

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                try { WatchCatchableExitSignal } catch { }

                # Save
                $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            }

            try { WatchCatchableExitSignal } catch { }

            # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
            # Close the document to remove in-memory references to already deleted images
            $script:COMWord.ActiveDocument.Saved = $true
            $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)

            try { WatchCatchableExitSignal } catch { }

            # Export to .htm
            Write-Host "$Indent      Export to HTM format"
            $path = $([System.IO.Path]::ChangeExtension($path, '.docx'))

            try { WatchCatchableExitSignal } catch { }

            $script:COMWord.Documents.Open($path, $false, $false, $false) | Out-Null

            try { WatchCatchableExitSignal } catch { }

            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatFilteredHTML')
            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

            $script:WordWebOptions = $script:COMWord.ActiveDocument.WebOptions

            $script:COMWord.ActiveDocument.WebOptions.TargetBrowser = 4 # IE6, which is the maximum
            $script:COMWord.ActiveDocument.WebOptions.BrowserLevel = 2 # IE6, which is the maximum
            $script:COMWord.ActiveDocument.WebOptions.AllowPNG = $true
            $script:COMWord.ActiveDocument.WebOptions.OptimizeForBrowser = $false
            $script:COMWord.ActiveDocument.WebOptions.RelyOnCSS = $true
            $script:COMWord.ActiveDocument.WebOptions.RelyOnVML = $false
            $script:COMWord.ActiveDocument.WebOptions.Encoding = 65001 # Outlook uses 65001 (UTF8) for .htm, but 1200 (UTF16LE a.k.a Unicode) for .txt
            $script:COMWord.ActiveDocument.WebOptions.OrganizeInFolder = $true
            $script:COMWord.ActiveDocument.WebOptions.PixelsPerInch = 96
            $script:COMWord.ActiveDocument.WebOptions.ScreenSize = 10 # 1920x1200
            $script:COMWord.ActiveDocument.WebOptions.UseLongFileNames = $true

            $script:COMWord.ActiveDocument.WebOptions.UseDefaultFolderSuffix()
            $pathHtmlFolderSuffix = $script:COMWord.ActiveDocument.WebOptions.FolderSuffix

            try {
                # Overcome Word security warning when export contains embedded pictures
                if ($null -eq (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                }

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                try { WatchCatchableExitSignal } catch { }

                # Save
                $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            } catch {
                # Restore original security setting after error
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null

                Start-Sleep -Seconds 2

                # Overcome Word security warning when export contains embedded pictures
                if ($null -eq (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                }

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                try { WatchCatchableExitSignal } catch { }

                # Save
                $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            }

            try { WatchCatchableExitSignal } catch { }

            # Restore original WebOptions
            try {
                if ($script:WordWebOptions) {
                    foreach ($property in @('TargetBrowser', 'BrowserLevel', 'AllowPNG', 'OptimizeForBrowser', 'RelyOnCSS', 'RelyOnVML', 'Encoding', 'OrganizeInFolder', 'PixelsPerInch', 'ScreenSize', 'UseLongFileNames')) {
                        $script:COMWord.ActiveDocument.WebOptions.$property = $script:WordWebOptions.$property
                    }
                }
            } catch {}

            try { WatchCatchableExitSignal } catch { }

            # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
            $script:COMWord.ActiveDocument.Saved = $true

            Write-Host "$Indent        Export high-res images"

            if ($DocxHighResImageConversion) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('DocxHighResImageConversion')))) {
                    $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)

                    Write-Host "$Indent          Can not export high-res images." -ForegroundColor Yellow
                    Write-Host "$Indent          The 'DocxHighResImageConversion' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                    Write-Host "$Indent          Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DocxHighResImageConversion()

                    if ($FeatureResult -ne 'true') {
                        try {
                            $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)
                        } catch {
                        }
                        Write-Host "$Indent          Error converting high resolution images from DOCX template." -ForegroundColor Yellow
                        Write-Host "$Indent          $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "$Indent          Parameter 'DocxHighResImageConversion' is not enabled, skipping task."

                $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent        Copy HTM image width and height attributes to style attribute"
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

        if ($($PSVersionTable.PSEdition) -ieq 'Core') {
            $AngleSharpConfig = [AngleSharp.CssConfigurationExtensions]::WithCss([AngleSharp.Configuration]::Default)
            $AngleSharpBrowsingContext = [AngleSharp.BrowsingContext]::New($AngleSharpConfig)
            $AngleSharpHtmlParser = $AngleSharpBrowsingContext.GetType().GetMethod('GetService').MakeGenericMethod([AngleSharp.Html.Parser.IHtmlParser]).Invoke($AngleSharpBrowsingContext, $null)
            $AngleSharpParsedDocument = $AngleSharpHtmlParser.ParseDocument("$(Get-Content -LiteralPath $path -Encoding UTF8 -Raw)")

            foreach ($image in @($AngleSharpParsedDocument.images)) {
                if (-not $image.Attributes['style']) {
                    $image.SetAttribute('style', '')
                }

                if ($image.Attributes['width'].TextContent) {
                    $image.SetAttribute('style', $('' + $image.Attributes['style'].TextContent + ';width:' + $($image.Attributes['width'].TextContent) + ';'))
                }

                if ($image.Attributes['height'].TextContent) {
                    $image.SetAttribute('style', $('' + $image.Attributes['style'].TextContent + ';height:' + $($image.Attributes['height'].TextContent) + ';'))
                }
            }

            [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $AngleSharpParsedDocument.documentelement.outerhtml)
        } else {
            $tempVerbosePreference = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'
            $html = New-Object -ComObject 'HTMLFile'
            $VerbosePreference = $tempVerbosePreference

            try {
                # PowerShell Desktop with Office
                $html.IHTMLDocument2_write((Get-Content -LiteralPath $path -Encoding UTF8 -Raw))
            } catch {
                # PowerShell Desktop without Office, PowerShell 6+
                $html.write([System.Text.Encoding]::Unicode.GetBytes((Get-Content -LiteralPath $path -Encoding UTF8 -Raw)))
            }

            foreach ($image in @($html.images)) {
                $image.style.setAttribute('width', ($image.attributes | Where-Object { $_.nodename -ieq 'width' }).textContent)
                $image.style.setAttribute('height', ($image.attributes | Where-Object { $_.nodename -ieq 'height' }).textContent)
            }

            [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $html.documentelement.outerhtml)

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($html) | Out-Null
            Remove-Variable -Name 'html'
        }

        try { WatchCatchableExitSignal } catch { }

        if ($MoveCSSInline) {
            Write-Host "$Indent        Move CSS inline"

            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
            $tempFileContent = Get-Content -LiteralPath $path -Encoding UTF8 -Raw

            # Use a separate runspace for PreMailer.Net, as there are DLL conflicts in PowerShell 5.x with Invoke-RestMethod
            # Do not use jobs, as they fall back to Constrained Language Mode in secured environments, which makes Import-Module fail
            $MoveCSSInlineResult = MoveCssInline $tempFileContent

            if ($MoveCSSInlineResult.StartsWith('Failed: ')) {
                Write-Host "$Indent          $MoveCSSInlineResult" -ForegroundColor Yellow
            } else {
                [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $MoveCSSInlineResult)
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent        Add marker to final HTM file"
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
        $tempFileContent = Get-Content -LiteralPath $path -Encoding UTF8 -Raw

        if ($tempFileContent -inotmatch [regex]::escape($HTMLMarkerTag)) {
            if ($tempFileContent -imatch '<\s*head\b[^>]*>') {
                $tempFileContent = $tempFileContent -ireplace '<\s*head\b[^>]*>', "`${0} $($HTMLMarkerTag)"
            } else {
                $tempFileContent = $tempFileContent -ireplace '<\s*html\b[^>]*>', "`${0} <HEAD> $($HTMLMarkerTag) </HEAD>"
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent        Modify connected folder name"

        foreach ($pathConnectedFolderName in $pathConnectedFolderNames) {
            try { WatchCatchableExitSignal } catch { }

            $tempFileContent = $tempFileContent -ireplace ('(\s*src=")(' + [regex]::escape($pathConnectedFolderName) + '\/)'), ('$1' + "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.value)).files/")
            Rename-Item (Join-Path -Path (Split-Path $path) -ChildPath $($pathConnectedFolderName)) $([System.IO.Path]::GetFileNameWithoutExtension($Signature.value) + '.files') -ErrorAction SilentlyContinue
        }

        [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $tempFileContent)

        try { WatchCatchableExitSignal } catch { }

        if (-not $ProcessOOF) {
            if ($EmbedImagesInHtml) {
                Write-Host "$Indent        Embed local images"

                [SetOutlookSignatures.Common]::ConvertToSingleFileHtml($path, $path)
            }
        } else {
            [SetOutlookSignatures.Common]::ConvertToSingleFileHtml($path, ((Join-Path -Path $script:tempDir -ChildPath $Signature.Value)))
        }

        try { WatchCatchableExitSignal } catch { }

        if (-not $ProcessOOF) {
            if ($CreateRtfSignatures) {
                Write-Host "$Indent      Export to RTF format"

                try { WatchCatchableExitSignal } catch { }

                # If possible, use .docx file to avoid problems with MS Information Protection
                $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
                $script:COMWord.Documents.Open($path, $false, $false, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, 65001) | Out-Null

                try { WatchCatchableExitSignal } catch { }

                $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatRTF')
                $path = $([System.IO.Path]::ChangeExtension($path, '.rtf'))

                try {
                    # Overcome Word security warning when export contains embedded pictures
                    if ($null -eq (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                    }

                    if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                        $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                    }

                    if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                    }

                    try { WatchCatchableExitSignal } catch { }

                    # Save
                    $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                    # Restore original security setting
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                } catch {
                    # Restore original security setting after error
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null

                    Start-Sleep -Seconds 2

                    # Overcome Word security warning when export contains embedded pictures
                    if ($null -eq (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                    }

                    if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                        $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                    }

                    if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                    }

                    try { WatchCatchableExitSignal } catch { }

                    # Save
                    $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                    # Restore original security setting
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }

                try { WatchCatchableExitSignal } catch { }

                # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
                # Close the document as conversion to .rtf happens from .htm
                $script:COMWord.ActiveDocument.Saved = $true
                $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null

                try { WatchCatchableExitSignal } catch { }

                Write-Host "$Indent        Shrink RTF file"
                $((Get-Content -LiteralPath $path -Raw -Encoding Ascii) -ireplace '\{\\nonshppict[\s\S]*?\}\}', '') | Set-Content -LiteralPath $path -Encoding Ascii
            }

            try { WatchCatchableExitSignal } catch { }

            if ($CreateTxtSignatures) {
                Write-Host "$Indent      Export to TXT format"

                $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

                if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                    $AngleSharpConfig = [AngleSharp.CssConfigurationExtensions]::WithRenderDevice([AngleSharp.CssConfigurationExtensions]::WithCss([AngleSharp.Configuration]::Default), (New-Object -TypeName AngleSharp.Css.DefaultRenderDevice -Property @{DeviceWidth = 1920; DeviceHeight = 1080; ViewPortWidth = 1920; ViewPortHeight = 1080 }))
                    $AngleSharpBrowsingContext = [AngleSharp.BrowsingContext]::New($AngleSharpConfig)
                    $AngleSharpHtmlParser = $AngleSharpBrowsingContext.GetType().GetMethod('GetService').MakeGenericMethod([AngleSharp.Html.Parser.IHtmlParser]).Invoke($AngleSharpBrowsingContext, $null)
                    $AngleSharpParsedDocument = $AngleSharpHtmlParser.ParseDocument("$(Get-Content -LiteralPath $path -Encoding UTF8 -Raw)")
                } else {
                    $tempVerbosePreference = $VerbosePreference
                    $VerbosePreference = 'SilentlyContinue'
                    $html = New-Object -ComObject 'HTMLFile'
                    $VerbosePreference = $tempVerbosePreference

                    try {
                        # PowerShell Desktop with Office
                        $html.IHTMLDocument2_write((Get-Content -LiteralPath $path -Encoding UTF8 -Raw))
                    } catch {
                        # PowerShell Desktop without Office, PowerShell 6+
                        $html.write([System.Text.Encoding]::Unicode.GetBytes((Get-Content -LiteralPath $path -Encoding UTF8 -Raw)))
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                $path = $([System.IO.Path]::ChangeExtension($path, '.txt'))

                if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                    $AngleSharpParsedDocumentInnerText = [AngleSharp.Dom.ElementExtensions]::GetInnerText($AngleSharpParsedDocument.body)
                    $AngleSharpParsedDocumentInnerText = $AngleSharpParsedDocumentInnerText -ireplace "(?<!`n)`n `n", "`n"

                    (1..1000) | ForEach-Object {
                        $AngleSharpParsedDocumentInnerText = $AngleSharpParsedDocumentInnerText -ireplace "(?<!`n)`n{$($_)}(?!`n)", ("`n" * $(if ($_ % 2 -eq 0) { ($_ / 2) + $(if ($_ -gt 2) { 1 }else { 0 }) } else { (($_ + 1) / 2) }))
                    }

                    $AngleSharpParsedDocumentInnerText | Out-File -LiteralPath $path -Encoding Unicode -Force # Outlook uses 65001 (UTF8) for .htm, but 1200 (UTF16LE a.k.a Unicode) for .txt
                } else {
                    $html.body.innerText | Out-File -LiteralPath $path -Encoding Unicode -Force # Outlook uses 65001 (UTF8) for .htm, but 1200 (UTF16LE a.k.a Unicode) for .txt
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        if (-not $ProcessOOF) {
            Write-Host "$Indent      Upload signature to Exchange Online as roaming signature"

            if ($MirrorCloudSignatures -eq $true) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('RoamingSignaturesUpload')))) {
                    Write-Host "$Indent        Can not upload signature to Exchange Online." -ForegroundColor Yellow
                    Write-Host "$Indent        The 'MirrorCloudSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Yellow
                    Write-Host "$Indent        Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesUpload()

                    if ($FeatureResult -ne 'true') {
                        Write-Host "$Indent        Error uploading roaming signatures to the cloud." -ForegroundColor Yellow
                        Write-Host "$Indent        $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "$Indent        Parameter 'MirrorCloudSignatures' is not enabled, skipping task."
            }

            foreach ($SignaturePath in $SignaturePaths) {
                try { WatchCatchableExitSignal } catch { }

                Write-Host "$Indent      Copy signature files to '$SignaturePath'"

                RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.htm')))

                foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                    try { WatchCatchableExitSignal } catch { }

                    RemoveItemAlternativeRecurse -LiteralPath ((Join-Path -Path $SignaturePath -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.value))") + $ConnectedFilesFolderName)
                }

                Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.htm')) -Destination $((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.htm')))) -Force

                try { WatchCatchableExitSignal } catch { }

                if ($EmbedImagesInHtml -eq $false) {
                    if (Test-Path (Join-Path -Path (Split-Path $path) -ChildPath "$([System.IO.Path]::ChangeExtension($Signature.value, '.files'))")) {
                        Copy-Item -LiteralPath (Join-Path -Path (Split-Path $path) -ChildPath "$([System.IO.Path]::ChangeExtension($Signature.value, '.files'))") -Destination $SignaturePath -Force -Recurse
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($CreateRtfSignatures -eq $true) {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))
                    Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.rtf')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))) -Force
                } else {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))
                }

                try { WatchCatchableExitSignal } catch { }

                if ($CreateTxtSignatures -eq $true) {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))
                    Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.txt')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))) -Force
                } else {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))
                }

                try { WatchCatchableExitSignal } catch { }

                if ($SignatureFilesWriteProtect.containskey($TemplateIniSettingsIndex)) {
                    Write-Host "$Indent      Write protect signature files"
                    @('.htm', '.rtf', '.txt') | ForEach-Object {
                        $file = Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, $_))
                        if (Test-Path -Path $file -PathType Leaf) {
                            (Get-Item $file -Force).Attributes += 'ReadOnly'
                        }
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($macOSSignaturesScriptable) {
                    Write-Host "$Indent      Create Outlook for Mac signature"

                    @($(@"
tell application "Microsoft Outlook"
    try
        set signatureName to "$(Split-Path $signature.value -LeafBase)"
        set htmlContent to (read POSIX file "$(([System.IO.Path]::ChangeExtension($path, '.htm')))" as «class utf8»)

        -- Check if the signature exists
        set signatureList to signatures
        set signatureExists to false

        repeat with aSignature in signatureList
            if name of aSignature is signatureName then
                set signatureExists to true
                exit repeat
            end if
        end repeat

        if signatureExists then
            -- Update the existing signature
            set content of signature signatureName to htmlContent
        else
            -- Create a new signature
            make new signature with properties {name:signatureName, content:htmlContent}
        end if
    on error errorMessage
        log "$Indent        Error: " & errorMessage
    end try
end tell
"@ | osascript *>&1)) | ForEach-Object { Write-Host $_.tostring() }
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent      Remove temporary files"
        foreach ($extension in ('.docx', '.htm', '.rtf', '.txt')) {
            Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, $extension)) -ErrorAction SilentlyContinue | Out-Null

            if ($pathHighResHtml) {
                Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($pathHighResHtml, $extension)) -ErrorAction SilentlyContinue | Out-Null
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Foreach ($file in @(Get-ChildItem -Path ("$($script:tempDir)\*" + [System.IO.Path]::GetFileNameWithoutExtension($path) + '*') -Directory).FullName) {
            Remove-Item -LiteralPath $file -Force -Recurse -ErrorAction SilentlyContinue
        }

        try { WatchCatchableExitSignal } catch { }

        if ($pathHighResHtml) {
            Foreach ($file in @(Get-ChildItem -Path ("$($script:tempDir)\*" + [System.IO.Path]::GetFileNameWithoutExtension($pathHighResHtml) + '*') -Directory).FullName) {
                Remove-Item -LiteralPath $file -Force -Recurse -ErrorAction SilentlyContinue
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath $([System.IO.Path]::ChangeExtension($signature.value, '.files'))) -Force -Recurse -ErrorAction SilentlyContinue
    }

    try { WatchCatchableExitSignal } catch { }

    if ((-not $ProcessOOF)) {
        # Set default signature for new emails
        if ($SignatureFilesDefaultNew.containskey($TemplateIniSettingsIndex)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                try { WatchCatchableExitSignal } catch { }

                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    if ($CurrentTemplateIsForAliasSmtp) {
                        $NewSigExpected."$($CurrentTemplateIsForAliasSmtp.ToLower())" = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }

                    $NewSigExpected."$(($MailAddresses[$AccountNumberRunning]).ToLower())" = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                    if (-not $SimulateUser) {
                        if ($RegistryPaths[$j] -ilike '*\9375CFF0413111d3B88A00104B2A6676\*') {
                            Write-Host "$Indent      Set signature as default for new messages (Outlook profile '$(($RegistryPaths[$j] -split '\\')[8])')"

                            if ($OutlookFileVersion -ge '16.0.0.0') {
                                New-ItemProperty -Path $RegistryPaths[$j] -Name 'New Signature' -PropertyType String -Value (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') -Force | Out-Null
                            } else {
                                New-ItemProperty -Path $RegistryPaths[$j] -Name 'New Signature' -PropertyType Binary -Value ([byte[]](([System.Text.Encoding]::Unicode.GetBytes(((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')) + "`0")))) -Force | Out-Null
                            }
                        } else {
                            $script:GraphUserDummyMailboxDefaultSigNew = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                        }
                    } else {
                        @('htm', 'rtf', 'txt') | ForEach-Object {
                            if (Test-Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)"))) {
                                $script:GraphUserDummyMailboxDefaultSigNew = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                                if ($_ -ieq 'htm') {
                                    [SetOutlookSignatures.Common]::ConvertToSingleFileHtml($(Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)")), $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultNew.$($_)")))
                                } else {
                                    Copy-Item -LiteralPath $(Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)")) -Destination $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultNew.$($_)")) -Force
                                }
                            }
                        }
                    }
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        # Set default signature for replies and forwarded emails
        try { WatchCatchableExitSignal } catch { }

        if ($SignatureFilesDefaultReplyFwd.containskey($TemplateIniSettingsIndex)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                try { WatchCatchableExitSignal } catch { }

                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    if ($CurrentTemplateIsForAliasSmtp) {
                        $ReplySigExpected."$($CurrentTemplateIsForAliasSmtp.ToLower())" = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }

                    $ReplySigExpected."$(($MailAddresses[$AccountNumberRunning]).ToLower())" = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                    if (-not $SimulateUser) {
                        if ($RegistryPaths[$j] -ilike '*\9375CFF0413111d3B88A00104B2A6676\*') {
                            Write-Host "$Indent      Set signature as default for reply/forward messages (Outlook profile '$(($RegistryPaths[$j] -split '\\')[8])')"

                            if ($OutlookFileVersion -ge '16.0.0.0') {
                                New-ItemProperty -Path $RegistryPaths[$j] -Name 'Reply-Forward Signature' -PropertyType String -Value (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') -Force | Out-Null
                            } else {
                                New-ItemProperty -Path $RegistryPaths[$j] -Name 'Reply-Forward Signature' -PropertyType Binary -Value ([byte[]](([System.Text.Encoding]::Unicode.GetBytes(((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')) + "`0")))) -Force | Out-Null
                            }
                        } else {
                            $script:GraphUserDummyMailboxDefaultSigReply = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                        }
                    } else {
                        @('htm', 'rtf', 'txt') | ForEach-Object {
                            if (Test-Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)"))) {
                                $script:GraphUserDummyMailboxDefaultSigReply = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                                if ($_ -ieq 'htm') {
                                    [SetOutlookSignatures.Common]::ConvertToSingleFileHtml($(Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)")), $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultReplyFwd.$($_)")))
                                } else {
                                    Copy-Item -LiteralPath $(Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)")) -Destination $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultReplyFwd.$($_)")) -Force
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function CheckADConnectivity {
    param (
        [array]$CheckDomains,
        [string]$CheckProtocolText,
        [string]$Indent
    )

    try { WatchCatchableExitSignal } catch { }

    [void][runspacefactory]::CreateRunspacePool()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 25)
    $RunspacePool.Open()

    for ($DomainNumber = 0; $DomainNumber -lt $CheckDomains.count; $DomainNumber++) {
        try { WatchCatchableExitSignal } catch { }

        if ($($CheckDomains[$DomainNumber]) -eq '') {
            continue
        }

        $PowerShell = [powershell]::Create()
        $PowerShell.RunspacePool = $RunspacePool

        [void]$PowerShell.AddScript({
                Param (
                    [string]$CheckDomain,
                    [string]$CheckProtocolText
                )
                $DebugPreference = 'Continue'
                Write-Debug "Start(Ticks) = $((Get-Date).Ticks)"
                Write-Output "$CheckDomain"
                $Search = New-Object DirectoryServices.DirectorySearcher
                $Search.PageSize = 1000
                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("$($CheckProtocolText)://$CheckDomain")
                $Search.filter = '(objectclass=user)'
                try {
                    try { WatchCatchableExitSignal } catch { }
                    $null = ([ADSI]"$(($Search.FindOne()).path)")
                    try { WatchCatchableExitSignal } catch { }
                    Write-Output 'QueryPassed'
                } catch {
                    Write-Output 'QueryFailed'
                }
            }).AddArgument($($CheckDomains[$DomainNumber])).AddArgument($CheckProtocolText)

        $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
        $Handle = $PowerShell.BeginInvoke($Object, $Object)
        $temp = '' | Select-Object PowerShell, Handle, Object, StartTime, Done
        $temp.PowerShell = $PowerShell
        $temp.Handle = $Handle
        $temp.Object = $Object
        $temp.StartTime = $null
        $temp.Done = $false
        [void]$script:jobs.Add($Temp)
    }

    while (($script:jobs.Done | Where-Object { $_ -eq $false }).count -ne 0) {
        try { WatchCatchableExitSignal } catch { }

        foreach ($job in $script:jobs) {
            try { WatchCatchableExitSignal } catch { }

            if (($null -eq $job.StartTime) -and ($job.Powershell.Streams.Debug[0].Message -imatch 'Start')) {
                $StartTicks = $job.powershell.Streams.Debug[0].Message -ireplace '[^0-9]'
                $job.StartTime = [Datetime]::MinValue + [TimeSpan]::FromTicks($StartTicks)
            }

            if ($null -ne $job.StartTime) {
                if ((($job.handle.IsCompleted -eq $true) -and ($job.Done -eq $false)) -or (($job.Done -eq $false) -and ((New-TimeSpan -Start $job.StartTime -End (Get-Date)).TotalSeconds -ge 5))) {
                    $data = $job.Object[0..$(($job.object).count - 1)]

                    Write-Host "$Indent$($data[0])"

                    if ($data -icontains 'QueryPassed') {
                        Write-Host "$Indent  $CheckProtocolText query successful"
                        $returnvalue = $true
                    } else {
                        Write-Host "$Indent  $CheckProtocolText query failed, remove domain from list." -ForegroundColor Red
                        Write-Host "$Indent  If this error is permanent, check firewalls, DNS and AD trust. Consider parameter 'TrustsToCheckForGroups' to not use this domain." -ForegroundColor Red

                        if ($TrustsToCheckForGroups -icontains $data[0]) {
                            $TrustsToCheckForGroups.remove($data[0])
                        }

                        $LookupDomainsToTrusts.remove($data[0])

                        $returnvalue = $false
                    }

                    $job.Done = $true
                }
            }
        }

        Start-Sleep -Seconds 1
    }

    try { WatchCatchableExitSignal } catch { }

    return $returnvalue
}


function MoveCssInline {
    param (
        $HtmlCode
    )

    try { WatchCatchableExitSignal } catch { }

    [void][runspacefactory]::CreateRunspacePool()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 1)
    $RunspacePool.Open()

    $PowerShell = [powershell]::Create()
    $PowerShell.RunspacePool = $RunspacePool

    [void]$PowerShell.AddScript({
            Param (
                $HtmlCode,
                $path
            )

            $DebugPreference = 'Continue'
            Write-Debug "Start(Ticks) = $((Get-Date).Ticks)"

            try {
                Import-Module (Join-Path -Path $path -ChildPath 'PreMailer.Net.dll') -Force -ErrorAction Stop

                if ($UseHtmTemplates) {
                    Write-Debug $([PreMailer.Net.PreMailer]::MoveCssInline($HtmlCode).html)
                } else {
                    Write-Debug $([PreMailer.Net.PreMailer]::MoveCssInline($HtmlCode, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true, $true).html)
                }
            } catch {
                $MoveCSSInlineError = $_
                Write-Debug "Failed: $MoveCSSInlineError"
            }
        }).AddArgument($HtmlCode).AddArgument($script:PreMailerNetModulePath)

    $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $Handle = $PowerShell.BeginInvoke($Object, $Object)
    $temp = '' | Select-Object PowerShell, Handle, Object, StartTime, Done
    $temp.PowerShell = $PowerShell
    $temp.Handle = $Handle
    $temp.Object = $Object
    $temp.StartTime = $null
    $temp.Done = $false
    [void]$script:jobs.Add($Temp)

    while (($script:jobs.Done | Where-Object { $_ -eq $false }).count -ne 0) {
        try { WatchCatchableExitSignal } catch { }

        foreach ($job in $script:jobs) {
            try { WatchCatchableExitSignal } catch { }

            if (($null -eq $job.StartTime) -and ($job.Powershell.Streams.Debug[0].Message -imatch 'Start')) {
                $StartTicks = $job.powershell.Streams.Debug[0].Message -ireplace '[^0-9]'
                $job.StartTime = [Datetime]::MinValue + [TimeSpan]::FromTicks($StartTicks)
            }

            if ($null -ne $job.StartTime) {
                if ((($job.handle.IsCompleted -eq $true) -and ($job.Done -eq $false)) -or (($job.Done -eq $false) -and ((New-TimeSpan -Start $job.StartTime -End (Get-Date)).TotalSeconds -ge 5))) {
                    $data = $job.Object[0..$(($job.object).count - 1)]
                    #if ($job.Powershell.Streams.Debug[1].Message.StartsWith('Failed: ')) {
                    #    $returnvalue = $HtmlCode
                    #} else {
                    $returnvalue = $job.Powershell.Streams.Debug[1].Message
                    #}
                    $job.Done = $true
                }
            }
        }

        Start-Sleep -Seconds 1
    }

    try { WatchCatchableExitSignal } catch { }

    return $returnvalue
}


$CheckPathScriptblock = {
    # A script block runs in the scope of the caller, which is different from functions
    # This makes it interesting for manipulating variables, so take care of variable names
    [cmdletbinding()]
    param (
        [ref]$CheckPathRefPath,
        [switch]$CheckPathSilent = $false,
        [switch]$CheckPathCreate = $false,
        [string]$ExpectedPathType = 'Container'
    )

    try { WatchCatchableExitSignal } catch { }

    $CheckPathPath = $CheckPathRefPath.Value

    try {
        Write-Verbose "      Execute config file '$GraphConfigFile'"

        if (Test-Path -LiteralPath $GraphConfigFile -PathType Leaf) {
            . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $GraphConfigFile -Encoding UTF8 -Raw)))
        } elseif (Test-Path -LiteralPath $(Join-Path -Path $PSScriptRoot -ChildPath '.\config\default graph config.ps1') -PathType Leaf) {
            Write-Verbose '        Not accessible, use default Graph config file'
            . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $(Join-Path -Path $PSScriptRoot -ChildPath '.\config\default graph config.ps1') -Encoding UTF8 -Raw)))
        } else {
            Write-Verbose '        Not accessible, and default Graph config file not found'
        }

        try { WatchCatchableExitSignal } catch { }

        @(
            'displayName',
            'givenName',
            'id',
            'mail',
            'mailNickname',
            'onPremisesDistinguishedName',
            'onPremisesDomainName',
            'onPremisesExtensionAttributes',
            'onPremisesImmutableId',
            'onPremisesSamAccountName',
            'onPremisesSecurityIdentifier',
            'onPremisesUserPrincipalName',
            'proxyAddresses',
            'securityIdentifier',
            'surname',
            'userPrincipalName'
        ) | ForEach-Object {
            if ($GraphUserProperties -inotcontains $_) {
                $GraphUserProperties += $_
            }
        }

        if (-not $GraphUserAttributeMapping) {
            $GraphUserAttributeMapping = @{}
        }

        $GraphUserAttributeMapping['distinguishedname'] = 'onPremisesDistinguishedName'
        $GraphUserAttributeMapping['mailboxsettings'] = 'mailboxSettings'
        $GraphUserAttributeMapping['mailNickname'] = 'mailNickname'
        $GraphUserAttributeMapping['objectguid'] = 'id'
        $GraphUserAttributeMapping['objectsid'] = 'securityIdentifier'
        $GraphUserAttributeMapping['onpremisesdomainname'] = 'onPremisesDomainName'
        $GraphUserAttributeMapping['onpremisessecurityidentifier'] = 'onPremisesSecurityIdentifier'
        $GraphUserAttributeMapping['userprincipalname'] = 'userPrincipalName'
    } catch {
        Write-Host $error[0]
        Write-Host "        Problem executing content of '$GraphConfigFile'. Exit." -ForegroundColor Red
        $script:ExitCode = 22
        $script:ExitCodeDescription = 'Problem executing content of GraphConfigFile';
        exit
    }

    try { WatchCatchableExitSignal } catch { }

    if ($CheckPathCreate -eq $false) {
        Write-Verbose "      Try to access '$($CheckPathPath)'."

        if (
            -not $(
                $(
                    (((
                            [uri]$(
                                if (-not [System.Uri]::IsWellFormedUriString($CheckPathPath, [System.UriKind]::Absolute)) {
                                    $([uri]($CheckPathPath -ireplace '@SSL\\', '/' -ireplace '^\\\\', 'https://' -ireplace '\\', '/')).AbsoluteUri
                                } else {
                                    $CheckPathPath
                                }
                            )
                        ).DnsSafeHost -split '\.')[1..999] -join '.') -iin $CloudEnvironmentSharePointOnlineDomains
                ) -and
                $GraphClientID
            ) -and
            $(Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue)
        ) {
            Write-Verbose "        '$($CheckPathPath)' is accessible, nothing more to do."
        } else {
            Write-Verbose "        '$($CheckPathPath)' is not yet accessible."

            if (-not [System.Uri]::IsWellFormedUriString($CheckPathPath, [System.UriKind]::Absolute)) {
                $CheckPathPath = ([uri]($CheckPathPath -ireplace '@SSL\\', '/' -ireplace '^\\\\', 'https://' -ireplace '\\', '/')).AbsoluteUri
            }

            if (
                (((([uri]$CheckPathPath).DnsSafeHost -split '\.')[1..999] -join '.') -iin $CloudEnvironmentSharePointOnlineDomains) -and
                $GraphClientID
            ) {
                # SharePoint Online with Graph client ID
                if (-not $CheckPathSilent) {
                    Write-Host '    SharePoint via Graph, may be slow'
                }

                $CheckPathPath = [uri]::UnescapeDataString($CheckPathPath.Trimend('/'))
                $CheckPathPathSplitBySlash = @($CheckPathPath -split '\/' | Where-Object { $_ })

                try { WatchCatchableExitSignal } catch { }

                # graph auth
                if (-not $GraphToken) {
                    try {
                        $GraphToken = GraphGetToken
                    } catch {
                        $GraphToken = $null
                    }

                    if ($GraphToken.error -eq $false) {
                        Write-Verbose "      Graph Token metadata: $((ParseJwtToken $GraphToken.AccessToken) | ConvertTo-Json)"

                        if ($SimulateAndDeployGraphCredentialFile) {
                            Write-Verbose "      App Graph Token metadata: $((ParseJwtToken $GraphToken.AppAccessToken) | ConvertTo-Json)"
                        }
                    } else {
                        Write-Host '      Problem connecting to Microsoft Graph. Exit.' -ForegroundColor Red
                        Write-Host $GraphToken.error -ForegroundColor Red
                        $script:ExitCode = 23
                        $script:ExitCodeDescription = 'Problem connecting to Microsoft Graph.';
                        exit
                    }
                }

                if ($SimulateUser) {
                    $script:GraphUser = $SimulateUser
                }

                try { WatchCatchableExitSignal } catch { }

                Write-Verbose '    Get SharePoint Online site ID'

                $(
                    if ($CheckPathPathSplitbySlash[2] -iin @('sites', 'teams')) {
                        "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/sites/$(([uri]$CheckPathPath).DnsSafeHost):/$($CheckPathPathSplitbySlash[2])/$($CheckPathPathSplitbySlash[3])"
                    } else {
                        "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/sites/$(([uri]$CheckPathPath).DnsSafeHost)"
                    }
                ) | ForEach-Object {
                    Write-Verbose "      Query: '$($_)'"

                    $siteId = (GraphGenericQuery -method 'Get' -uri $_).result.id

                    Write-Verbose "      siteId: $($siteID)"
                }


                try { WatchCatchableExitSignal } catch { }

                if ($siteid) {
                    Write-Verbose '    Get DocLib drive ID'

                    "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/sites/$($siteId)/drives" | ForEach-Object {
                        $docLibDriveIdQueryResult = (GraphGenericQuery -method 'Get' -uri $_).result.value
                        $docLibDriveId = ($docLibDriveIdQueryResult | Where-Object {
                                $_.webUrl -ieq $(
                                    if ($CheckPathPathSplitbySlash[2] -iin @('sites', 'teams')) {
                                        [uri]::EscapeUriString($(($CheckPathPath -split '/')[0..5] -join '/'))
                                    } else {
                                        [uri]::EscapeUriString($(($CheckPathPath -split '/')[0..3] -join '/'))
                                    }
                                )
                            }
                        ).id

                        Write-Verbose "      Query: '$($_)'"
                        Write-Verbose "      Return value: '$(ConvertTo-Json $docLibDriveIdQueryResult -Compress -Depth 10)'"
                        Write-Verbose "      webUrl: '$([uri]::EscapeUriString($(($CheckPathPath -split '/')[0..5] -join '/')))'"

                        Write-Verbose "      docLibDriveId: $docLibDriveId"
                    }

                    try { WatchCatchableExitSignal } catch { }

                    if ($docLibDriveId) {
                        Write-Verbose '      Get DocLib drive items'
                        $docLibDriveItems = (GraphGenericQuery -method 'Get' -uri "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/drives/$($docLibDriveId)/list/items?`$expand=DriveItem").result.value

                        $tempDir = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))
                        $null = New-Item $tempDir -ItemType Directory

                        switch (($docLibDriveItems | Where-Object { ([uri]($_.webUrl)).AbsoluteUri -eq ([uri]($CheckPathPath)).AbsoluteUri }).contentType.name) {
                            'document' {
                                Write-Verbose '    Download file to local temp folder'

                                $CheckPathPathNew = $(Join-Path -Path $tempDir -ChildPath $([uri]::UnEscapeDataString((Split-Path $($docLibDriveItems | Where-Object { ([uri]($_.webUrl)).AbsoluteUri -eq ([uri]($CheckPathPath)).AbsoluteUri }).webUrl -Leaf))))

                                $(New-Object Net.WebClient).DownloadFile(
                                    $($docLibDriveItems | Where-Object { ([uri]($_.webUrl)).AbsoluteUri -eq ([uri]($CheckPathPath)).AbsoluteUri }).driveItem.'@microsoft.graph.downloadUrl',
                                    $CheckPathPathNew
                                )

                                Write-Verbose "      '$($CheckPathRefPath.Value)' -> '$($CheckPathPathNew)'"
                                $CheckPathPath = $CheckPathRefPath.Value = $CheckPathPathNew

                                break
                            }

                            'folder' {
                                Write-Verbose '    Create temp folders locally'

                                @(
                                    @($docLibDriveItems | Where-Object { ($_.contentType.name -ieq 'Folder') -and ($_.webUrl -ilike "$([uri]::EscapeUriString($CheckPathPath))/*") }).webUrl | ForEach-Object {
                                        [uri]::UnescapeDataString(($_ -ireplace "^$([uri]::EscapeUriString($CheckPathPath))/", '')) -replace '/', '\'
                                    }
                                ) | Sort-Object | ForEach-Object {
                                    if (-not (Test-Path (Join-Path -Path $tempDir -ChildPath $_) -PathType Container)) {
                                        $null = New-Item -ItemType Directory -Path (Join-Path -Path $tempDir -ChildPath $_)
                                    }
                                }

                                Write-Verbose '      Create dummy files in local temp folders'
                                @($docLibDriveItems | Where-Object { ($_.contentType.name -ieq 'Document') -and ($_.webUrl -ilike "$([uri]::EscapeUriString($CheckPathPath))/*") }) | Sort-Object -Property { $_.webUrl } | ForEach-Object {
                                    $CheckPathPathNew = $(Join-Path -Path $tempDir -ChildPath ([uri]::UnescapeDataString(($_.webUrl -ireplace "^$([uri]::EscapeUriString($CheckPathPath))/", '')) -replace '/', '\'))

                                    if (-not $script:SpoDownloadUrls) {
                                        $script:SpoDownloadUrls = @{}
                                    }

                                    $script:SpoDownloadUrls.Add(
                                        $CheckPathPathNew,
                                        $_.driveItem.'@microsoft.graph.downloadUrl'
                                    )

                                    $null = New-Item -Path $CheckPathPathNew -ItemType File
                                }

                                Write-Verbose "      '$($CheckPathRefPath.Value)' -> '$($tempDir)'"
                                $CheckPathPath = $CheckPathRefPath.Value = $tempDir

                                break
                            }

                            default {
                                Write-Host " '$($CheckPathPath)' does not exist. Exiting." -ForegroundColor Red
                                $script:ExitCode = 24
                                $script:ExitCodeDescription = "Path '$($CheckPathPath)' does not exist.";
                                exit
                            }
                        }
                    } else {
                        Write-Host '    SharePoint via Graph: No DriveID. Wrong path or missing permission in SharePoint?' -ForegroundColor Yellow
                    }
                } else {
                    Write-Host '    SharePoint via Graph: No SiteID. Wrong path or missing permission in Entra ID app?' -ForegroundColor Yellow
                }
            }

            try { WatchCatchableExitSignal } catch { }

            if ((Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue)) {
                Write-Verbose "      '$($CheckPathPath)' is accessible, nothing more to do."
            } else {
                # SharePoint Online without Graph client ID or SharePoint on-prem
                # Or normal file path that does not exist

                if ($IsWindows) {
                    # Windows. Use old way with "net use", Internet-Explorer-Cookie.

                    if (($CheckPathPath.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($CheckPathPath -ilike '*@SSL\*')) {
                        Write-Host '    SharePoint via WebDAV, may be slow and path length problems may occur (fully qualified file names must be less than 260 characters).' -ForegroundColor Yellow
                        $CheckPathPath = $CheckPathPath -ireplace '@SSL\\', '\'
                        $CheckPathPath = ([uri]::UnescapeDataString($CheckPathPath) -ireplace ('https://', '\\'))
                        $CheckPathPath = ([System.URI]$CheckPathPath).AbsoluteURI -ireplace 'file:\/\/(.*?)\/(.*)', '\\${1}@SSL\$2' -ireplace '/', '\'
                        $CheckPathPath = [uri]::UnescapeDataString($CheckPathPath)
                    } else {
                        try {
                            $CheckPathPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CheckPathPath)
                            $CheckPathPath = ([System.URI]$CheckPathPath).absoluteuri -ireplace 'file:///', '' -ireplace 'file://', '\\' -ireplace '/', '\'
                            $CheckPathPath = [uri]::UnescapeDataString($CheckPathPath)
                        } catch {
                            if ($CheckPathSilent -eq $false) {
                                Write-Host "Problem connecting or reading '$CheckPathPath'. Exit." -ForegroundColor Red
                                $script:ExitCode = 25
                                $script:ExitCodeDescription = "Problem connecting or reading '$CheckPathPath'.";
                                exit
                            }
                        }
                    }

                    if (-not (Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue)) {
                        # Reconnect already connected network drives at the OS level
                        # New-PSDrive is not enough for this
                        foreach ($NetworkConnection in @(Get-CimInstance Win32_NetworkConnection)) {
                            try { WatchCatchableExitSignal } catch { }
                            & net use $NetworkConnection.LocalName $NetworkConnection.RemoteName 2>&1 | Out-Null
                        }

                        if (-not (Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue)) {
                            try { WatchCatchableExitSignal } catch { }

                            # Connect network drives
                            $([System.Environment]::NewLine) | & net use "$CheckPathPath" 2>&1 | Out-Null

                            try { WatchCatchableExitSignal } catch { }

                            try {
                                (Test-Path -LiteralPath $CheckPathPath -ErrorAction Stop) | Out-Null
                            } catch {
                                if ($_.CategoryInfo.Category -eq 'PermissionDenied') {
                                    try { WatchCatchableExitSignal } catch { }
                                    & net use "$CheckPathPath" 2>&1
                                }
                            }

                            try { WatchCatchableExitSignal } catch { }

                            & net use "$CheckPathPath" /d 2>&1 | Out-Null
                        }

                        try { WatchCatchableExitSignal } catch { }

                        if (($CheckPathPath -ilike '*@SSL\*') -and (-not (Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue))) {
                            if ((Get-Service -ServiceName 'WebClient' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Status -ine 'Running') {
                                if (-not $CheckPathSilent) {
                                    Write-Host
                                    Write-Host 'WebClient service not running.' -ForegroundColor Red
                                }
                            } else {
                                Try {
                                    if (-not [string]::IsNullOrWhitespace($GraphHtmlMessageboxText)) {
                                        if ($IsWindows -and (-not (Test-Path env:SSH_CLIENT))) {
                                            Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

                                            $window = New-Object System.Windows.Window -Property @{
                                                Width                 = 1
                                                Height                = 1
                                                WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
                                                ShowActivated         = $false
                                                Topmost               = $true
                                            }

                                            $window.Show()
                                            $window.Hide()

                                            $MessageBoxResult = [System.Windows.MessageBox]::Show($window, "$($GraphHtmlMessageboxText)", $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }), [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Information, [System.Windows.MessageBoxResult]::None)

                                            if ($MessageBoxResult -ieq 'Cancel') {
                                                $window.Close()

                                                Write-Host
                                                Write-Host 'Authentication cancelled by user. Exiting.' -ForegroundColor Red

                                                $script:ExitCode = 26
                                                $script:ExitCodeDescription = 'Authentication cancelled by user.';
                                                exit
                                            }

                                            $window.Close()
                                        }
                                    }

                                    # Add site to trusted sites in internet options
                                    New-Item ('HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\' + (New-Object System.Uri -ArgumentList ($CheckPathPath -ireplace ('@SSL', ''))).Host) -Force | New-ItemProperty -Name * -Value 1 -Type DWORD -Force | Out-Null

                                    # Open site in new IE process
                                    $oIE = New-Object -com InternetExplorer.Application
                                    $oIE.Navigate('https://' + ((($CheckPathPath -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')) + '?web=1')
                                    $oIE.Visible = $true

                                    # Wait until an IE tab with the corresponding URL is open
                                    $app = New-Object -com shell.application

                                    $i = 0

                                    $compareurl = ('*' + ([uri]::UnescapeDataString([uri]::UnescapeDataString((($CheckPathPath -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))) + '*') -split '\/' -join '*'

                                    while ($i -lt 1) {
                                        $i += @($app.windows() | Where-Object {
                                            ([uri]::UnescapeDataString([uri]::UnescapeDataString($_.LocationURL)) -ilike $compareurl)
                                            }).count

                                        Start-Sleep -Seconds 1

                                        try { WatchCatchableExitSignal } catch { }
                                    }

                                    # Wait until the corresponding URL is fully loaded, then close the tab
                                    @($app.windows() | Where-Object {
                                        ([uri]::UnescapeDataString([uri]::UnescapeDataString($_.LocationURL)) -ilike $compareurl)
                                        }) | ForEach-Object {

                                        while ($_.Busy) {
                                            Start-Sleep -Milliseconds 100

                                            try { WatchCatchableExitSignal } catch { }
                                        }

                                        $_.Quit()
                                    }

                                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
                                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($oIE) | Out-Null

                                    Remove-Variable -Name 'app'
                                    Remove-Variable -Name 'oIE'
                                } catch {
                                    $_
                                }
                            }
                        }
                    }
                } else {
                    if (($CheckPathPath.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($CheckPathPath -ilike '*@SSL\*')) {
                        Write-Host '    SharePoint via WebDAV is only supported on Windows platforms.' -ForegroundColor Yellow
                    }
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        if ((Test-Path -LiteralPath $CheckPathPath -PathType $ExpectedPathType) -eq $false) {
            if ($CheckPathSilent -eq $false) {
                Write-Host "Problem connecting or reading $($ExpectedPathType) '$($CheckPathPath)'. Exit." -ForegroundColor Red
                $script:ExitCode = 27
                $script:ExitCodeDescription = "Problem connecting or reading $($ExpectedPathType) '$($CheckPathPath)'.";
                exit
            } else {
                return $false
            }
        } else {
            if ($CheckPathSilent -eq $false) {
                # Write-Host
            } else {
                return $true
            }
        }
    } else {
        Write-Verbose "      Try to create '$($CheckPathPath)'."

        if ($CheckPathPath.StartsWith('https://', 'CurrentCultureIgnoreCase')) {
            $CheckPathPath = ((([uri]::UnescapeDataString($CheckPathPath) -ireplace ('https://', '\\')) -ireplace ('(.*?)/(.*)', '${1}@SSL\$2')) -ireplace ('/', '\'))
        } else {
            # '@SSL' seems to be case sensitive, so we make sure that the first occurrence is in uppercase letters
            $CheckPathPath = ([regex]"(?i)$([regex]::escape('@ssl\'))").replace($CheckPathPath, '@SSL\', 1)
        }

        $CheckPathPathTarget = $CheckPathPath

        for (
            $i = 1
            $i -lt @($CheckPathPathTarget -split [regex]::escape([IO.Path]::DirectorySeparatorChar)).count
            $i++
        ) {
            try { WatchCatchableExitSignal } catch { }

            $CheckPathPathTemp = @($CheckPathPathTarget -split [regex]::escape([IO.Path]::DirectorySeparatorChar))[0..$i] -join [IO.Path]::DirectorySeparatorChar

            if ((. $CheckPathScriptblock ([ref]$CheckPathPathTemp) -CheckPathSilent) -eq $true) {
                if (-not (Test-Path $CheckPathPathTemp -PathType Container -ErrorAction SilentlyContinue)) {
                    Write-Host "'$CheckPathPathTemp' is a file, '$CheckPathPathTarget' is not valid. Exit." -ForegroundColor Red
                    $script:ExitCode = 28
                    $script:ExitCodeDescription = "'$CheckPathPathTemp' is a file, '$CheckPathPathTarget' is not valid.";
                    exit
                }

                if ($CheckPathPathTemp -eq $CheckPathPathTarget) {
                    break
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    Write-Verbose "      Try to create '$($CheckPathPathTarget)'."

                    New-Item -ItemType Directory -Path $CheckPathPathTarget -ErrorAction SilentlyContinue | Out-Null

                    if (Test-Path -Path $CheckPathPathTarget -PathType Container) {
                        break
                    }
                }
            }
        }

        if ((. $CheckPathScriptblock ([ref]$CheckPathPathTarget) -CheckPathSilent) -ne $true) {
            Write-Host "Problem connecting or reading '$CheckPathPathTarget'. Exit." -ForegroundColor Red
            $script:ExitCode = 29
            $script:ExitCodeDescription = "Problem connecting or reading '$CheckPathPathTarget'.";
            exit
        } else {
            # Write-Host
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function ConnectEWS([string]$MailAddress = $MailAddresses[0], [string]$Indent = '') {
    try { WatchCatchableExitSignal } catch { }

    Write-Host "$($Indent)Connect to Outlook Web"

    $local:exchServiceAvailable = $false

    if ($script:exchService) {
        try {
            if (
                $($script:exchService.SetOutlookSignaturesMailaddress -ieq $MailAddress) -and
                $(([Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:exchService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)).DisplayName)
            ) {
                Write-Host "$($Indent)  Existing connection matches required parameters and is working"

                $local:exchServiceAvailable = $true
            } else {
                Write-Host "$($Indent)  Existing connecting does not match required parameters or does not work"
            }
        } catch {
            Write-Host "$($Indent)  Existing connecting does not match required parameters or does not work"
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ($local:exchServiceAvailable -eq $false) {
        Write-Host "$($Indent)  Creating new connection"

        $script:exchService = $null

        try {
            Import-Module -Name $script:WebServicesDllPath -Force -ErrorAction Stop

            try { WatchCatchableExitSignal } catch { }

            $script:exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

            try { WatchCatchableExitSignal } catch { }

            $tempEwsRedirectUrl = $null

            function ExchServiceEwsTraceHandler() {
                $sourceCode = @'
using System.Management.Automation;
using System.Text;
using System.Text.RegularExpressions;

public class ExchServiceEwsTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
{
    public void Trace(System.String traceType, System.String traceMessage)
    {
        string tempEwsRedirectUrl = string.Empty;

        Match match = Regex.Match(traceMessage, "Redirection URL found: '(.*?)'");

        if (match.Success)
        {
            tempEwsRedirectUrl = match.Groups[1].Value;
        }

        StringBuilder sb = new StringBuilder();
        // sb.AppendLine("Write-Verbose \"$($Indent)      traceType: $($('" + System.Management.Automation.Language.CodeGeneration.EscapeSingleQuotedStringContent(traceType) + "'))\"");
        // sb.AppendLine("Write-Verbose \"$($Indent)      traceMessage: $($('" + System.Management.Automation.Language.CodeGeneration.EscapeSingleQuotedStringContent(traceMessage) + "'))\"");
        sb.AppendLine("$tempEwsRedirectUrl = $($('" + System.Management.Automation.Language.CodeGeneration.EscapeSingleQuotedStringContent(tempEwsRedirectUrl) + "'))");

        var defRunspace = System.Management.Automation.Runspaces.Runspace.DefaultRunspace;
        var pipeline = defRunspace.CreateNestedPipeline();
        pipeline.Commands.AddScript(sb.ToString());
        pipeline.Invoke();
    }
}
'@

                Add-Type -TypeDefinition $sourceCode -Language CSharp -ReferencedAssemblies $script:WebServicesDllPath, System.Management.Automation, System.Text.RegularExpressions
                $ExchServiceEwsTraceListener = New-Object ExchServiceEwsTraceListener
                return $ExchServiceEwsTraceListener
            }

            $script:exchService.TraceEnabled = $true
            $script:exchService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::AutodiscoverConfiguration, [Microsoft.Exchange.WebServices.Data.TraceFlags]::AutodiscoverRequest, [Microsoft.Exchange.WebServices.Data.TraceFlags]::AutodiscoverResponse
            $script:exchService.TraceListener = ExchServiceEwsTraceHandler

            try { WatchCatchableExitSignal } catch { }

            try {
                Write-Verbose "$($Indent)    Try Autodiscover with Integrated Windows Authentication"

                $script:exchService.UseDefaultCredentials = $true
                $script:exchService.ImpersonatedUserId = $null
                $script:exchService.AutodiscoverUrl($MailAddress, { $true }) | Out-Null

            } catch {
                Write-Verbose "$($Indent)      Autodiscover with Integrated Windows Authentication failed."
                Write-Verbose "$($Indent)        $($_)"
                Write-Verbose "$($Indent)      This is OK when:"
                Write-Verbose "$($Indent)        - Not connected to internal network"
                Write-Verbose "$($Indent)        - Connected to internal network with no Exchange on prem."
                Write-Verbose "$($Indent)        - Connected to internal network with Exchange on prem, but your mailbox is in Exchange Online."
                Write-Verbose "$($Indent)        - Connected to internal network but not logged-on with Active Directory credentials."
                Write-Verbose "$($Indent)      Else, you should check your internal and/or external Autodiscover configuration:"
                Write-Verbose "$($Indent)        - External: https://testconnectivity.microsoft.com"
                Write-Verbose "$($Indent)        - Internal: https://learn.microsoft.com/en-us/exchange/architecture/client-access/autodiscover"
                Write-Verbose "$($Indent)        - Check your loadbalancer configuration."

                if ([System.Uri]::IsWellFormedUriString($tempEwsRedirectUrl, [System.UriKind]::Absolute)) {
                    Write-Verbose "$($Indent)      Anyhow:"
                    Write-Verbose "$($Indent)        - Redirect URL '$($tempEwsRedirectUrl)' was returned."
                    Write-Verbose "$($Indent)        - No need to try Autodiscver with OAuth, skipping to OAuth with fixed URL."
                } else {
                    $tempEwsRedirectUrl = $null
                }

                if (
                    $($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile -and !$GraphToken.AppAccessTokenExo) -or
                    !$GraphToken.AccessTokenExo
                ) {
                    throw "Integrated Windows Authentication failed, and there is no EXO OAuth access token available. Did you forget '-GraphOnly true' or are you missing AD attributes?"
                }

                try { WatchCatchableExitSignal } catch { }

                try {
                    Write-Verbose "$($Indent)    Try Autodiscover with OAuth"

                    if ([System.Uri]::IsWellFormedUriString($tempEwsRedirectUrl, [System.UriKind]::Absolute)) {
                        throw 'Autodiscover with IWA failed before, but returned a redirect URL. We will use this fixed URL without Autodiscover.'
                    } else {
                        $tempEwsRedirectUrl = $null
                    }

                    $script:exchService.UseDefaultCredentials = $false

                    if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) {
                        $script:exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailAddress)
                        $script:exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AppAccessTokenExo)
                    } else {
                        $script:exchService.ImpersonatedUserId = $null
                        $script:exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AccessTokenExo)
                    }

                    $script:exchService.AutodiscoverUrl($MailAddress, { $true }) | Out-Null
                } catch {
                    if ([System.Uri]::IsWellFormedUriString($tempEwsRedirectUrl, [System.UriKind]::Absolute)) {
                        Write-Verbose "$($Indent)      Skipping Autodiscover with OAuth because"
                        Write-Verbose "$($Indent)        $($_)"
                    } else {
                        Write-Verbose "$($Indent)      Autodiscover with OAuth failed."
                        Write-Verbose "$($Indent)        $($_)"
                        Write-Verbose "$($Indent)      This is OK when"
                        Write-Verbose "$($indent)        - Connected to internal network with Exchange on prem without Hybrid Modern Authentication"
                        Write-Verbose "$($Indent)      Else, you should check your internal and/or external Autodiscover configuration:"
                        Write-Verbose "$($Indent)        - External: https://testconnectivity.microsoft.com"
                        Write-Verbose "$($Indent)        - Internal: https://learn.microsoft.com/en-us/exchange/architecture/client-access/autodiscover"
                        Write-Verbose "$($Indent)        - Check your loadbalancer configuration."
                    }

                    try { WatchCatchableExitSignal } catch { }

                    Write-Verbose "$($Indent)    Try OAuth with fixed URL"

                    $script:exchService.UseDefaultCredentials = $false

                    if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) {
                        $script:exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailAddress)
                        $script:exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AppAccessTokenExo)
                    } else {
                        $script:exchService.ImpersonatedUserId = $null
                        $script:exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AccessTokenExo)
                    }

                    if ([System.Uri]::IsWellFormedUriString($tempEwsRedirectUrl, [System.UriKind]::Absolute)) {
                        $script:exchService.Url = "$(([uri]$tempEwsRedirectUrl).GetLeftPart([UriPartial]::Authority))/EWS/Exchange.asmx"
                    } else {
                        $script:exchService.Url = "$($CloudEnvironmentExchangeOnlineEndpoint)/EWS/Exchange.asmx"
                    }

                    Write-Verbose "$($Indent)      Fixed URL: '$($script:exchService.Url)'"
                }
            }

            if (([Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:exchService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)).DisplayName) {
                Add-Member -InputObject $script:exchService -MemberType NoteProperty -Name 'SetOutlookSignaturesMailaddress' -Value $MailAddress -Force
            } else {
                throw 'Could not connect to Outlook Web, although the EWS DLL threw no error.'
            }
        } catch {
            Write-Host "$($Indent)    Error connecting to Outlook Web: $($_)" -ForegroundColor Red
            Write-Host "$($Indent)    Check verbose output for details and solution hints." -ForegroundColor Red

            $script:exchService = $null
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function GraphGenericQuery {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory = $true)]
        [string]$method,

        [Parameter(Mandatory = $true)]
        [uri]$uri
    )

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = $method
            Uri         = $uri
            Headers     = $script:AuthorizationHeader
            ContentType = 'application/json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error  = $error[0] | Out-String
            result = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error  = $false
            result = $local:x
        }
    } else {
        return @{
            error  = $error[0] | Out-String
            result = $null
        }
    }
}


function GraphGetToken {
    param(
        [switch]$EXO
    )

    try { WatchCatchableExitSignal } catch { }

    if (-not $EXO) {
        Write-Host '    Graph authentication'
    }


    try {
        Invoke-WebRequest $CloudEnvironmentGraphApiEndpoint -UseBasicParsing -TimeoutSec 5
    } catch {
        return @{
            error             = "Endpoint '$($CloudEnvironmentGraphApiEndpoint)' is not accessible: $($_)"
            AccessToken       = $null
            authHeader        = $null
            AccessTokenExo    = $null
            authHeaderExo     = $null
            AppAccessToken    = $null
            AppAuthHeader     = $null
            AppAccessTokenExo = $null
            AppAuthHeaderExo  = $null
        }
    }

    if (-not $EXO) {
        try {
            Invoke-WebRequest $CloudEnvironmentAzureADEndpoint -UseBasicParsing -TimeoutSec 5
        } catch {
            return @{
                error             = "Endpoint '$($CloudEnvironmentAzureADEndpoint)' is not accessible: $($_)"
                AccessToken       = $null
                authHeader        = $null
                AccessTokenExo    = $null
                authHeaderExo     = $null
                AppAccessToken    = $null
                AppAuthHeader     = $null
                AppAccessTokenExo = $null
                AppAuthHeaderExo  = $null
            }
        }
    } else {
        try {
            Invoke-WebRequest $CloudEnvironmentExchangeOnlineEndpoint -UseBasicParsing -TimeoutSec 5
        } catch {
            return @{
                error             = "Endpoint '$($CloudEnvironmentExchangeOnlineEndpoint)' is not accessible: $($_)"
                AccessToken       = $null
                authHeader        = $null
                AccessTokenExo    = $null
                authHeaderExo     = $null
                AppAccessToken    = $null
                AppAuthHeader     = $null
                AppAccessTokenExo = $null
                AppAuthHeaderExo  = $null
            }
        }
    }

    if ($SimulateAndDeployGraphCredentialFile) {
        Write-Host "        Via SimulateAndDeployGraphCredentialFile '$SimulateAndDeployGraphCredentialFile'"

        try {
            try {
                $auth = Import-Clixml -Path $SimulateAndDeployGraphCredentialFile
            } catch {
                Start-Sleep -Seconds 2
                $auth = Import-Clixml -Path $SimulateAndDeployGraphCredentialFile
            }

            $script:AuthorizationToken = $auth.AccessToken

            $script:ExoAuthorizationToken = $auth.AccessTokenExo

            $script:AuthorizationHeader = @{
                Authorization = $auth.AuthHeader
            }

            $script:ExoAuthorizationHeader = @{
                Authorization = $auth.AuthHeaderExo
            }

            $script:AppAuthorizationHeader = @{
                Authorization = $auth.AppAuthHeader
            }

            $script:AppExoAuthorizationHeader = @{
                Authorization = $auth.AppAuthHeaderExo
            }

            return @{
                error             = $false
                AccessToken       = $auth.AccessToken
                AuthHeader        = $auth.authHeader
                AccessTokenExo    = $auth.AccessTokenExo
                AuthHeaderExo     = $auth.AuthHeaderExo
                AppAccessToken    = $auth.AppAccessToken
                AppAuthHeader     = $auth.AppAuthHeader
                AppAccessTokenExo = $auth.AppAccessTokenExo
                AppAuthHeaderExo  = $auth.AppAuthHeaderExo
            }
        } catch {
            return @{
                error             = ($error[0] | Out-String)
                AccessToken       = $null
                AuthHeader        = $null
                AccessTokenExo    = $null
                AuthHeaderExo     = $null
                AppAccessToken    = $null
                AppAuthHeader     = $null
                AppAccessTokenExo = $null
                AppAuthHeaderExo  = $null
            }
        }
    } else {
        if (-not  $script:MsalModulePath) {
            Write-Host '      Load MSAL.PS'

            $script:MsalModulePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))
            Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\MSAL.PS')) -Destination (Join-Path -Path $script:MsalModulePath -ChildPath 'MSAL.PS') -Recurse

            if (-not $IsLinux) {
                Get-ChildItem $script:MsalModulePath -Recurse | Unblock-File
            }

            try { WatchCatchableExitSignal } catch { }

            try {
                Import-Module (Join-Path -Path $script:MsalModulePath -ChildPath 'MSAL.PS') -Force -ErrorAction Stop
            } catch {
                Write-Host $error[0]
                Write-Host '        Problem importing MSAL.PS module. Exit.' -ForegroundColor Red
                $script:ExitCode = 30
                $script:ExitCodeDescription = 'Problem importing MSAL.PS module.';
                exit
            }
        }

        try { WatchCatchableExitSignal } catch { }

        # On Linux/macOS, unlock keyring/keychain if required
        if (-not [string]::IsNullOrWhitespace($GraphUnlockKeyringKeychainMessageboxText)) {
            if ($IsLinux) {
                $keyringPath = (dbus-send --session --dest=org.freedesktop.secrets --type=method_call --print-reply /org/freedesktop/secrets org.freedesktop.Secret.Service.ReadAlias string:'default' | grep -oP '(?<=object path \")/[^"]+')

                if ($((gdbus call -e -d org.freedesktop.secrets -o $keyringPath -m org.freedesktop.DBus.Properties.Get org.freedesktop.Secret.Collection Locked *>&1) -ine '(<false>,)')) {
                    if ($(Get-Command -Name 'kdialog' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                        $null = kdialog `
                            --title $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }) `
                            --msgbox "$($GraphUnlockKeyringKeychainMessageboxText)"
                    } elseif ($(Get-Command -Name 'zenity' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                        $null = zenity `
                            --info `
                            --title=$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }) `
                            --text="$($GraphUnlockKeyringKeychainMessageboxText)"
                    } else {
                        Write-Host "        Neither kdialog nor zenity found, so no message box could be shown: $($GraphUnlockKeyringKeychainMessageboxText)"
                    }
                }
            } elseif ($IsMacOS) {
                security unlock-keychain -p 'Set-OutlookSignatures dummy password' *>$null

                if ($LastExitCode -ne 0) {
                    Write-Host $("display alert ""$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })"" message ""$($GraphUnlockKeyringKeychainMessageboxText)""  buttons { ""OK"" } default button 1" | osascript *>$1; '')
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        try {
            Write-Host '      Search for login hint in Graph token cache'

            $script:GraphUser = $null

            $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId 'organizations' -AuthenticationBroker | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

            $script:GraphUser = ($script:msalClientApp | get-msalaccount | Select-Object -First 1).username

            Write-Host "        Graph token cache: $($script:msalClientApp.cacheInfo)"
            Write-Host "        Result: '$($script:GraphUser)'"
        } catch {
            return @{
                error             = ($error[0] | Out-String)
                AccessToken       = $null
                AuthHeader        = $null
                AccessTokenExo    = $null
                AuthHeaderExo     = $null
                AppAccessToken    = $null
                AppAuthHeader     = $null
                AppAccessTokenExo = $null
                AppAuthHeaderExo  = $null
            }
        }

        try { WatchCatchableExitSignal } catch { }

        # Graph authentication
        Write-Host "      Authentication against $(if(-not $EXO) { 'Graph' } else { 'Exchange Online' })"

        try {
            Write-Host '        Silent via Integrated Windows Authentication without login hint'

            $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId 'organizations' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

            $auth = $script:msalClientApp | Get-MsalToken -IntegratedWindowsAuth -AzureCloudInstance $CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($CloudEnvironmentGraphApiEndpoint)/.default" }else { "$($CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -Timeout (New-TimeSpan -Minutes 1)

            Write-Host "          Success: '$(($script:msalClientApp | get-msalaccount | Select-Object -First 1).username)'"
        } catch {
            Write-Host "          Failed: $($error[0])"

            try { WatchCatchableExitSignal } catch { }

            try {
                Write-Host '        Silent via Integrated Windows Authentication with login hint'
                # Required, because IWA without login hint may fail when account enumeration is blocked at OS level

                if (-not ([string]::IsNullOrWhiteSpace($script:GraphUser))) {
                    $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId 'organizations' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                    $auth = $script:msalClientApp | Get-MsalToken -IntegratedWindowsAuth -LoginHint $script:GraphUser -AzureCloudInstance $CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($CloudEnvironmentGraphApiEndpoint)/.default" }else { "$($CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -Timeout (New-TimeSpan -Minutes 1)
                } else {
                    throw 'No login hint found before'
                }

                Write-Host "          Success: '$(($script:msalClientApp | get-msalaccount | Select-Object -First 1).username)'"
            } catch {
                Write-Host "          Failed: $($error[0])"

                try { WatchCatchableExitSignal } catch { }

                try {
                    Write-Host '        Silent via Authentication Broker without login hint'

                    $script:msalClientApp = New-MsalClientApplication -AuthenticationBroker -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId 'organizations' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                    $auth = $script:msalClientApp | Get-MsalToken -Silent -AuthenticationBroker -AzureCloudInstance $CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($CloudEnvironmentGraphApiEndpoint)/.default" }else { "$($CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -ForceRefresh -Timeout (New-TimeSpan -Minutes 1)

                    Write-Host "          Success: '$(($script:msalClientApp | get-msalaccount | Select-Object -First 1).username)'"
                } catch {
                    Write-Host "          Failed: $($error[0])"

                    try { WatchCatchableExitSignal } catch { }

                    try {
                        Write-Host '        Silent via Authentication Broker with login hint'

                        if (-not ([string]::IsNullOrWhiteSpace($script:GraphUser))) {
                            $script:msalClientApp = New-MsalClientApplication -AuthenticationBroker -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId 'organizations' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                            $auth = $script:msalClientApp | Get-MsalToken -Silent -AuthenticationBroker -LoginHint $script:GraphUser -AzureCloudInstance $CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($CloudEnvironmentGraphApiEndpoint)/.default" }else { "$($CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -ForceRefresh -Timeout (New-TimeSpan -Minutes 1)
                        } else {
                            throw 'No login hint found before'
                        }

                        Write-Host "          Success: '$(($script:msalClientApp | get-msalaccount | Select-Object -First 1).username)'"
                    } catch {
                        Write-Host "          Failed: $($error[0])"

                        try {
                            Write-Host '        Silent via refresh token, with login hint'

                            if (-not ([string]::IsNullOrWhiteSpace($script:GraphUser))) {
                                $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId 'organizations' -RedirectUri 'http://localhost' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                                $auth = $script:msalClientApp | Get-MsalToken -Silent -LoginHint $script:GraphUser -AzureCloudInstance $CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($CloudEnvironmentGraphApiEndpoint)/.default" }else { "$($CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -ForceRefresh -Timeout (New-TimeSpan -Minutes 1)
                            } else {
                                throw 'No login hint found before'
                            }

                            Write-Host "          Success: '$(($script:msalClientApp | get-msalaccount | Select-Object -First 1).username)'"
                        } catch {
                            Write-Host "          Failed: $($error[0])"

                            try { WatchCatchableExitSignal } catch { }

                            # Interactive authentication methods
                            Write-Host '        All silent authentication methods failed, switching to interactive authentication methods.'

                            if (-not [string]::IsNullOrWhitespace($GraphHtmlMessageboxText)) {
                                if ($IsWindows -and (-not (Test-Path env:SSH_CLIENT))) {
                                    Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

                                    $window = New-Object System.Windows.Window -Property @{
                                        Width                 = 1
                                        Height                = 1
                                        WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
                                        ShowActivated         = $false
                                        Topmost               = $true
                                    }

                                    $window.Show()
                                    $window.Hide()

                                    $MessageBoxResult = [System.Windows.MessageBox]::Show($window, "$($GraphHtmlMessageboxText)", $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }), [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Information, [System.Windows.MessageBoxResult]::None)

                                    $window.Close()

                                    if ($MessageBoxResult -ieq 'Cancel') {
                                        return @{
                                            error             = 'Authentication cancelled by user. Exiting.'
                                            AccessToken       = $null
                                            authHeader        = $null
                                            AccessTokenExo    = $null
                                            authHeaderExo     = $null
                                            AppAccessToken    = $null
                                            AppAuthHeader     = $null
                                            AppAccessTokenExo = $null
                                            AppAuthHeaderExo  = $null
                                        }
                                    }
                                } elseif ($IsLinux -and ((Test-Path env:DISPLAY))) {
                                    if ($(Get-Command -Name 'kdialog' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                                        $null = kdialog `
                                            --title $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }) `
                                            --msgbox "$($GraphHtmlMessageboxText)"
                                    } elseif ($(Get-Command -Name 'zenity' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                                        $null = zenity `
                                            --info `
                                            --title=$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }) `
                                            --text="$($GraphHtmlMessageboxText)"
                                    } else {
                                        Write-Host "          Neither kdialog nor zenity found, so no message box could be shown: $($GraphHtmlMessageboxText)"
                                    }
                                } elseif ($IsMacOS -and ((Test-Path env:DISPLAY))) {
                                    Write-Host $("display alert ""$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })"" message ""$($GraphHtmlMessageboxText)""  buttons { ""OK"" } default button 1" | osascript *>&1; '')
                                }

                                try { WatchCatchableExitSignal } catch { }
                            }

                            $MsalInteractiveParams = @{}

                            if (-not [string]::IsNullOrWhiteSpace($GraphBrowserRedirectSuccess)) {
                                $MsalInteractiveParams.BrowserRedirectSuccess = $GraphBrowserRedirectSuccess
                            }

                            if (-not [string]::IsNullOrWhiteSpace($GraphBrowserRedirectError)) {
                                $MsalInteractiveParams.BrowserRedirectError = $GraphBrowserRedirectError
                            }

                            if (-not [string]::IsNullOrWhiteSpace($GraphHtmlMessageSuccess)) {
                                $MsalInteractiveParams.HtmlMessageSuccess = $GraphHtmlMessageSuccess
                            }

                            if (-not [string]::IsNullOrWhiteSpace($GraphHtmlMessageError)) {
                                $MsalInteractiveParams.HtmlMessageError = $GraphHtmlMessageError
                            }

                            try { WatchCatchableExitSignal } catch { }

                            try {
                                Write-Host '        Interactive via Authentication Broker'

                                if (-not $IsWindows) {
                                    throw 'Interactive with Authentication Broker on Linux/macOS only works in the console. Browser is preferred for better user experience.'
                                }

                                $script:msalClientApp = New-MsalClientApplication -AuthenticationBroker -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId 'organizations' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                                Write-Host '          Opening authentication broker window and waiting for you to authenticate. Stopping script execution after five minutes.'
                                $auth = $script:msalClientApp | Get-MsalToken -Interactive -AuthenticationBroker -LoginHint $(if ($script:GraphUser) { $script:GraphUser } else { '' }) -AzureCloudInstance $CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($CloudEnvironmentGraphApiEndpoint)/.default" }else { "$($CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -Timeout (New-TimeSpan -Minutes 5) -Prompt 'NoPrompt' -UseEmbeddedWebView:$false @MsalInteractiveParams

                                Write-Host "          Success: '$(($script:msalClientApp | get-msalaccount | Select-Object -First 1).username)'"
                            } catch {
                                Write-Host "          Failed: $($error[0])"

                                try {
                                    Write-Host '        Interactive via browser'

                                    $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId 'organizations' -RedirectUri 'http://localhost' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                                    Write-Host '          Opening new browser window and waiting for you to authenticate. Stopping script execution after five minutes.'
                                    $auth = $script:msalClientApp | Get-MsalToken -Interactive -LoginHint $(if ($script:GraphUser) { $script:GraphUser } else { '' }) -AzureCloudInstance $CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($CloudEnvironmentGraphApiEndpoint)/.default" }else { "$($CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -Timeout (New-TimeSpan -Minutes 5) -Prompt 'NoPrompt' -UseEmbeddedWebView:$false @MsalInteractiveParams

                                    Write-Host "          Success: '$(($script:msalClientApp | get-msalaccount | Select-Object -First 1).username)'"
                                } catch {
                                    Write-Host "          Failed: $($error[0])"
                                    Write-Host '        No authentication possible'

                                    $auth = $null

                                    return @{
                                        error             = (($error[0] | Out-String) + @"
No authentication possible.
1. Did you follow the Quick Start Guide in '.\docs\README' and configure the Entra ID app correctly?
2. Run Set-OutlookSignatures with the "-Verbose" parameter and check for authentication messages
3. If the "Interactive" message is displayed:
   - When using an Authentication Broker (which is preferred on supported platforms):
     - Does the account picker window show up?
     - Check if authentication happens within five minutes
     - Check if your firewall or anti-malware software blocks Set-OutlookSignatures from creating a temporary listener port for localhost.
     - Check if the correct user account is selected/entered and if the authentication is successful
   - When not using an Authentication Broker (on a system without support for it, or when broker auth failed):
     - Does a browser (the system default browser, if configured) open and ask for authentication?
      - Yes:
       - Check if authentication happens within five minutes
       - Ensure that your browser does not block access to 'http://localhost', errors such as 'connection refused' point to this problem. ('https://localhost' is currently not technically feasible, see 'https://learn.microsoft.com/en-us/entra/msal/dotnet/acquiring-tokens/using-web-browsers' and 'https://learn.microsoft.com/en-us/entra/msal/dotnet/acquiring-tokens/using-web-browsers' for details)
         This is typically due to enforced redirection to HTTPS being applied to localhost. If not configured via policies: edge://net-internals/#hsts or chrome://net-internals/#hsts, delete domain security policies for localhost.
       - Check if your firewall or anti-malware software blocks Set-OutlookSignatures from creating a temporary listener port for localhost.
       - Check if the correct user account is selected/entered and if the authentication is successful
     - No:
       - Check if a default browser is set and if the PowerShell command 'start https://github.com/Set-OutlookSignatures/Set-OutlookSignatures' opens it
       - Make sure that Set-OutlookSignatures is executed in the security context of the currently logged-in user
       - Run Set-OutlookSignatures in a new PowerShell session
       - Check your anti-malware configuration (errors such as 'error sending the request' or 'connection refused' point at a problem there)
       - Make sure that the current PowerShell session allows TLS 1.2+ (see https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/85 for details)
4. Delete the Graph token cache: $($script:msalClientApp.cacheInfo).
"@)
                                        AccessToken       = $null
                                        AuthHeader        = $null
                                        AccessTokenExo    = $null
                                        AuthHeaderExo     = $null
                                        AppAccessToken    = $null
                                        AppAuthHeader     = $null
                                        AppAccessTokenExo = $null
                                        AppAuthHeaderExo  = $null
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if ($auth) {
            try {
                $script:GraphUser = $auth.account.username

                if (-not $EXO) {
                    $script:AuthorizationHeader = @{
                        Authorization = $auth.CreateAuthorizationHeader()
                    }

                    $script:AuthorizationToken = $auth.AccessToken
                } else {
                    $script:ExoAuthorizationHeader = @{
                        Authorization = $auth.CreateAuthorizationHeader()
                    }

                    $script:ExoAuthorizationToken = $auth.AccessToken
                }

                if (-not $EXO) {
                    $authExo = GraphGetToken -EXO

                    if ($authExo) {
                        return @{
                            error             = $false
                            AccessToken       = $script:AuthorizationToken
                            AuthHeader        = $script:AuthorizationHeader
                            AccessTokenExo    = $script:ExoAuthorizationToken
                            AuthHeaderExo     = $script:ExoAuthorizationHeader
                            AppAccessToken    = $null
                            AppAuthHeader     = $null
                            AppAccessTokenExo = $null
                            AppAuthHeaderExo  = $null
                        }
                    } else {
                        throw 'No Exchange Online token'
                    }
                } else {
                    return @{
                        error             = $false
                        AccessToken       = $null
                        AuthHeader        = $null
                        AccessTokenExo    = $auth.AccessToken
                        AuthHeaderExo     = $script:ExoAuthorizationHeader
                        AppAccessToken    = $null
                        AppAuthHeader     = $null
                        AppAccessTokenExo = $null
                        AppAuthHeaderExo  = $null
                    }
                }
            } catch {
                Write-Host "          Failed: $($error[0])"

                return @{
                    error             = ($error[0] | Out-String)
                    AccessToken       = $null
                    authHeader        = $null
                    AccessTokenExo    = $null
                    authHeaderExo     = $null
                    AppAccessToken    = $null
                    AppAuthHeader     = $null
                    AppAccessTokenExo = $null
                    AppAuthHeaderExo  = $null
                }
            }
        }
    }
}


function GraphGetMe {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s)
    #   Delegated: User.Read.All
    #   Application: User.Read.All (/me is not supported in applications)

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/me?`$select=" + [System.Net.WebUtility]::UrlEncode(($GraphUserProperties -join ','))
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error = $error[0] | Out-String
            me    = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error = $false
            me    = $local:x
        }
    } else {
        return @{
            error = $error[0] | Out-String
            me    = $null
        }
    }
}


function GraphGetUpnFromSmtp($user) {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users?`$filter=proxyAddresses/any(x:x eq 'smtp:$($user)')"
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error      = $error[0] | Out-String
            properties = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error      = $false
            properties = $local:x
        }
    } else {
        return @{
            error      = $error[0] | Out-String
            properties = $null
        }
    }
}


function GraphGetUserProperties($user, $authHeader = $script:AuthorizationHeader) {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    try { WatchCatchableExitSignal } catch { }

    $user = GraphGetUpnFromSmtp($user)

    if ($user.properties.value.userprincipalname) {
        try {
            $requestBody = @{
                Method      = 'Get'
                Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user.properties.value.userprincipalname)?`$select=" + [System.Net.WebUtility]::UrlEncode($(@($GraphUserProperties | Select-Object -Unique) -join ','))
                Headers     = $authHeader
                ContentType = 'Application/Json; charset=utf-8'
            }

            $OldProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            $local:x = @()
            $local:uri = $null

            do {
                try { WatchCatchableExitSignal } catch { }

                if ($local:uri) {
                    $requestBody['Uri'] = $local:uri
                }

                $local:pagedResults = Invoke-RestMethod @requestBody
                $local:x += $local:pagedResults

                if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                    $local:uri = $null
                } else {
                    $local:uri = $local:pagedResults.'@odata.nextlink'
                }
            } until (!($local:uri))


            if (($user.properties.value.userprincipalname -ieq $script:GraphUser) -and ((-not $SimulateUser) -or ($SimulateUser -and $SimulateAndDeployGraphCredentialFile -and ($authHeader -eq $script:AppAuthorizationHeader))) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true) -or ($MirrorCloudSignatures -eq $true))) {
                try {
                    $requestBody = @{
                        Method      = 'Get'
                        Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user.properties.value.userprincipalname)?`$select=mailboxsettings"
                        Headers     = $authHeader
                        ContentType = 'Application/Json; charset=utf-8'
                    }

                    $OldProgressPreference = $ProgressPreference
                    $ProgressPreference = 'SilentlyContinue'

                    $local:y = @()

                    do {
                        try { WatchCatchableExitSignal } catch { }

                        if ($local:uri) {
                            $requestBody['Uri'] = $local:uri
                        }

                        $local:pagedResults = Invoke-RestMethod @requestBody
                        $local:y += $local:pagedResults

                        if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                            $local:uri = $null
                        } else {
                            $local:uri = $local:pagedResults.'@odata.nextlink'
                        }
                    } until (!($local:uri))

                    $local:x | Add-Member -MemberType NoteProperty -Name 'mailboxSettings' -Value $local:y.mailboxSettings -Force
                } catch {
                    Write-Host $error[0]
                    Write-Host "      Problem getting mailboxSettings for '$($script:GraphUser)' from Microsoft Graph." -ForegroundColor Yellow
                    Write-Host '      This is a Microsoft Graph API problem, which can only be solved by Microsoft itself.' -ForegroundColor Yellow
                    Write-Host '      Disabling SetCurrentUserOutlookWebSignature and SetCurrentUserOOFMessage to be able to continue.' -ForegroundColor Yellow

                    $SetCurrentUserOutlookWebSignature = $false
                    $SetCurrentUserOOFMessage = $false
                }
            }

            $ProgressPreference = $OldProgressPreference
        } catch {
            return @{
                error      = $error[0] | Out-String
                properties = $null
            }
        }

        try { WatchCatchableExitSignal } catch { }

        if (($user.properties.value.userprincipalname -ieq $script:GraphUser) -and ($SimulateUser -and $SimulateAndDeployGraphCredentialFile -and ($authHeader -eq $script:AuthorizationHeader))) {
            $temp = GraphGetUserProperties -user $($user.properties.value.userprincipalname) -authHeader $script:AppAuthorizationHeader

            if ($temp.error -eq $false) {
                $local:x = $temp.properties
            } else {
            }
        }

        if ($null -ne $local:x) {
            return @{
                error      = $false
                properties = $local:x
            }
        } else {
            return @{
                error      = $error[0] | Out-String
                properties = $null
            }
        }
    } else {
        return @{
            error      = $user.error
            properties = $null
        }
    }
}


function GraphGetUserManager($user) {
    # Current mailbox manager
    # https://docs.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user)/manager"
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error      = $error[0] | Out-String
            properties = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error      = $false
            properties = $local:x
        }
    } else {
        return @{
            error      = $error[0] | Out-String
            properties = $null
        }
    }

}


function GraphGetUserTransitiveMemberOf($user) {
    # https://learn.microsoft.com/en-us/graph/api/user-list-transitivememberof?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user)/transitiveMemberOf"
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error    = $error[0] | Out-String
            memberof = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error    = $false
            memberof = $local:x
        }
    } else {
        return @{
            error    = $error[0] | Out-String
            memberof = $null
        }
    }
}


function GraphGetUserPhoto($user) {
    # https://docs.microsoft.com/en-us/graph/api/profilephoto-get?view=graph-rest-1.0
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user)/photo/`$value"
            Headers     = $script:AuthorizationHeader
            ContentType = 'image/jpg'
        }

        $local:tempFile = (Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ((New-Guid).Guid))
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $null = Invoke-RestMethod @requestBody -OutFile $local:tempFile

        $ProgressPreference = $OldProgressPreference

        try { WatchCatchableExitSignal } catch { }

        $local:x = [System.IO.File]::ReadAllBytes($local:tempFile)

        Remove-Item $local:tempFile -Force -ErrorAction SilentlyContinue
    } catch {
        return @{
            error = $error[0] | Out-String
            photo = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error = $false
            photo = $local:x
        }
    } else {
        return @{
            error = $error[0] | Out-String
            photo = $null
        }
    }
}


function GraphPatchUserMailboxsettings($user, $OOFInternal, $OOFExternal, $authHeader = $script:AuthorizationHeader) {
    # https://learn.microsoft.com/en-us/graph/api/user-updatemailboxsettings?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: Mailboxsettings.ReadWrite
    #   Application: Mailboxsettings.ReadWrite

    try { WatchCatchableExitSignal } catch { }

    try {
        if ($OOFInternal -or $OOFExternal) {
            $body = @{}
            $body.add('automaticRepliesSetting', @{})

            if ($OOFInternal) { $Body.'automaticRepliesSetting'.add('internalReplyMessage', $OOFInternal) }

            if ($OOFExternal) { $Body.'automaticRepliesSetting'.add('externalReplyMessage', $OOFExternal) }

            $body = ConvertTo-Json -InputObject $body

            $requestBody = @{
                Method      = 'Patch'
                Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user)/mailboxsettings"
                Headers     = $authHeader
                ContentType = 'Application/Json; charset=utf-8'
                Body        = $body
            }

            $OldProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            $null = Invoke-RestMethod @requestBody

            $ProgressPreference = $OldProgressPreference
        }

        return @{
            error = $false
        }
    } catch {
        return @{
            error = $error[0] | Out-String
        }
    }
}


function GraphFilterGroups($filter) {
    # https://docs.microsoft.com/en-us/graph/api/group-get?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: GroupMember.Read.All
    #   Application: GroupMember.Read.All

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/groups?`$select=securityidentifier&`$filter=" + [System.Net.WebUtility]::UrlEncode($filter)
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error  = $error[0] | Out-String
            groups = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error  = $false
            groups = $local:x
        }
    } else {
        return @{
            error  = $error[0] | Out-String
            groups = $null
        }
    }
}


function GraphFilterUsers($filter) {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users?`$select=securityidentifier&`$filter=" + [System.Net.WebUtility]::UrlEncode($filter)
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error = $error[0] | Out-String
            users = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error = $false
            users = $local:x
        }
    } else {
        return @{
            error = $error[0] | Out-String
            users = $null
        }
    }
}


function GetIniContent ($filePath, $additionalLines) {
    try { WatchCatchableExitSignal } catch { }

    $local:ini = [ordered]@{}
    $local:SectionIndex = -1

    if ($filePath -ne '') {
        try {
            Write-Verbose '    Original ini content'

            foreach ($line in @(@(Get-Content -LiteralPath $FilePath -Encoding UTF8 -ErrorAction Stop) + @($additionalLines -split '\r?\n'))) {
                Write-Verbose "      $line"
                switch -regex ($line) {
                    # Comments starting with ; or # or //, or empty line, whitespace(s) before are ignored
                    '(^\s*(;|#|//))|(^\s*$)' { continue }

                    # Section in square brackets, whitespace(s) before and after brackets are ignored
                    '^\s*\[(.+)\]\s*' {
                        $local:section = ($matches[1]).trim().trim('"').trim('''')
                        if ($null -ne $local:section) {
                            $local:SectionIndex++
                            $local:ini["$($local:SectionIndex)"] = [ordered]@{ '<Set-OutlookSignatures template>' = $local:section }
                        }
                        continue
                    }

                    # Key and value, whitespace(s) before and after brackets are ignored
                    '^\s*(.+?)\s*=\s*(.*)\s*' {
                        if ($null -ne $local:section) {
                            $local:ini["$($local:SectionIndex)"][($matches[1]).trim().trim('"').trim('''')] = ($matches[2]).trim().trim('"').trim('''')
                            continue
                        }
                    }

                    # Key only, whitespace(s) before and after brackets are ignored
                    '^\s*(.*)\s*' {
                        if ($null -ne $local:section) {
                            $local:ini["$($local:SectionIndex)"][($matches[1]).trim().trim('"').trim('''')] = $null
                            continue
                        }
                    }
                }
            }
        } catch {
            Write-Host $error[0]
            Write-Host "Error accessing '$FilePath'. Exit." -ForegroundColor red
            $script:ExitCode = 31
            $script:ExitCodeDescription = "Error accessing '$FilePath'."
            exit
        }
    }

    try { WatchCatchableExitSignal } catch { }

    # default values for <Set-OutlookSignatures configuration>
    if (
        $(
            try {
        ((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))
            } catch {
                $false
            }
        )
    ) {

        if (
            -not $(
                try {
                    $((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))['SortOrder']
                } catch {
                    $false
                }
            )
        ) {
            $((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))['SortOrder'] = 'AsInThisFile'
        }

        if (
            -not $(
                try {
                    $((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))['SortCulture']
                } catch {
                    $false
                }
            )
        ) {
            $((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))['SortCulture'] = 'de-AT'
        }
    } else {
        $local:ini["$($local:ini.Count)"] = [ordered]@{
            '<Set-OutlookSignatures template>' = '<Set-OutlookSignatures configuration>'
            'SortOrder'                        = 'AsInThisFile'
            'SortCulture'                      = 'de-AT'
        }
    }

    try { WatchCatchableExitSignal } catch { }

    return $local:ini
}


function ConvertPath ([ref]$path) {
    try { WatchCatchableExitSignal } catch { }

    if ($path) {
        if (($path.value.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($path.value -ilike '*@SSL\*')) {
            if (-not [System.Uri]::IsWellFormedUriString($path.value, [System.UriKind]::Absolute)) {
                $path.value = ([uri]($path.value -ireplace '@SSL\\', '/' -ireplace '^\\\\', 'https://' -ireplace '\\', '/')).AbsoluteUri
            }
            $path.value = ([uri]$path.value).GetLeftPart([System.UriPartial]::Path) -ireplace "$(([uri]$path.value).GetLeftPart([System.UriPartial]::Authority))/:\S:/\S", $(([uri]$path.value).GetLeftPart([System.UriPartial]::Authority))
            $path.value = ([uri]::UnescapeDataString($path.value) -ireplace ('https://', '\\'))
            $path.value = ([System.URI]$path.value).AbsoluteURI -ireplace 'file:\/\/(.*?)\/(.*)', '\\${1}@SSL\$2' -ireplace '/', '\'
            $path.value = [uri]::UnescapeDataString($path.value)
        } else {
            $path.value = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path.value)

            if ($IsWindows) {
                $path.value = ([System.URI]$path.value).absoluteuri -ireplace 'file:///', '' -ireplace 'file://', '\\' -ireplace '/', '\'
                $path.value = [uri]::UnescapeDataString($path.value)
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function RemoveItemAlternativeRecurse {
    # Function to avoid problems with OneDrive throwing "Access to the cloud file is denied"

    param(
        [alias('LiteralPath')][string] $Path,
        [switch] $SkipFolder # when $Path is a folder, do not delete $path, only it's content
    )

    try { WatchCatchableExitSignal } catch { }

    $local:ToDelete = @()

    if (Test-Path -LiteralPath $path) {
        foreach ($SinglePath in @(Get-Item -LiteralPath $Path)) {
            try { WatchCatchableExitSignal } catch { }

            if (Test-Path -LiteralPath $SinglePath -PathType Container) {
                if (-not $SkipFolder) {
                    $local:ToDelete += @(Get-ChildItem -LiteralPath $SinglePath -Recurse -Force | Sort-Object -Culture $TemplateFilesSortCulture -Property PSIsContainer, @{expression = { $_.FullName.split([IO.Path]::DirectorySeparatorChar).count }; descending = $true }, fullname)
                    $local:ToDelete += @(Get-Item -LiteralPath $SinglePath -Force)
                } else {
                    $local:ToDelete += @(Get-ChildItem -LiteralPath $SinglePath -Recurse -Force | Sort-Object -Culture $TemplateFilesSortCulture -Property PSIsContainer, @{expression = { $_.FullName.split([IO.Path]::DirectorySeparatorChar).count }; descending = $true }, fullname)
                }
            } elseif (Test-Path -LiteralPath $SinglePath -PathType Leaf) {
                $local:ToDelete += (Get-Item -LiteralPath $SinglePath -Force)
            }
        }
    } else {
        # Item to delete does not exist, nothing to do
    }

    foreach ($SingleItemToDelete in $local:ToDelete) {
        try { WatchCatchableExitSignal } catch { }

        try {
            if ((Test-Path $SingleItemToDelete.FullName) -eq $true) {
                Remove-Item $SingleItemToDelete.FullName -Force -Recurse
            }
        } catch {
            Write-Verbose "Could not delete $($SingleItemToDelete.FullName), error: $($_.Exception.Message)"
            Write-Verbose $_
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function ParseJwtToken {
    # Idea for this code: https://www.michev.info/blog/post/2140/decode-jwt-access-and-id-tokens-via-powershell

    [cmdletbinding()]
    param([Parameter(Mandatory = $true)][string]$token)

    try { WatchCatchableExitSignal } catch { }

    # Validate as per https://tools.ietf.org/html/rfc7519
    # Access and ID tokens are fine, Refresh tokens will not work
    if (!$token.Contains('.') -or !$token.StartsWith('eyJ')) {
        return @{
            error   = 'Invalid token'
            header  = $null
            payload = $null
        }
    } else {
        # Header
        $tokenheader = $token.Split('.')[0].Replace('-', '+').Replace('_', '/')

        # Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
        while ($tokenheader.Length % 4) { $tokenheader += '=' }

        # Convert from Base64 encoded string to PSObject all at once
        $tokenHeader = [System.Text.Encoding]::UTF8.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json

        # Payload
        $tokenPayload = $token.Split('.')[1].Replace('-', '+').Replace('_', '/')

        # Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
        while ($tokenPayload.Length % 4) { $tokenPayload += '=' }

        # Convert to Byte array
        $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)

        # Convert to string array
        $tokenArray = [System.Text.Encoding]::UTF8.GetString($tokenByteArray)

        # Convert from JSON to PSObject
        $tokenPayload = $tokenArray | ConvertFrom-Json

        return @{
            error   = $false
            header  = $tokenHeader
            payload = $tokenPayload
        }
    }
}


### ▼▼▼ WatchCatchableExitSignal initiation code below ▼▼▼
##
#
# Place this code in your main script, as early in the code as possible
#
# Call WatchCatchableExitSignal wherever you want to gracefully exit in case of
#   - a Logoff/Reboot/Shutdown message on Windows
#   - a catchable POSIX signal on Linux and macOS
#
# If $WatchCatchableExitSignalNonExitScriptBlock is of type [scriptblock],
#   it is executed when no catchable exit signal is detected.
#
# Clean-up is triggered by WatchCatchableExitSignal running the "$script:ExitCode = 1; $script:ExitCodeDescription = ''; exit" command
#   This triggers the Finally part of a Try/Catch/Finally block
#
# Place the following two lines of code at the end of your clean-up routine
#   WatchCatchableExitSignal -CleanupDone
#

$global:WatchCatchableExitSignalStatus = [hashtable]::Synchronized(@{})
$global:WatchCatchableExitSignalStatus.0 = 'Nothing detected yet'
# Possible values for $global:WatchCatchableExitSignalStatus.0
#   "Nothing detected yet" when no catchable exit signal has been found until now
#   "Detected '<description>', initiate clean-up and exit" when a catchable exit signal has been found
#   "Clean-up done" after clean-up is done

$WatchCatchableExitSignalRunspace = [runspacefactory]::CreateRunspace()
$WatchCatchableExitSignalRunspace.Open()
$WatchCatchableExitSignalRunspace.SessionStateProxy.SetVariable('WatchCatchableExitSignalStatus', $global:WatchCatchableExitSignalStatus)
$WatchCatchableExitSignalPowershell = [powershell]::Create()
$WatchCatchableExitSignalPowershell.Runspace = $WatchCatchableExitSignalRunspace

if ($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) {
    $WatchCatchableExitSignalForm = $null

    $WatchCatchableExitSignalRunspace.SessionStateProxy.SetVariable('WatchCatchableExitSignalForm', [ref]$WatchCatchableExitSignalForm)

    $null = $WatchCatchableExitSignalPowershell.AddScript(
        {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -TypeDefinition @'
            using System;
            using System.Windows.Forms;
            using System.Management.Automation;
            using System.Management.Automation.Runspaces;
            using System.Collections.ObjectModel;

            public class CustomForm : Form {
                public event Action<Message> EndSessionInitiateCleanup;

                protected override CreateParams CreateParams {
                    // Hide the window from Alt-Tab
                    get {
                        CreateParams cp = base.CreateParams;
                        cp.ExStyle |= 0x80;  // WS_EX_TOOLWINDOW
                        return cp;
                    }
                }

                protected override void WndProc(ref Message m) {
                    if (EndSessionInitiateCleanup != null && !this.IsDisposed) {
                        try {
                            EndSessionInitiateCleanup.Invoke(m);
                        } catch {
                            // Do nothing
                        }
                    }

                    base.WndProc(ref m);
                }
            }
'@ -ReferencedAssemblies $(
                if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                    $(@('System.Windows.Forms', 'System.ComponentModel.Primitives', 'System.Management.Automation', 'System.Windows.Forms.Primitives'))
                } else {
                    $(@('System.Windows.Forms', 'System.ComponentModel.Primitives', 'System.Management.Automation'))
                }
            )

            $formRef = $ExecutionContext.SessionState.PSVariable.GetValue('WatchCatchableExitSignalForm')
            $formRef.Value = [CustomForm]::new()
            $formRef.Value.Text = 'Set-OutlookSignatures non-blocking window for WM_* detection'
            $formRef.Value.Width = 300
            $formRef.Value.Height = 300
            $formRef.Value.ShowInTaskbar = $false
            $formRef.Value.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
            $formRef.Value.Opacity = 0
            $formRef.Value.WindowState = [System.Windows.Forms.FormWindowState]::Minimized

            # Event handler
            $formRef.Value.add_EndSessionInitiateCleanup(
                {
                    param($message)

                    try {
                        $WindowsMessagesByDecimal = @{
                            0 = 'WM_NULL'; 2 = 'WM_DESTROY'; 3 = 'WM_MOVE'; 5 = 'WM_SIZE'; 6 = 'WM_ACTIVATE'; 7 = 'WM_SETFOCUS'; 8 = 'WM_KILLFOCUS'; 10 = 'WM_ENABLE'; 11 = 'WM_SETREDRAW'; 12 = 'WM_SETTEXT'; 13 = 'WM_GETTEXT'; 14 = 'WM_GETTEXTLENGTH'; 15 = 'WM_PAINT'; 16 = 'WM_CLOSE'; 17 = 'WM_QUERYENDSESSION'; 18 = 'WM_QUIT'; 19 = 'WM_QUERYOPEN'; 20 = 'WM_ERASEBKGND'; 21 = 'WM_SYSCOLORCHANGE'; 22 = 'WM_ENDSESSION'; 23 = 'WM_SYSTEMERROR'; 24 = 'WM_SHOWWINDOW'; 25 = 'WM_CTLCOLOR'; 26 = 'WM_SETTINGCHANGE'; 27 = 'WM_DEVMODECHANGE'; 28 = 'WM_ACTIVATEAPP'; 29 = 'WM_FONTCHANGE'; 30 = 'WM_TIMECHANGE'; 31 = 'WM_CANCELMODE'; 32 = 'WM_SETCURSOR'; 33 = 'WM_MOUSEACTIVATE'; 34 = 'WM_CHILDACTIVATE'; 35 = 'WM_QUEUESYNC'; 36 = 'WM_GETMINMAXINFO'; 38 = 'WM_PAINTICON'; 39 = 'WM_ICONERASEBKGND'; 40 = 'WM_NEXTDLGCTL'; 42 = 'WM_SPOOLERSTATUS'; 43 = 'WM_DRAWITEM'; 44 = 'WM_MEASUREITEM'; 45 = 'WM_DELETEITEM'; 46 = 'WM_VKEYTOITEM'; 47 = 'WM_CHARTOITEM'; 48 = 'WM_SETFONT'; 49 = 'WM_GETFONT'; 50 = 'WM_SETHOTKEY'; 51 = 'WM_GETHOTKEY'; 55 = 'WM_QUERYDRAGICON'; 57 = 'WM_COMPAREITEM'; 65 = 'WM_COMPACTING'; 70 = 'WM_WINDOWPOSCHANGING'; 71 = 'WM_WINDOWPOSCHANGED'; 72 = 'WM_POWER'; 74 = 'WM_COPYDATA'; 75 = 'WM_CANCELJOURNAL'; 78 = 'WM_NOTIFY'; 80 = 'WM_INPUTLANGCHANGEREQUEST'; 81 = 'WM_INPUTLANGCHANGE'; 82 = 'WM_TCARD'; 83 = 'WM_HELP'; 84 = 'WM_USERCHANGED'; 85 = 'WM_NOTIFYFORMAT'; 123 = 'WM_CONTEXTMENU'; 124 = 'WM_STYLECHANGING'; 125 = 'WM_STYLECHANGED'; 126 = 'WM_DISPLAYCHANGE'; 127 = 'WM_GETICON'; 128 = 'WM_SETICON'; 129 = 'WM_NCCREATE'; 130 = 'WM_NCDESTROY'; 131 = 'WM_NCCALCSIZE'; 132 = 'WM_NCHITTEST'; 133 = 'WM_NCPAINT'; 134 = 'WM_NCACTIVATE'; 135 = 'WM_GETDLGCODE'; 160 = 'WM_NCMOUSEMOVE'; 161 = 'WM_NCLBUTTONDOWN'; 162 = 'WM_NCLBUTTONUP'; 163 = 'WM_NCLBUTTONDBLCLK'; 164 = 'WM_NCRBUTTONDOWN'; 165 = 'WM_NCRBUTTONUP'; 166 = 'WM_NCRBUTTONDBLCLK'; 167 = 'WM_NCMBUTTONDOWN'; 168 = 'WM_NCMBUTTONUP'; 169 = 'WM_NCMBUTTONDBLCLK'; 256 = 'WM_KEYDOWN'; 257 = 'WM_KEYUP'; 258 = 'WM_CHAR'; 259 = 'WM_DEADCHAR'; 260 = 'WM_SYSKEYDOWN'; 261 = 'WM_SYSKEYUP'; 262 = 'WM_SYSCHAR'; 263 = 'WM_SYSDEADCHAR'; 264 = 'WM_KEYLAST'; 269 = 'WM_IME_STARTCOMPOSITION'; 270 = 'WM_IME_ENDCOMPOSITION'; 271 = 'WM_IME_COMPOSITION'; 272 = 'WM_INITDIALOG'; 273 = 'WM_COMMAND'; 274 = 'WM_SYSCOMMAND'; 275 = 'WM_TIMER'; 276 = 'WM_HSCROLL'; 277 = 'WM_VSCROLL'; 278 = 'WM_INITMENU'; 279 = 'WM_INITMENUPOPUP'; 287 = 'WM_MENUSELECT'; 288 = 'WM_MENUCHAR'; 289 = 'WM_ENTERIDLE'; 306 = 'WM_CTLCOLORMSGBOX'; 307 = 'WM_CTLCOLOREDIT'; 308 = 'WM_CTLCOLORLISTBOX'; 309 = 'WM_CTLCOLORBTN'; 310 = 'WM_CTLCOLORDLG'; 311 = 'WM_CTLCOLORSCROLLBAR'; 312 = 'WM_CTLCOLORSTATIC'; 512 = 'WM_MOUSEMOVE'; 513 = 'WM_LBUTTONDOWN'; 514 = 'WM_LBUTTONUP'; 515 = 'WM_LBUTTONDBLCLK'; 516 = 'WM_RBUTTONDOWN'; 517 = 'WM_RBUTTONUP'; 518 = 'WM_RBUTTONDBLCLK'; 519 = 'WM_MBUTTONDOWN'; 520 = 'WM_MBUTTONUP'; 521 = 'WM_MBUTTONDBLCLK'; 522 = 'WM_MOUSEWHEEL'; 526 = 'WM_MOUSEHWHEEL'; 528 = 'WM_PARENTNOTIFY'; 529 = 'WM_ENTERMENULOOP'; 530 = 'WM_EXITMENULOOP'; 531 = 'WM_NEXTMENU'; 532 = 'WM_SIZING'; 533 = 'WM_CAPTURECHANGED'; 534 = 'WM_MOVING'; 536 = 'WM_POWERBROADCAST'; 537 = 'WM_DEVICECHANGE'; 544 = 'WM_MDICREATE'; 545 = 'WM_MDIDESTROY'; 546 = 'WM_MDIACTIVATE'; 547 = 'WM_MDIRESTORE'; 548 = 'WM_MDINEXT'; 549 = 'WM_MDIMAXIMIZE'; 550 = 'WM_MDITILE'; 551 = 'WM_MDICASCADE'; 552 = 'WM_MDIICONARRANGE'; 553 = 'WM_MDIGETACTIVE'; 560 = 'WM_MDISETMENU'; 561 = 'WM_ENTERSIZEMOVE'; 562 = 'WM_EXITSIZEMOVE'; 563 = 'WM_DROPFILES'; 564 = 'WM_MDIREFRESHMENU'; 641 = 'WM_IME_SETCONTEXT'; 642 = 'WM_IME_NOTIFY'; 643 = 'WM_IME_CONTROL'; 644 = 'WM_IME_COMPOSITIONFULL'; 645 = 'WM_IME_SELECT'; 646 = 'WM_IME_CHAR'; 656 = 'WM_IME_KEYDOWN'; 657 = 'WM_IME_KEYUP'; 673 = 'WM_MOUSEHOVER'; 674 = 'WM_NCMOUSELEAVE'; 675 = 'WM_MOUSELEAVE'; 768 = 'WM_CUT'; 769 = 'WM_COPY'; 770 = 'WM_PASTE'; 771 = 'WM_CLEAR'; 772 = 'WM_UNDO'; 773 = 'WM_RENDERFORMAT'; 774 = 'WM_RENDERALLFORMATS'; 775 = 'WM_DESTROYCLIPBOARD'; 776 = 'WM_DRAWCLIPBOARD'; 777 = 'WM_PAINTCLIPBOARD'; 778 = 'WM_VSCROLLCLIPBOARD'; 779 = 'WM_SIZECLIPBOARD'; 780 = 'WM_ASKCBFORMATNAME'; 781 = 'WM_CHANGECBCHAIN'; 782 = 'WM_HSCROLLCLIPBOARD'; 783 = 'WM_QUERYNEWPALETTE'; 784 = 'WM_PALETTEISCHANGING'; 785 = 'WM_PALETTECHANGED'; 786 = 'WM_HOTKEY'; 791 = 'WM_PRINT'; 792 = 'WM_PRINTCLIENT'; 856 = 'WM_HANDHELDFIRST'; 863 = 'WM_HANDHELDLAST'; 896 = 'WM_PENWINFIRST'; 911 = 'WM_PENWINLAST'; 912 = 'WM_COALESCE_FIRST'; 927 = 'WM_COALESCE_LAST'; 992 = 'WM_DDE_INITIATE'; 993 = 'WM_DDE_TERMINATE'; 994 = 'WM_DDE_ADVISE'; 995 = 'WM_DDE_UNADVISE'; 996 = 'WM_DDE_ACK'; 997 = 'WM_DDE_DATA'; 998 = 'WM_DDE_REQUEST'; 999 = 'WM_DDE_POKE'; 1000 = 'WM_DDE_EXECUTE'
                        }

                        if (
                            $(
                                $($WindowsMessagesByDecimal[$($message.Msg)] -ieq 'WM_ENDSESSION') -and
                                $($message.WParam -ne [IntPtr]::Zero)
                            ) -or
                            $($WindowsMessagesByDecimal[$($message.Msg)] -ieq 'WM_QUERYENDSESSION')
                        ) {
                            # Logoff/Reboot/Shutdown will happen.
                            # Set status, wait for clean-up and then return 0.
                            $global:WatchCatchableExitSignalStatus.0 = "Detected '$(@(@($($message.Msg), $($WindowsMessagesByDecimal[$($message.Msg)]), $($message.WParam), $($message.LParam)) | Where-Object {$_})-join ', ')', initiate clean-up and exit"

                            until (
                                $($global:WatchCatchableExitSignalStatus.0 -ieq 'Clean-up done')
                            ) {
                                Start-Sleep -Milliseconds 100
                            }

                            $message.Result = [IntPtr]::Zero

                            $formRef.Value.Close()
                        }
                    } catch {
                    }
                }
            )

            $formRef.Value.ShowDialog()
        }
    )
} elseif ($IsLinux -or $IsMacOS) {
    $null = $WatchCatchableExitSignalPowershell.AddScript(
        {
            # Use trap instead of try/catch, because trap reacts to catchable POSIX signals
            trap {
                $global:WatchCatchableExitSignalStatus.0 = "Detected '$($_)', initiate clean-up and exit"

                until ($global:WatchCatchableExitSignalStatus.0 -ieq 'Clean-up done') {
                    Start-Sleep -Milliseconds 100
                }
            }

            while ($true) {
                Start-Sleep -Milliseconds 100
            }
        }
    )
}

$null = $WatchCatchableExitSignalPowershell.BeginInvoke()


function global:WatchCatchableExitSignal {
    param (
        [ScriptBlock]$NonExitScriptBlock,
        [switch]$CleanupDone
    )

    if ($CleanupDone) {
        if ($WatchCatchableExitSignalForm) {
            try {
                $WatchCatchableExitSignalForm.Close()
            } catch {
                # Do nothing
            }
        }

        $global:WatchCatchableExitSignalStatus.0 = 'Clean-up done'
    } elseif ($global:WatchCatchableExitSignalStatus.0 -ilike "Detected '*', initiate clean-up and exit") {
        Write-Host
        Write-Host "WatchCatchableExitSignal: $($global:WatchCatchableExitSignalStatus.0)" -ForegroundColor Yellow

        if ($WatchCatchableExitSignalForm) {
            try {
                $WatchCatchableExitSignalForm.Close()
            } catch {
                # Do nothing
            }
        }

        $script:ExitCode = 1
        $script:ExitCodeDescription = 'Detected catchable exit signal.'
        exit
    } else {
        if ($WatchCatchableExitSignalNonExitScriptBlock -and ($WatchCatchableExitSignalNonExitScriptBlock -is [ScriptBlock])) {
            try {
                . $WatchCatchableExitSignalNonExitScriptBlock
            } catch {
                # Do nothing
            }
        }
    }
}
#
##
### ▲▲▲ WatchCatchableExitSignal initiation code above ▲▲▲


$WatchCatchableExitSignalNonExitScriptBlock = {
    try {
        $script:COMWord.Visible = $false
    } catch {
    }

    try {
        $script:COMWordDummy.Visible = $false
    } catch {
    }
}


#
# All functions have been defined above
# Initially executed code starts here
#


$script:ExitCode = 255
$script:ExitCodeDescription = 'Generic exit code, no details available.'


try {
    try {
        $TranscriptFullName = Join-Path -Path $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Set-OutlookSignatures\Logs') -ChildPath $("Set-OutlookSignatures_Log_$(Get-Date $([DateTime]::UtcNow) -Format FileDateTimeUniversal).txt")
        $TranscriptFullName = (Start-Transcript -LiteralPath $TranscriptFullName -Force).Path

        "This folder contains log files generated by Set-OutlookSignatures.$([Environment]::NewLine)Each file is named according to the pattern ""Set-OutlookSignatures_Log_yyyyMMddTHHmmssffffZ.txt"".$([Environment]::NewLine)Files older than 14 days are automatically deleted with each execution of Set-OutlookSignatures." | Out-File -LiteralPath $(Join-Path -Path (Split-Path -Path $TranscriptFullName) -ChildPath '_README.txt') -Encoding utf8 -Force
    } catch {
        $TranscriptFullName = $null
    }


    Write-Host
    Write-Host "Start Set-OutlookSignatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($TranscriptFullName) {
        Write-Host "  Log file: '$TranscriptFullName'"

        try {
            Get-ChildItem -LiteralPath $(Split-Path -LiteralPath $TranscriptFullName) -File -Force | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-14) } | ForEach-Object {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        } catch {
        }
    }

    if ($PSVersion -ge [version]'7.5') {
        Write-Host '  PowerShell 7.5 and higher versions are not yet supported because .Net 9 causes incompatibilities.'
        Write-Host '  Please use PowerShell 7.4 or lower versions. Exit.'
        $script:ExitCode = 254
        $script:ExitCodeDescription = 'PowerShell 7.5 or higher detected.'
        exit
    }

    if ($psISE) {
        Write-Host '  PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
        Write-Host '  Required features are not available in ISE. Exit.' -ForegroundColor Red
        $script:ExitCode = 2
        $script:ExitCodeDescription = 'PowerShell ISE detected.'
        exit
    }

    if (($ExecutionContext.SessionState.LanguageMode) -ine 'FullLanguage') {
        {
            Write-Host '' This PowerShell session runs in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode."" -ForegroundColor Red
            Write-Host '  Required features are only available in FullLanguage mode. Exit.' -ForegroundColor Red
            $script:ExitCode = 32
            $script:ExitCodeDescription = 'Not running in FullLanguage mode.'
            exit
        }
    }

    if ($global:SetOutlookSignaturesLastRunGuid) {
        Write-Host '  Set-OutlookSignatures has already been run before in this PowerShell session.' -ForegroundColor Yellow
        Write-Host '    It is strongly recommended to run Set-OutlookSignatures only once per session, ideally in a fresh one.' -ForegroundColor Yellow
        Write-Host '    This is the only way to avoid problem caused by .Net caching DLL files in memory.' -ForegroundColor Yellow
        Write-Host '    Use at your own risk!' -ForegroundColor Yellow

        # $script:ExitCode = 3
        # $script:ExitCodeDescription = 'Script already run in this PowerShell session, is only supported once.'
        # exit
    } else {
        $global:SetOutlookSignaturesLastRunGuid = (New-Guid).Guid
    }

    if (-not (Test-Path 'variable:IsWindows')) {
        $script:IsWindows = $true
        $script:IsLinux = $false
        $script:IsMacOS = $false
    }

    BlockSleep

    try { WatchCatchableExitSignal } catch { }

    $OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

    Set-Location $PSScriptRoot | Out-Null

    $ScriptInvocation = $MyInvocation

    $script:tempDir = (New-Item -Path ([System.IO.Path]::GetTempPath()) -Name (New-Guid).Guid -ItemType Directory).FullName
    $script:ScriptRunGuid = Split-Path -Path $script:tempDir -Leaf

    $script:SetOutlookSignaturesCommonDllFilePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))
    Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\Set-OutlookSignatures\Set-OutlookSignatures.Common.dll')) -Destination $script:SetOutlookSignaturesCommonDllFilePath
    if (-not $IsLinux) {
        Unblock-File -LiteralPath $script:SetOutlookSignaturesCommonDllFilePath
    }

    try {
        Import-Module -Name $script:SetOutlookSignaturesCommonDllFilePath -Force -ErrorAction Stop
    } catch {
        Write-Host $error[0]
        Write-Host '    Problem importing Set-OutlookSignatures.Common.dll. Exit.' -ForegroundColor Red
        $script:ExitCode = 4
        $script:ExitCodeDescription = 'Problem importing Set-OutlookSignatures.Common.dll.'
        exit
    }

    try { WatchCatchableExitSignal } catch { }

    main

    $script:ExitCode = 0
    $script:ExitCodeDescription = 'Success.'
} catch {
    Write-Host $error[0]
    Write-Host
    Write-Host 'Unexpected error. Exit.' -ForegroundColor red
} finally {
    Write-Host
    Write-Host "Clean-up @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    # Restore original Word AlertIfNotDefault setting
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value $script:WordAlertIfNotDefaultOriginal -ErrorAction SilentlyContinue | Out-Null

    # Restore original Word security setting
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction SilentlyContinue | Out-Null

    if ($script:COMWordDummy) {
        if ($script:COMWordDummy.ActiveDocument) {
            try {
                $script:COMWordDummy.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal
            } catch {
            }

            # Restore original WebOptions
            try {
                if ($script:WordWebOptions) {
                    foreach ($property in @('TargetBrowser', 'BrowserLevel', 'AllowPNG', 'OptimizeForBrowser', 'RelyOnCSS', 'RelyOnVML', 'Encoding', 'OrganizeInFolder', 'PixelsPerInch', 'ScreenSize', 'UseLongFileNames')) {
                        $script:COMWordDummy.ActiveDocument.WebOptions.$property = $script:WordWebOptions.$property
                    }
                }
            } catch {}

            # Restore original TextEncoding
            try {
                if ($script:WordTextEncoding) {
                    $script:COMWordDummy.ActiveDocument.TextEndocing = $script:WordTextEncoding
                }
            } catch {
            }
        }

        try {
            $script:COMWordDummy.Quit([ref]$false)
        } catch {}

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWordDummy) | Out-Null

        Remove-Variable -Name 'COMWordDummy' -Scope 'script'
    }

    if ($script:COMWord) {
        if ($script:COMWord.ActiveDocument) {
            try {
                $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal
            } catch {
            }

            # Restore original WebOptions
            try {
                if ($script:WordWebOptions) {
                    foreach ($property in @('TargetBrowser', 'BrowserLevel', 'AllowPNG', 'OptimizeForBrowser', 'RelyOnCSS', 'RelyOnVML', 'Encoding', 'OrganizeInFolder', 'PixelsPerInch', 'ScreenSize', 'UseLongFileNames')) {
                        $script:COMWord.ActiveDocument.WebOptions.$property = $script:WordWebOptions.$property
                    }
                }
            } catch {}

            # Restore original TextEncoding
            try {
                if ($script:WordTextEncoding) {
                    $script:COMWord.ActiveDocument.TextEndocing = $script:WordTextEncoding
                }
            } catch {
            }
        }

        try {
            $script:COMWord.Quit([ref]$false)
        } catch {}

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWord) | Out-Null

        Remove-Variable -Name 'COMWord' -Scope 'script'
    }

    if ($script:SetOutlookSignaturesCommonDllFilePath) {
        Remove-Module -Name $([System.IO.Path]::GetFileNameWithoutExtension($script:SetOutlookSignaturesCommonDllFilePath)) -Force -ErrorAction SilentlyContinue
        Remove-Item $script:SetOutlookSignaturesCommonDllFilePath -Force -ErrorAction SilentlyContinue
    }

    if ($script:BenefactorCircleLicenseFilePath) {
        Remove-Module -Name $([System.IO.Path]::GetFileNameWithoutExtension($script:BenefactorCircleLicenseFilePath)) -Force -ErrorAction SilentlyContinue
        Remove-Item $script:BenefactorCircleLicenseFilePath -Force -ErrorAction SilentlyContinue
    }

    if ($script:WebServicesDllPath) {
        Remove-Module -Name $([System.IO.Path]::GetFileNameWithoutExtension($script:WebServicesDllPath)) -Force -ErrorAction SilentlyContinue
        Remove-Item $script:WebServicesDllPath -Force -ErrorAction SilentlyContinue
    }

    if ($script:MsalModulePath) {
        Remove-Module -Name MSAL.PS -Force -ErrorAction SilentlyContinue
        Remove-Item $script:MsalModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:PreMailerNetModulePath) {
        Remove-Item $script:PreMailerNetModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:AngleSharpCssNetModulePath) {
        Remove-Module -Name AngleSharp.Css -Force -ErrorAction SilentlyContinue
        Remove-Module -Name AngleSharp -Force -ErrorAction SilentlyContinue
        Remove-Item $script:AngleSharpCssNetModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:QRCoderModulePath) {
        Remove-Module -Name AngleSharp -Force -ErrorAction SilentlyContinue
        Remove-Item $script:QRCoderModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:ScriptProcessPriorityOriginal -and $IsWindows) {
        $null = Get-CimInstance Win32_process -Filter "ProcessId = ""$PID""" | Invoke-CimMethod -Name SetPriority -Arguments @{Priority = $script:ScriptProcessPriorityOriginal }
    }

    if ($script:SystemNetServicePointManagerSecurityProtocolOld) {
        [System.Net.ServicePointManager]::SecurityProtocol = $script:SystemNetServicePointManagerSecurityProtocolOld
    }

    if ($script:tempDir) {
        Remove-Item $script:tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($TranscriptFullName) {
        Write-Host
        Write-Host 'Log file'
        Write-Host "  '$TranscriptFullName'"
    }

    Write-Host
    Write-Host 'Exit code' -ForegroundColor $(if ($script:ExitCode -eq 0) { (Get-Host).ui.rawui.ForegroundColor } else { 'Yellow' })
    Write-Host "  Code: $($script:ExitCode)" -ForegroundColor $(if ($script:ExitCode -eq 0) { (Get-Host).ui.rawui.ForegroundColor } else { 'Yellow' })
    Write-Host "  Description: '$($script:ExitCodeDescription)'" -ForegroundColor $(if ($script:ExitCode -eq 0) { (Get-Host).ui.rawui.ForegroundColor } else { 'Yellow' })

    if ($script:ExitCode -ne 0) {
        Write-Host '  Check for existing issues at https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues?q=' -ForegroundColor Yellow
        Write-Host '  or request commercial support from ExplicIT Consulting at https://explicitconsulting.at/open-source/set-outlooksignatures.' -ForegroundColor Yellow
    }

    Write-Host
    Write-Host "End Set-OutlookSignatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($TranscriptFullName) {
        Stop-Transcript | Out-Null
    }

    # Allow sleep
    BlockSleep -AllowSleep

    # Stop watching for catchable exit signals
    try { WatchCatchableExitSignal -CleanupDone } catch { }

    # End script with exit 0 or whatever is defined in $script:ExitCode
    exit $script:ExitCode
}
