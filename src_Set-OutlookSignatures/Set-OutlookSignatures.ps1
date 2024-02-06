<#
.SYNOPSIS
Set-OutlookSignatures XXXVersionStringXXX
The open source gold standard to centrally manage and deploy email signatures and out-of-office replies for Outlook and Exchange.

.DESCRIPTION
Email signatures and out-of-office replies are an integral part of corporate identity and corporate design, of successful concepts for media and internet presence, and of marketing campaigns.

Central administration and distribution ensures that CI/CD guidelines are met, guarantees the use of correct and up-to-date data, helps to comply with legal requirements, relieves staff and also opens up an additional marketing channel.

**With Set-OutlookSignatures, you can do all this and even more. Signatures and out-of-office replies can be:**
- Generated from **templates in DOCX or HTML** file format
- Customized with a **broad range of variables**, including **photos**, from Active Directory and other sources
  - Variables are available for the **currently logged-on user, this user's manager, each mailbox and each mailbox's manager**
  - Images in signatures can be **bound to the existence of certain variables** (useful for optional social network icons, for example)
- Applied to all **mailboxes (including shared mailboxes)**, specific **mailbox groups**, specific **email addresses** or specific **user or mailbox properties**, for **every mailbox across all Outlook profiles (Outlook, New Outlook, Outlook Web)** (**automapped and additional mailboxes** are optional)
- Created with different names from the same template (e.g., **one template can be used for multiple shared mailboxes**)
- Assigned **time ranges** within which they are valid
- Set as **default signature** for new emails, or for replies and forwards (signatures only)
- Set as **default OOF message** for internal or external recipients (OOF messages only)
- Set in **Outlook Web** for the currently logged-in user, including mirroring signatures the the cloud as **roaming signatures**
- Centrally managed only or **exist along user-created signatures** (signatures only)
- Copied to an **alternate path** for easy access on mobile devices not directly supported by this script (signatures only)
- **Write protected** (Outlook signatures only)

Set-OutlookSignatures can be **executed by users on clients and terminal servers, or on a central server without end user interaction**.
On clients, it can run as part of the logon script, as scheduled task, or on user demand via a desktop icon, start menu entry, link or any other way of starting a program.
Signatures and OOF messages can also be created and deployed centrally, without end user or client involvement.

**Sample templates** for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

**Simulation mode** allows content creators and admins to simulate the behavior of the software and to inspect the resulting signature files before going live.

The software is **designed to work in big and complex environments** (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). It works **on premises, in hybrid and cloud-only environments**.
All **national clouds are supported**: Public (AzurePublic), US Government L4 (AzureUSGovernment), US Government L5 (AzureUSGovernment DoD), China (AzureChinaCloud operated by 21Vianet).

It is **multi-client capable** by using different template paths, configuration files and script parameters.

Set-OutlookSignatures requires **no installation on servers or clients**. You only need a standard SMB file share on a central system, and Office on your clients.

A **documented implementation approach**, based on real life experiences implementing the software in multi-client environments with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and email and client administrators.
The implementation approach is **suited for service providers as well as for clients**, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The software core is **Free and Open-Source Software (FOSS)**. It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see `.\LICENSE.txt` for copyright and MIT license details.

**Some features are exclusive to Benefactor Circle members.** Benefactor Circle members have access to an extension file enabling the exclusive features. This extension file is chargeable, and it is distributed under a proprietary, non-free and non-open-source license.  Please see `.\docs\Benefactor Circle` for details.
.LINK
Github: https://github.com/Set-OutlookSignatures/Set-OutlookSignatures

.PARAMETER SignatureTemplatePath
Path to centrally managed signature templates.

Local and remote paths are supported.

Local paths can be absolute ('C:\Signature templates') or relative to the software path ('.\sample templates\Signatures DOCX').

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

Default value: '.\sample templates\Signatures DOCX'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '.\sample templates\Signatures DOCX'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '.\sample templates\Signatures DOCX'"

.PARAMETER SignatureIniPath
Path to ini file containing signature template tags.

The file must be UTF8 encoded.

See '.\sample templates\Signatures DOCX\_Signatures.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the software path ('.\sample templates\Signatures DOCX')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

The currently logged-in user needs at least read access to the path

Default value: '.\sample templates\Signatures DOCX\_Signatures.ini'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -SignatureIniPath '.\templates\Signatures DOCX\_Signatures.ini'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -SignatureIniPath '.\templates\Signatures DOCX\_Signatures.ini'"

.PARAMETER ReplacementVariableConfigFile
Path to a replacement variable config file.

The file must be UTF8 encoded.

Local and remote paths are supported.

Local paths can be absolute ('C:\Signature templates') or relative to the software path ('.\sample templates\Signatures DOCX').

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

Default value: '.\config\default replacement variables.ps1'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -ReplacementVariableConfigFile '.\config\default replacement variables.ps1'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -ReplacementVariableConfigFile '.\config\default replacement variables.ps1'"

.PARAMETER GraphConfigFile
Path to a Graph variable config file.

The file must be UTF8 encoded.

Local and remote paths are supported.

Local paths can be absolute ('C:\config\default graph config.ps1') or relative to the software path ('.\config\default graph config.ps1')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/config/default graph config.ps1' or '\\server.domain@SSL\SignatureSite\config\default graph config.ps1'

The currently logged-in user needs at least read access to the path

Default value: '.\config\default graph config.ps1'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -GraphConfigFile '.\config\default graph config.ps1'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 GraphConfigFile '.\config\default graph config.ps1'"

.PARAMETER TrustsToCheckForGroups
List of domains to check for group membership.

If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered.

If a string starts with a minus or dash ('-domain-a.local'), the domain after the dash or minus is removed from the list (no wildcards allowed).

All domains belonging to the Active Directory forest of the currently logged-in user are always considered, but specific domains can be removed ('*', '-childA1.childA.user.forest').

When a cross-forest trust is detected by the '*' option, all domains belonging to the trusted forest are considered but specific domains can be removed ('*', '-childX.trusted.forest').

Default value: '*'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -TrustsToCheckForGroups 'corp.example.com', 'corp.example.net'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -TrustsToCheckForGroups 'corp.example.com', 'corp.example.net'"

.PARAMETER IncludeMailboxForestDomainLocalGroups
Shall the software consider group membership in domain local groups in the mailbox's AD forest?

Per default, membership in domain local groups in the mailbox's forest is not considered as the required LDAP queries are slow and domain local groups are usually not used in Exchange.

Domain local groups across trusts behave differently, they are always considered as soon as the trusted domain/forest is included in TrustsToCheckForGroups.

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
Shall the software set the out of office (OOF) message of the currently logged-in user?

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

Local paths can be absolute ('C:\OOF templates') or relative to the software path ('.\sample templates\ Out of Office ').

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/OOFTemplates' or '\\server.domain@SSL\SignatureSite\OOFTemplates'

The currently logged-in user needs at least read access to the path.

Default value: '.\sample templates\Out of Office DOCX'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -OOFTemplatePath '.\templates\Out of Office DOCX'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -OOFTemplatePath '.\templates\Out of Office DOCX'"

.PARAMETER OOFIniPath
Path to ini file containing signature template tags.

The file must be UTF8 encoded.

See '.\sample templates\Out of Office DOCX\_OOF.ini' for a sample file with further explanations.

Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the software path ('.\sample templates\Signatures')

WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'

The currently logged-in user needs at least read access to the path

Default value: '.\sample templates\Out of Office DOCX\_OOF.ini'

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -OOFIniPath '.\templates\Out of Office DOCX\_OOF.ini'
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -OOFIniPath '.\templates\Out of Office DOCX\_OOF.ini'"

.PARAMETER AdditionalSignaturePath
An additional path that the signatures shall be copied to.

Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.

This way, the user can easily copy-paste the preferred preconfigured signature for use in an email app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.

Local and remote paths are supported.

Local paths can be absolute ('C:\Outlook signatures') or relative to the software path ('.\Outlook signatures').

WebDAV paths are supported (https only): 'https://server.domain/User' or '\\server.domain@SSL\User'

The currently logged-in user needs at least write access to the path.

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

Default value: $false

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
Try to connect to Microsoft Graph only, ignoring any local Active Directory.

The default behavior is to try Active Directory first and fall back to Graph.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false

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

Outlook 2016 and newer can handle images embedded directly into an HTML file as BASE64 string ('<img src="data:image/[...]"').

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
Disable signature roaming in Outlook. Has no effect on signature roaming via the MirrorLocalSignaturesToCloud parameter.

A value representing true disables roaming signatures, a value representing false enables roaming signatures, any other value leaves the setting as-is.

Only sets HKCU registry key, does not override configuration set by group policy.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no', $null, ''

Default value: $true

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DisableRoamingSignatures $true
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -DisableRoamingSignatures true
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DisableRoamingSignatures $true"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -DisableRoamingSignatures true"

.PARAMETER MirrorLocalSignaturesToCloud
Should local signatures be uploaded as roaming signature for the current user?

Possible for Exchange Online mailbox of currently logged-in user.

This feature requires a Benefactor Circle license.

Allowed values: 1, 'true', '$true', 'yes', 0, 'false', '$false', 'no'

Default value: $false

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -MirrorLocalSignaturesToCloud $false
Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -MirrorLocalSignaturesToCloud false
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -MirrorLocalSignaturesToCloud $false"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -MirrorLocalSignaturesToCloud false"

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

.PARAMETER BenefactorCircleId
The Benefactor Circle member Id matching your license file, which unlocks exclusive features.

Default value: ''

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -BenefactorCircleId 00000000-0000-0000-0000-000000000000
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -BenefactorCircleId 00000000-0000-0000-0000-000000000000"

.PARAMETER BenefactorCircleLicenseFile
The Benefactor Circle license file matching your member Id, which unlocks exclusive features.

Default value: ''

Usage example PowerShell: & .\Set-OutlookSignatures.ps1 -BenefactorCircleLicenseFile ".\license.dll"
Usage example Non-PowerShell: powershell.exe -command "& .\Set-OutlookSignatures.ps1 -BenefactorCircleLicenseFile "".\license.dll"""

.INPUTS
None. You cannot pipe objects to Set-OutlookSignatures.ps1.

.OUTPUTS
Set-OutlookSignatures.ps1 writes the current activities, warnings and error messages to the standard output stream.

.EXAMPLE
Run Set-OutlookSignatures with default values and sample templates
PS> .\Set-OutlookSignatures.ps1

.EXAMPLE
Use custom signature templates and custom ini file
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureIniPath '\\internal.example.com\share\Signature Templates\_Signatures.ini'

.EXAMPLE
Use custom signature templates, ignore trust to internal-test.example.com
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -TrustsToCheckForGroups '*', '-internal-test.example.com'

.EXAMPLE
Use custom signature templates, only check domains/trusts internal-test.example.com and company.b.com
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -TrustsToCheckForGroups 'internal-test.example.com', 'company.b.com'

.EXAMPLE
Passing arguments to PowerShell.exe from the command line or task scheduler can be very tricky when spaces are involved. See '.\docs\README' for details.
PowerShell.exe -Command "& '\\server\share\directory\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\server\share\directory\templates\Signatures DOCX' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -OOFTemplatePath '\\server\share\directory\templates\Out of Office DOCX' -OOFTemplatePath '\\internal.example.com\share\Signature Templates\_OOF.ini' -ReplacementVariableConfigFile '\\server\share\directory\config\default replacement variables.ps1' "

.EXAMPLE
Please see '.\docs\README' and https://github.com/Set-OutlookSignatures/Set-OutlookSignatures for more details.

.NOTES
Software: Set-OutlookSignatures
Version : XXXVersionStringXXX
Web     : https://github.com/Set-OutlookSignatures/Set-OutlookSignatures
License : MIT license (see '.\LICENSE.txt' for details and copyright)
#>


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
    [string]$BenefactorCircleId = '',

    # Path to centrally managed signature templates
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$SignatureTemplatePath = '.\sample templates\Signatures DOCX',

    # Path to ini file containing signature template tags
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$SignatureIniPath = '.\sample templates\Signatures DOCX\_Signatures.ini',

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

    # Use templates in .HTM file format instead of .DOCX
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $UseHtmTemplates = $false,

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

    # Shall the software set the out of office (OOF) message(s) of the currently logged-in user?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SetCurrentUserOOFMessage = $true,

    # Path to centrally managed out of office (OOF, automatic reply) templates
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$OOFTemplatePath = '.\sample templates\Out of Office DOCX',

    # Path to ini file containing OOF template tags
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$OOFIniPath = '.\sample templates\Out of Office DOCX\_OOF.ini',

    # Path to a replacement variable config file.
    [Parameter(Mandatory = $false, ParameterSetName = 'D: Replacement variables')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$ReplacementVariableConfigFile = '.\config\default replacement variables.ps1',

    # Try to connect to Microsoft Graph only, ignoring any local Active Directory.
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $GraphOnly = $false,

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
    $MirrorLocalSignaturesToCloud = $false,

    # Word process priority
    [Parameter(Mandatory = $false, ParameterSetName = 'H: Word')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet('Idle', 64, 'BelowNormal', 16384, 'Normal', 32, 'AboveNormal', 32768, 'High', 128, 'RealTime', 256)]
    $WordProcessPriority = 'Normal',

    # Script process priority
    [Parameter(Mandatory = $false, ParameterSetName = 'I: Script')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet('Idle', 64, 'BelowNormal', 16384, 'Normal', 32, 'AboveNormal', 32768, 'High', 128, 'RealTime', 256)]
    $ScriptProcessPriority = 'Normal'
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
        Add-Member -InputObject $current -MemberType NoteProperty -Name Rank -Value $rank
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

    # Windows reserved file names and device names (CON, PRN, AUX, COMx, LPTx, ...)
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


function main {
    Set-Location $PSScriptRoot | Out-Null

    $ScriptVersion = 'XXXVersionStringXXX'

    $OldProgressPreference = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    try {
        $GitHubReleases = @()

        (Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Set-OutlookSignatures/Set-OutlookSignatures/main/docs/releases.txt' -UseBasicParsing).Content -split '\r?\n' | Where-Object { $_ } | ForEach-Object {
            $GitHubReleases += $_
        }
    } catch {
        $GitHubReleases = @(, 'v0.0.0')
    }

    $ProgressPreference = $OldProgressPreference

    $GitHubReleases += $ScriptVersion

    $GitHubReleasesSemVer = @()

    foreach ($v in $GitHubReleases) {
        $GitHubReleasesSemVer += toSemVer $v
    }

    $GitHubReleasesSemVerRanked = rankedSemVer($GitHubReleasesSemVer) | Sort-Object -Culture $TemplateFilesSortCulture -Property Rank -Unique

    $GitHubReleasesNewer = @(@($GitHubReleasesSemVerRanked | Where-Object { $_.Rank -gt ($GitHubReleasesSemVerRanked | Where-Object { $_.VersionString -ieq $ScriptVersion }).Rank }).VersionString)

    Write-Host
    Write-Host "Script notes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Write-Host '  Software: Set-OutlookSignatures'
    Write-Host "  Version : $ScriptVersion"

    if ($GitHubReleasesNewer) {
        Write-Host "            At least one release newer than $($ScriptVersion) is available: $(@($GitHubReleasesNewer[0, -1] | Select-Object -Unique) -join $(if($GitHubReleasesNewer.Count -gt 2) { ', ..., ' } else {', '}))" -ForegroundColor Yellow
    }

    Write-Host '  Web     : https://github.com/Set-OutlookSignatures/Set-OutlookSignatures'
    Write-Host "  License : MIT license (see '.\LICENSE.txt' for details and copyright)"


    Write-Host
    Write-Host "Check parameters and script environment @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    Write-Host "  PowerShell: '$((($($PSVersionTable.PSVersion), $($PSVersionTable.PSEdition), $($PSVersionTable.Platform), $($PSVersionTable.OS)) | Where-Object {$_}) -join "', '")'"

    Write-Host "  PowerShell bitness: $(if ([Environment]::Is64BitProcess -eq $false) {'Non-'})64-bit process on a $(if ([Environment]::Is64OperatingSystem -eq $false) {'Non-'})64-bit operating system"

    Write-Host "  PowerShell parameters: '$ScriptPassedParameters'"

    Write-Host "  Script path: '$PSCommandPath'"

    if ((Test-Path 'variable:IsWindows')) {
        # Automatic variable $IsWindows is available, must be cross-platform PowerShell version v6+
        if ($IsWindows -eq $false) {
            Write-Host "  Your OS: $($PSVersionTable.Platform), $($PSVersionTable.OS), $(Invoke-Expression '(lsb_release -ds || cat /etc/*release || uname -om) 2>/dev/null | head -n1')" -ForegroundColor Red
            Write-Host '  This script is supported on Windows only. Exit.' -ForegroundColor Red
            exit 1
        }
    } else {
        # Automatic variable $IsWindows is not available, must be PowerShell <v6 running on Windows
    }

    if (($ExecutionContext.SessionState.LanguageMode) -ine 'FullLanguage') {
        Write-Host "  This PowerShell session runs in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode." -ForegroundColor Red
        Write-Host '  Required features are only available in FullLanguage mode. Exit.' -ForegroundColor Red
        exit 1
    }

    $script:tempDir = [System.IO.Path]::GetTempPath()
    $script:jobs = New-Object System.Collections.ArrayList
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    Add-Type -AssemblyName System.Web
    $Search = New-Object DirectoryServices.DirectorySearcher
    $Search.PageSize = 1000

    $HTMLMarkerTag = '<meta name=data-SignatureFileInfo content="Set-OutlookSignatures">'

    # Connected Files - description and folder name sources:
    #   https://docs.microsoft.com/en-us/windows/win32/shell/manage#connected-files
    #   https://docs.microsoft.com/en-us/office/vba/api/word.defaultWebOptions.foldersuffix
    $ConnectedFilesFolderNames = ('.files', '_archivos', '_arquivos', '_bestanden', '_bylos', '_datoteke', '_dosyalar', '_elemei', '_failid', '_fails', '_fajlovi', '_ficheiros', '_fichiers', '_file', '_files', '_fitxategiak', '_fitxers', '_pliki', '_soubory', '_tiedostot', '-Dateien', '-filer')


    Write-Host ('  TrustsToCheckForGroups: ' + ('''' + $($TrustsToCheckForGroups -join ''', ''') + ''''))

    Write-Host "  IncludeMailboxForestDomainLocalGroups: '$IncludeMailboxForestDomainLocalGroups'"
    if ($IncludeMailboxForestDomainLocalGroups -iin (1, '1', 'true', '$true', 'yes')) {
        $IncludeMailboxForestDomainLocalGroups = $true
    } else {
        $IncludeMailboxForestDomainLocalGroups = $false
    }

    Write-Host "  SignatureTemplatePath: '$SignatureTemplatePath'" -NoNewline
    ConvertPath ([ref]$SignatureTemplatePath)
    CheckPath $SignatureTemplatePath

    Write-Host "  SignatureIniPath: '$SignatureIniPath'" -NoNewline
    if ($SignatureIniPath) {
        ConvertPath ([ref]$SignatureIniPath)
        CheckPath $SignatureIniPath
        $SignatureIniSettings = GetIniContent $SignatureIniPath

        Write-Verbose '    Parsed ini content'
        foreach ($section in $SignatureIniSettings.GetEnumerator()) {
            Write-Verbose "      Signature ini index #: '$($section.name)'"
            $local:tags = @()
            foreach ($key in $SignatureIniSettings[$($section.name)].GetEnumerator()) {
                if ($key.value) {
                    $local:tags += "$($key.name) = $($key.value)"
                } else {
                    $local:tags += "$($key.name)"
                }
            }
            Write-Verbose "        Tags: [$($local:tags -join '] [')]"
        }
    } else {
        $SignatureIniSettings = @{}
        Write-Host
    }

    Write-Host "  SetCurrentUserOutlookWebSignature: '$SetCurrentUserOutlookWebSignature'"
    if ($SetCurrentUserOutlookWebSignature -iin (1, '1', 'true', '$true', 'yes')) {
        $SetCurrentUserOutlookWebSignature = $true
    } else {
        $SetCurrentUserOutlookWebSignature = $false
    }

    Write-Host "  SetCurrentUserOOFMessage: '$SetCurrentUserOOFMessage'"
    if ($SetCurrentUserOOFMessage -iin (1, '1', 'true', '$true', 'yes')) {
        $SetCurrentUserOOFMessage = $true
    } else {
        $SetCurrentUserOOFMessage = $false
    }

    if ($SetCurrentUserOOFMessage) {
        Write-Host "  OOFTemplatePath: '$OOFTemplatePath'" -NoNewline
        ConvertPath ([ref]$OOFTemplatePath)
        CheckPath $OOFTemplatePath
        Write-Host "  OOFIniPath: '$OOFIniPath'" -NoNewline
        if ($OOFIniPath) {
            ConvertPath ([ref]$OOFIniPath)
            CheckPath $OOFIniPath
            $OOFIniSettings = GetIniContent $OOFIniPath

            Write-Verbose '    Parsed ini content'
            foreach ($section in $OOFIniSettings.GetEnumerator()) {
                Write-Verbose "      OOF ini index #: '$($section.name)'"
                $local:tags = @()
                foreach ($key in $OOFIniSettings[$($section.name)].GetEnumerator()) {
                    if ($key.value) {
                        $local:tags += "$($key.name) = $($key.value)"
                    } else {
                        $local:tags += "$($key.name)"
                    }
                }
                Write-Verbose "        Tags: [$($local:tags -join '] [')]"
            }
        } else {
            $OOFIniSettings = @{}
            Write-Host
        }
    }

    Write-Host "  UseHtmTemplates: '$UseHtmTemplates'"
    if ($UseHtmTemplates -iin (1, '1', 'true', '$true', 'yes')) {
        $UseHtmTemplates = $true
    } else {
        $UseHtmTemplates = $false
    }

    Write-Host "  GraphOnly: '$GraphOnly'"
    if ($GraphOnly -iin (1, '1', 'true', '$true', 'yes')) {
        $GraphOnly = $true
    } else {
        $GraphOnly = $false
    }

    Write-Host "  CloudEnvironment: '$CloudEnvironment'" -NoNewline
    # Endpoints from https://github.com/microsoft/CSS-Exchange/blob/main/Shared/AzureFunctions/Get-CloudServiceEndpoint.ps1
    # Environment names must match https://learn.microsoft.com/en-us/dotnet/api/microsoft.identity.client.azurecloudinstance?view=msal-dotnet-latest
    switch ($CloudEnvironment) {
        { $_ -iin @('Public', 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC') } {
            $CloudEnvironmentEnvironmentName = 'AzurePublic'
            $CloudEnvironmentGraphApiEndpoint = 'https://graph.microsoft.com'
            $CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook.office.com'
            $CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.outlook.com'
            $CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.com'
            break
        }

        { $_ -iin @('AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4') } {
            $CloudEnvironmentEnvironmentName = 'AzureUSGovernment'
            $CloudEnvironmentGraphApiEndpoint = 'https://graph.microsoft.us'
            $CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook.office365.us'
            $CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.office365.us'
            $CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.us'
            break
        }

        { $_ -iin @('AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5') } {
            $CloudEnvironmentEnvironmentName = 'AzureUSGovernment'
            $CloudEnvironmentGraphApiEndpoint = 'https://dod-graph.microsoft.us'
            $CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook-dod.office365.us'
            $CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s-dod.office365.us'
            $CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.us'
            break
        }

        { $_ -iin @('China', 'AzureChina', 'ChinaCloud', 'AzureChinaCloud') } {
            $CloudEnvironmentEnvironmentName = 'AzureChina'
            $CloudEnvironmentGraphApiEndpoint = 'https://microsoftgraph.chinacloudapi.cn'
            $CloudEnvironmentExchangeOnlineEndpoint = 'https://partner.outlook.cn'
            $CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.partner.outlook.cn'
            $CloudEnvironmentAzureADEndpoint = 'https://login.partner.microsoftonline.cn'
            break
        }
    }

    Write-Host "  GraphConfigFile: '$GraphConfigFile'" -NoNewline
    if ($GraphConfigFile) {
        ConvertPath ([ref]$GraphConfigFile)
        CheckPath $GraphConfigFile
        foreach ($line in @(Get-Content -LiteralPath $GraphConfigFile -Encoding UTF8)) {
            Write-Verbose "    $($line)"
        }
    } else {
        Write-Host
    }

    Write-Host "  SimulateAndDeploy: '$SimulateAndDeploy'"
    if ($SimulateAndDeploy -iin (1, '1', 'true', '$true', 'yes')) {
        $SimulateAndDeploy = $true
    } else {
        $SimulateAndDeploy = $false
    }

    Write-Host "  SimulateAndDeployGraphCredentialFile: '$SimulateAndDeployGraphCredentialFile'" -NoNewline
    if ($SimulateAndDeployGraphCredentialFile) {
        ConvertPath ([ref]$SimulateAndDeployGraphCredentialFile)
        CheckPath $SimulateAndDeployGraphCredentialFile
    } else {
        Write-Host
    }

    Write-Host "  ReplacementVariableConfigFile: '$ReplacementVariableConfigFile'" -NoNewline
    if ($ReplacementVariableConfigFile) {
        ConvertPath ([ref]$ReplacementVariableConfigFile)
        CheckPath $ReplacementVariableConfigFile
        foreach ($line in @(Get-Content -LiteralPath $ReplacementVariableConfigFile -Encoding UTF8)) {
            Write-Verbose "    $($line)"
        }
    } else {
        Write-Host
    }

    Write-Host "  MoveCSSInline: '$MoveCSSInline'"
    if ($MoveCSSInline -iin (1, '1', 'true', '$true', 'yes')) {
        $MoveCSSInline = $true
    } else {
        $MoveCSSInline = $false
    }

    if ($MoveCSSInline) {
        $script:PreMailerNetModulePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))

        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\PreMailer.net\net462')) -Destination $script:PreMailerNetModulePath -Recurse -ErrorAction SilentlyContinue

        Get-ChildItem $script:PreMailerNetModulePath -Recurse | Unblock-File
    }

    Write-Host "  EmbedImagesInHtml: '$EmbedImagesInHtml'"
    if ($EmbedImagesInHtml -iin (1, '1', 'true', '$true', 'yes')) {
        $EmbedImagesInHtml = $true
    } else {
        $EmbedImagesInHtml = $false
    }

    Write-Host "  EmbedImagesInHtmlAdditionalSignaturePath: '$EmbedImagesInHtmlAdditionalSignaturePath'"
    if ($EmbedImagesInHtmlAdditionalSignaturePath -iin (1, '1', 'true', '$true', 'yes')) {
        $EmbedImagesInHtmlAdditionalSignaturePath = $true
    } else {
        $EmbedImagesInHtmlAdditionalSignaturePath = $false
    }

    Write-Host "  CreateRtfSignatures: '$CreateRtfSignatures'"
    if ($CreateRtfSignatures -iin (1, '1', 'true', '$true', 'yes')) {
        $CreateRtfSignatures = $true
    } else {
        $CreateRtfSignatures = $false
    }

    Write-Host "  CreateTxtSignatures: '$CreateTxtSignatures'"
    if ($CreateTxtSignatures -iin (1, '1', 'true', '$true', 'yes')) {
        $CreateTxtSignatures = $true
    } else {
        $CreateTxtSignatures = $false
    }

    Write-Host "  DocxHighResImageConversion: '$DocxHighResImageConversion'"
    if ($DocxHighResImageConversion -iin (1, '1', 'true', '$true', 'yes')) {
        $DocxHighResImageConversion = $true
    } else {
        $DocxHighResImageConversion = $false
    }

    Write-Host "  DeleteUserCreatedSignatures: '$DeleteUserCreatedSignatures'"
    if ($DeleteUserCreatedSignatures -iin (1, '1', 'true', '$true', 'yes')) {
        $DeleteUserCreatedSignatures = $true
    } else {
        $DeleteUserCreatedSignatures = $false
    }

    Write-Host "  DeleteScriptCreatedSignaturesWithoutTemplate: '$DeleteScriptCreatedSignaturesWithoutTemplate'"
    if ($DeleteScriptCreatedSignaturesWithoutTemplate -iin (1, '1', 'true', '$true', 'yes')) {
        $DeleteScriptCreatedSignaturesWithoutTemplate = $true
    } else {
        $DeleteScriptCreatedSignaturesWithoutTemplate = $false
    }

    Write-Host "  SignaturesForAutomappedAndAdditionalMailboxes: '$SignaturesForAutomappedAndAdditionalMailboxes'"
    if ($SignaturesForAutomappedAndAdditionalMailboxes -iin (1, '1', 'true', '$true', 'yes')) {
        $SignaturesForAutomappedAndAdditionalMailboxes = $true
    } else {
        $SignaturesForAutomappedAndAdditionalMailboxes = $false
    }

    Write-Host "  AdditionalSignaturePath: '$AdditionalSignaturePath'" -NoNewline
    if ($AdditionalSignaturePath) {
        ConvertPath ([ref]$AdditionalSignaturePath)
        checkpath $AdditionalSignaturePath -create
    } else {
        Write-Host
    }

    Write-Host "  SimulateUser: '$SimulateUser'"

    $tempSimulateMailboxes = $SimulateMailboxes
    [string[]]$SimulateMailboxes = $null
    foreach ($tempSimulateMailbox in $tempSimulateMailboxes) {
        $SimulateMailboxes += $tempSimulateMailbox.Address
    }

    Write-Host ('  SimulateMailboxes: ' + ('''' + $($SimulateMailboxes -join ''', ''') + ''''))

    Write-Host "  SimulateTime: '$($SimulateTime)'$(if ($SimulateTime) {" ($([DateTime]::ParseExact($SimulateTime, 'yyyyMMddHHmm', $null)))"})"

    Write-Host "  DisableRoamingSignatures: '$($DisableRoamingSignatures)'"
    if ($DisableRoamingSignatures -iin (1, '1', 'true', '$true', 'yes')) {
        $DisableRoamingSignatures = $true
    } elseif ($DisableRoamingSignatures -iin (0, '0', 'false', '$false', 'no')) {
        $DisableRoamingSignatures = $false
    } else {
        $DisableRoamingSignatures = $null
    }

    Write-Host "  MirrorLocalSignaturesToCloud: '$($MirrorLocalSignaturesToCloud)'"
    if ($MirrorLocalSignaturesToCloud -iin (1, '1', 'true', '$true', 'yes')) {
        $MirrorLocalSignaturesToCloud = $true
    } else {
        $MirrorLocalSignaturesToCloud = $false
    }

    switch ($WordProcessPriority) {
        'Idle' { $WordProcessPriorityText = $WordProcessPriority; $WordProcessPriority = 64; break }
        'BelowNormal' { $WordProcessPriorityText = $WordProcessPriority; $WordProcessPriority = 16384; break }
        'Normal' { $WordProcessPriorityText = $WordProcessPriority; $WordProcessPriority = 32; break }
        'AboveNormal' { $WordProcessPriorityText = $WordProcessPriority; $WordProcessPriority = 32768; break }
        'High' { $WordProcessPriorityText = $WordProcessPriority; $WordProcessPriority = 128; break }
        'Realtime' { $WordProcessPriorityText = $WordProcessPriority; $WordProcessPriority = 256; break }
        64 { $WordProcessPriorityText = 'Idle'; break }
        16384 { $WordProcessPriorityText = 'BelowNormal'; break }
        32 { $WordProcessPriorityText = 'Normal'; break }
        32768 { $WordProcessPriorityText = 'AboveNormal'; break }
        128 { $WordProcessPriorityText = 'High'; break }
        256 { $WordProcessPriorityText = 'Realtime'; break }
        default { $WordProcessPriority = 32; $WordProcessPriorityText = 'Normal' }
    }
    Write-Host "  WordProcessPriority: '$($WordProcessPriorityText)' ('$($WordProcessPriority)')"

    switch ((Get-Process -Id $pid).PriorityClass.ToString()) {
        'Idle' { $script:ScriptProcessPriorityOriginal = 64; break }
        'BelowNormal' { $script:ScriptProcessPriorityOriginal = 16384; break }
        'Normal' { $script:ScriptProcessPriorityOriginal = 32; break }
        'AboveNormal' { $script:ScriptProcessPriorityOriginal = 32768; break }
        'High' { $script:ScriptProcessPriorityOriginal = 128; break }
        'Realtime' { $script:ScriptProcessPriorityOriginal = 256; break }
        default { $script:ScriptProcessPriorityOriginal = 32 }
    }
    switch ($ScriptProcessPriority) {
        'Idle' { $ScriptProcessPriorityText = $ScriptProcessPriority; $ScriptProcessPriority = 64; break }
        'BelowNormal' { $ScriptProcessPriorityText = $ScriptProcessPriority; $ScriptProcessPriority = 16384; break }
        'Normal' { $ScriptProcessPriorityText = $ScriptProcessPriority; $ScriptProcessPriority = 32; break }
        'AboveNormal' { $ScriptProcessPriorityText = $ScriptProcessPriority; $ScriptProcessPriority = 32768; break }
        'High' { $ScriptProcessPriorityText = $ScriptProcessPriority; $ScriptProcessPriority = 128; break }
        'Realtime' { $ScriptProcessPriorityText = $ScriptProcessPriority; $ScriptProcessPriority = 256; break }
        64 { $ScriptProcessPriorityText = 'Idle'; break }
        16384 { $ScriptProcessPriorityText = 'BelowNormal'; break }
        32 { $ScriptProcessPriorityText = 'Normal'; break }
        32768 { $ScriptProcessPriorityText = 'AboveNormal'; break }
        128 { $ScriptProcessPriorityText = 'High'; break }
        256 { $ScriptProcessPriorityText = 'Realtime'; break }
        default { $ScriptProcessPriority = 32; $ScriptProcessPriorityText = 'Normal' }
    }
    Write-Host "  ScriptProcessPriority: '$($ScriptProcessPriorityText)' ('$($ScriptProcessPriority)')"
    $null = Get-CimInstance Win32_process -Filter "ProcessId = ""$PID""" | Invoke-CimMethod -Name SetPriority -Arguments @{Priority = $ScriptProcessPriority }

    Write-Host "  BenefactorCircleLicenseFile: '$BenefactorCircleLicenseFile'" -NoNewline
    if ($BenefactorCircleLicenseFile) {
        ConvertPath ([ref]$BenefactorCircleLicenseFile)
        CheckPath $BenefactorCircleLicenseFile
        $script:BenefactorCircleLicenseFilePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))
        Copy-Item -Path $BenefactorCircleLicenseFile -Destination $script:BenefactorCircleLicenseFilePath -Force
        Unblock-File -LiteralPath $script:BenefactorCircleLicenseFilePath
        Import-Module -Name $script:BenefactorCircleLicenseFilePath -Force -ErrorAction Stop
    } else {
        Write-Host
    }

    Write-Host "  BenefactorCircleId: '$BenefactorCircleId'"


    if ($SimulateUser) {
        Write-Host
        Write-Host 'Simulation mode enabled' -ForegroundColor Yellow

        if (-not $AdditionalSignaturePath) {
            Write-Host '  Simulation mode requires AdditionalSignaturePath. Exit.' -ForegroundColor Red
            exit 1
        }

        if (-not $SimulateMailboxes) {
            Write-Host '  SimulateUser is defined, but not SimulateMailboxes.' -ForegroundColor Yellow
        }
    } else {
        if ($SimulateMailboxes) {
            Write-Host
            Write-Host 'SimulateMailboxes is defined, but not SimulateUser. Exit.' -ForegroundColor Red
            exit 1
        }

        if ($SimulateTime) {
            Write-Host
            Write-Host 'SimulateTime is defined, but not SimulateUser. Exit.' -ForegroundColor Red
            exit 1
        }
    }

    if ($BenefactorCircleLicenseFile -and $BenefactorCircleId) {
        Write-Host
        Write-Host "Benefactor Circle license information @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $result = [SetOutlookSignatures.BenefactorCircle]::GetLicenseDetails()
        if ($result -ilike 'License is not valid: *') {
            $result -split '\r?\n' | ForEach-Object { Write-Host "  $($_)" -ForegroundColor Red }
            Write-Host
            Write-Host 'Continuing without Benefactor Circle exclusive features.' -ForegroundColor Red
            Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Red
        } else {
            $result -split '\r?\n' | ForEach-Object {
                if ($_.trim().startswith('Warning!')) {
                    Write-Host "  $($_)" -ForegroundColor Yellow
                } else {
                    Write-Host "  $($_)"
                }
            }
        }
    } elseif ($BenefactorCircleLicenseFile -or $BenefactorCircleId) {
        Write-Host
        Write-Host 'Benefactor Circle Id and license file must both be set for access to exclusive features.' -ForegroundColor Red
        Write-Host 'Continuing without these exclusive features.' -ForegroundColor Red
        Write-Host "Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Red
    }

    if (
        (-not ($BenefactorCircleLicenseFile -and $BenefactorCircleId)) -or
        ($result -ilike 'License is not valid: *')
    ) {
        Write-Host
        Write-Host @'
Thank you for using the free and open-source version of Set-OutlookSignatures!

If some or all of your mailboxes are hosted in the cloud, make sure you read the 'Quick Start Guide' first.
This guide is part of the documention file '.\docs\README.html'.
The '.\docs' folder also contains additional information.

Go to 'https://github.com/Set-OutlookSignatures/Set-OutlookSignatures' for new releases or to report issues.
'@ -ForegroundColor Green

        if ($GitHubReleasesNewer) {
            Write-Host "At least one release newer than $($ScriptVersion) is available: $(@($GitHubReleasesNewer[0, -1] | Select-Object -Unique) -join $(if($GitHubReleasesNewer.Count -gt 2) { ', ..., ' } else {', '}))" -ForegroundColor Yellow
        }

        Write-Host @'

You may be interested in some of the software located in the '.\sample code' folder.

Unlock additional features with a Benefactor Circle membership:
  - Time-based campaigns by assigning time range constraints to templates
  - Signatures for automapped and additional mailboxes
  - Set current user Outlook Web signature (classic Outlook Web signature and roaming signatures)
  - Download and upload roaming signatures
  - Set current user out-of-office replies
  - Delete signatures created by the software, for which the templates no longer exist or apply
  - Delete user-created signatures
  - Additional signature path (when used outside of simulation mode)
  - High resolution images from DOCX templates
  - Digitally signed components for tamper protection and easy integration into locked-down environments

Learn more about Set-OutlookSignatures Benefactor Circle in '.\docs\Benefactor Circle.html',
or visit 'https://explicitconsulting.at'.
'@ -ForegroundColor Green

        Write-Host

        30..0 | ForEach-Object {
            Write-Host ("`rAutomatically continue in {0:00} seconds, or stop program with Ctrl+C" -f $_) -NoNewline

            if ($_ -gt 0) { Start-Sleep -Seconds 1 }
        }

        Write-Host
    }


    Write-Host
    Write-Host "Get basic Outlook and Word information @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $OutlookProfiles = @()
    $OutlookUseNewOutlook = $null

    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled, skip Outlook checks' -ForegroundColor Yellow
    } else {
        if ($(Get-Command -Name 'Get-AppPackage' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
            $NewOutlook = Get-AppPackage -Name 'Microsoft.OutlookForWindows' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        } else {
            $NewOutlook = $null
        }

        Write-Host '  Outlook'
        $OutlookRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Outlook.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))

        try {
            # [Microsoft.Win32.RegistryView]::Registry32 makes sure view the registry as a 32 bit application would
            # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
            # Covers:
            #   Office x86 on Windows x86
            #   Office x86 on Windows x64
            #   Any PowerShell process bitness
            $OutlookFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry32)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '') -ErrorAction Stop
        } catch {
            try {
                # [Microsoft.Win32.RegistryView]::Registry64 makes sure we view the registry as a 64 bit application would
                # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
                # Covers:
                #   Office x64 on Windows x64
                #   Any PowerShell process bitness
                $OutlookFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry64)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '') -ErrorAction Stop
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
                $OutlookBitness = 'Error'
                $OutlookFileVersion = $null
            }
        } else {
            $OutlookBitness = $null
            $OutlookFileVersion = $null
        }

        if ($OutlookRegistryVersion.major -eq 0) {
            $OutlookRegistryVersion = $null
        } elseif ($OutlookRegistryVersion.major -gt 16) {
            Write-Host "    Outlook version $OutlookRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
            exit 1
        } elseif ($OutlookRegistryVersion.major -eq 16) {
            $OutlookRegistryVersion = '16.0'
        } elseif ($OutlookRegistryVersion.major -eq 15) {
            $OutlookRegistryVersion = '15.0'
        } elseif ($OutlookRegistryVersion.major -eq 14) {
            $OutlookRegistryVersion = '14.0'
        } elseif ($OutlookRegistryVersion.major -lt 14) {
            Write-Host "    Outlook version $OutlookRegistryVersion is older than Outlook 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
            exit 1
        }

        Write-Host "    Set 'Send Pictures With Document' registry value to '1'"
        $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Options\Mail" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'Send Pictures With Document' -Type DWORD -Value 1 -Force

        if (($DisableRoamingSignatures -in @($true, $false)) -and ($OutlookFileVersion -ge '16.0.0.0')) {
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

            foreach ($RegistryFolder in (
                    "registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                    "registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                    "registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                    "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup"
                )) {

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
            } else {
                $OutlookUseNewOutlook = $false
            }
        } else {
            $OutlookDefaultProfile = $null
            $OutlookDisableRoamingSignatures = $null
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
            Write-Host '      Consider parameter ''EmbedImagesInHtml false'' to avoid problems with images in templates.' -ForegroundColor Yellow
            Write-Host '      Microsoft supports Outlook 2013 until April 2023, older versions are already out of support.' -ForegroundColor Yellow
        }
        Write-Host "    Bitness: $OutlookBitness"
        Write-Host "    Default profile: $OutlookDefaultProfile"
        Write-Host "    Is C2R Beta: $OutlookIsBetaversion"
        Write-Host "    DisableRoamingSignatures: $OutlookDisableRoamingSignatures"
        Write-Host "    UseNewOutlook: $OutlookUseNewOutlook"

        Write-Host '  New Outlook'
        Write-Host "    Version: $($NewOutlook.Version)"
        Write-Verbose "    Status: $($NewOutlook.Status)"
        Write-Host "    UseNewOutlook: $OutlookUseNewOutlook"
    }

    Write-Host '  Word'
    $WordRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Word.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
    if ($WordRegistryVersion.major -eq 0) {
        $WordRegistryVersion = $null
    } elseif ($WordRegistryVersion.major -gt 16) {
        Write-Host "    Word version $WordRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
        exit 1
    } elseif ($WordRegistryVersion.major -eq 16) {
        $WordRegistryVersion = '16.0'
    } elseif ($WordRegistryVersion.major -eq 15) {
        $WordRegistryVersion = '15.0'
    } elseif ($WordRegistryVersion.major -eq 14) {
        $WordRegistryVersion = '14.0'
    } elseif ($WordRegistryVersion.major -lt 14) {
        Write-Host "    Word version $WordRegistryVersion is older than Word 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
        exit 1
    }

    try {
        # [Microsoft.Win32.RegistryView]::Registry32 makes sure view the registry as a 32 bit application would
        # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
        # Covers:
        #   Office x86 on Windows x86
        #   Office x86 on Windows x64
        #   Any PowerShell process bitness
        $WordFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry32)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '') -ErrorAction Stop
    } catch {
        try {
            # [Microsoft.Win32.RegistryView]::Registry64 makes sure we view the registry as a 64 bit application would
            # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
            # Covers:
            #   Office x64 on Windows x64
            #   Any PowerShell process bitness
            $WordFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry64)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '') -ErrorAction Stop
        } catch {
            $WordFilePath = $null
        }
    }

    if ($WordFilePath) {
        Write-Host "    Set 'DontUseScreenDpiOnOpen' registry value to '1'"
        $null = "HKCU:\Software\Microsoft\Office\$($WordRegistryVersion)\Word\Options" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DontUseScreenDpiOnOpen' -Type DWORD -Value 1 -Force

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

    Write-Host "    Registry version: $WordRegistryVersion"
    Write-Host "    File version: $WordFileVersion"
    Write-Host "    Bitness: $WordBitness"


    Write-Host
    Write-Host "Get Outlook signature file path(s) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $SignaturePaths = @()

    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled. Skip task, use AdditionalSignaturePath instead' -ForegroundColor Yellow
        if ($AdditionalSignaturePath) {
            $SignaturePaths += $AdditionalSignaturePath
        }
    } elseif ($OutlookProfiles -and ($OutlookUseNewOutlook -ne $true)) {
        $x = (Get-ItemProperty "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\common\general" -ErrorAction SilentlyContinue).'Signatures'

        if ($x) {
            Push-Location ((Join-Path -Path ($env:AppData) -ChildPath 'Microsoft'))
            $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($x))

            if (Test-Path $x -IsValid) {
                if (-not (Test-Path $x -type container)) {
                    New-Item -Path $x -ItemType directory -Force | Out-Null
                }

                if ($x -inotin $SignaturePaths) {
                    $SignaturePaths += $x
                    Write-Host "  $x"
                }
            }

            Pop-Location
        }
    } else {
        $SignaturePaths = @(((New-Item -ItemType Directory (Join-Path -Path $script:tempDir -ChildPath ((New-Guid).guid))).fullname))
        Write-Host "  $($SignaturePaths[-1]) (Outlook Web/New Outlook)"
    }

    # If Outlook is installed, synch profile folders anyway
    # Also makes sure that signatures are already there when starting Outlook for the first time
    if ((-not $SimulateUser) -and $OutlookFileVersion) {
        $x = (Get-ItemProperty "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\common\general" -ErrorAction SilentlyContinue).'Signatures'

        if ($x) {
            Push-Location ((Join-Path -Path ($env:AppData) -ChildPath 'Microsoft'))
            $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($x))

            if (Test-Path $x -IsValid) {
                if (-not (Test-Path $x -type container)) {
                    New-Item -Path $x -ItemType directory -Force | Out-Null
                }

                if ($x -inotin $SignaturePaths) {
                    $SignaturePaths += $x
                    Write-Host "  $x"
                }
            }

            Pop-Location
        }

        $SignaturePaths = @($SignaturePaths | Select-Object -Unique)
    }


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
            $UserForest = (([ADSI]"LDAP://$(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1) -split ',DC=')[1..999] -join '.')/RootDSE").rootDomainNamingContext -ireplace [Regex]::Escape('DC='), '' -ireplace [Regex]::Escape(','), '.').tolower()
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

                $TrustedDomains = @($Search.FindAll())

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

                # Internal trusts
                foreach ($TrustedDomain in $TrustedDomains) {
                    if (($TrustedDomain.properties.trustattributes -eq 32) -and ($TrustedDomain.properties.name -ine $UserForest) -and (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower()))) {
                        Write-Host "    Child domain: $($TrustedDomain.properties.name.tolower())"

                        if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                            $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $UserForest.tolower())
                        }
                    }
                }

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
                            Write-Host '    Trusted domain/forest already in list.' -ForegroundColor Yellow
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


                Write-Host
                Write-Host "Check trusts for open LDAP port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                CheckADConnectivity @(@(@($TrustsToCheckForGroups) + @($LookupDomainsToTrusts.GetEnumerator() | ForEach-Object { $_.Name })) | Select-Object -Unique) 'LDAP' '  ' | Out-Null


                Write-Host
                Write-Host "Check trusts for open Global Catalog port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                CheckADConnectivity $TrustsToCheckForGroups 'GC' '  ' | Out-Null
            } else {
                Write-Host '  Problem connecting to logged-in user''s Active Directory (no error message, but forest root domain name is empty).' -ForegroundColor Yellow
                Write-Host '  Assuming Graph/Entra ID/Azure AD from now on.' -ForegroundColor Yellow
                $GraphOnly = $true
            }
        } catch {
            $y = ''
            Write-Verbose $error[0]
            Write-Host '  Problem connecting to logged-in user''s Active Directory (see verbose stream for error message).' -ForegroundColor Yellow
            Write-Host '  Assuming Graph/Entra ID/Azure AD from now on.' -ForegroundColor Yellow
            $GraphOnly = $true
        }
    } else {
        Write-Host "  Parameter GraphOnly set to '$GraphOnly', ignore user's Active Directory in favor of Graph/Entra ID/Azure AD."
    }


    Write-Host
    Write-Host "Get AD properties of currently logged-in user and assigned manager @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if (-not $SimulateUser) {
        Write-Host '  Currently logged-in user'
    } else {
        Write-Host "  Simulate '$SimulateUser' as currently logged-in user" -ForegroundColor Yellow
    }

    if ($GraphOnly -eq $false) {
        if ($null -ne $TrustsToCheckForGroups[0]) {
            try {
                if (-not $SimulateUser) {
                    $Search.SearchRoot = "GC://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$(([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName)))"
                    $ADPropsCurrentUser = $Search.FindOne().Properties
                    $ADPropsCurrentUser = [hashtable]::new($ADPropsCurrentUser, [StringComparer]::OrdinalIgnoreCase)

                    $Search.SearchRoot = "LDAP://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$(([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName)))"
                    $ADPropsCurrentUserLdap = $Search.FindOne().Properties
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
                        $SimulateUserDN = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1)
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
                        Remove-Variable -Name 'objTrans'
                        Remove-Variable -Name 'objNT'

                        $Search.SearchRoot = "GC://$(($SimulateUserDN -split ',DC=')[1..999] -join '.')"
                        $Search.Filter = "((distinguishedname=$SimulateUserDN))"
                        $ADPropsCurrentUser = $Search.FindOne().Properties
                        $ADPropsCurrentUser = [hashtable]::new($ADPropsCurrentUser, [StringComparer]::OrdinalIgnoreCase)

                        $Search.SearchRoot = "LDAP://$(($SimulateUserDN -split ',DC=')[1..999] -join '.')"
                        $Search.Filter = "((distinguishedname=$SimulateUserDN))"
                        $ADPropsCurrentUserLdap = $Search.FindOne().Properties
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
                        Write-Verbose $error[0]
                        Write-Host "    Simulation user '$($SimulateUser)' not found. Exit." -ForegroundColor REd
                        exit 1
                    }
                }
            } catch {
                $ADPropsCurrentUser = $null
                Write-Host '    Problem connecting to Active Directory, or user is a local user. Exit.' -ForegroundColor Red
                $error[0]
                exit 1
            }
        }
    }

    if (
        ($GraphOnly -eq $true) -or
        (($GraphOnly -eq $false) -and ($ADPropsCurrentUser.msexchrecipienttypedetails -ge 2147483648) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true))) -or
        (($GraphOnly -eq $false) -and ($null -eq $ADPropsCurrentUser)) -or
        ($OutlookUseNewOutlook -eq $true)
    ) {
        Write-Host "    Set up environment for connection to Microsoft Graph @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $script:CurrentUser = (Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value)\IdentityCache\$(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value)" -Name 'UserName' -ErrorAction SilentlyContinue)
        $script:MsalModulePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))
        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\msal.ps')) -Destination (Join-Path -Path $script:MsalModulePath -ChildPath 'msal.ps') -Recurse -ErrorAction SilentlyContinue
        Get-ChildItem $script:MsalModulePath -Recurse | Unblock-File
        try {
            Import-Module (Join-Path -Path $script:MsalModulePath -ChildPath 'msal.ps') -Force -ErrorAction Stop
        } catch {
            Write-Host '        Problem importing MSAL.PS module. Exit.' -ForegroundColor Red
            $error[0]
            exit 1
        }

        if (Test-Path -Path $GraphConfigFile -PathType Leaf) {
            try {
                Write-Host "      Execute config file '$GraphConfigFile'"
                . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $GraphConfigFile -Encoding UTF8 -Raw)))

                if (-not $GraphUserProperties) {
                    $GraphUserProperties = @()
                }

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
                $GraphUserAttributeMapping['mailnickname'] = 'mailNickname'
                $GraphUserAttributeMapping['objectguid'] = 'id'
                $GraphUserAttributeMapping['objectsid'] = 'securityIdentifier'
                $GraphUserAttributeMapping['onpremisesdomainname'] = 'onPremisesDomainName'
                $GraphUserAttributeMapping['onpremisessecurityidentifier'] = 'onPremisesSecurityIdentifier'
                $GraphUserAttributeMapping['userprincipalname'] = 'userPrincipalName'
            } catch {
                Write-Host "        Problem executing content of '$GraphConfigFile'. Exit." -ForegroundColor Red
                $error[0]
                exit 1
            }
        } else {
            Write-Host "      Problem connecting to or reading from file '$GraphConfigFile'. Exit." -ForegroundColor Red
            exit 1
        }

        if (-not $SimulateAndDeployGraphCredentialFile) {
            Write-Host "      MSAL.PS Graph token cache: '$([TokenCacheHelper]::CacheFilePath)'"
        }

        if (-not $SimulateAndDeployGraphCredentialFile) {
            Write-Verbose "      Current user: $($script:CurrentUser)"
        }

        $GraphToken = GraphGetToken

        if ($GraphToken.error -eq $false) {
            Write-Verbose "        Graph Token metadata: $((ParseJwtToken $GraphToken.AccessToken) | ConvertTo-Json)"

            if (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true)) {
                Write-Verbose "        EXO Token metadata: $((ParseJwtToken $GraphToken.AccessTokenExo) | ConvertTo-Json)"

                if (-not $($GraphToken.AccessTokenExo)) {
                    Write-Host '      Problem connecting to Exchange Online with Graph token. Exit.' -ForegroundColor Red
                    exit 1
                }
            }

            if ($SimulateAndDeployGraphCredentialFile) {
                Write-Verbose "        App Graph Token metadata: $((ParseJwtToken $GraphToken.AppAccessToken) | ConvertTo-Json)"
                Write-Verbose "        App EXO Token metadata: $((ParseJwtToken $GraphToken.AppAccessTokenExo) | ConvertTo-Json)"
            }

            if ($SimulateUser) {
                $script:CurrentUser = $SimulateUser
            }

            if ($null -eq $script:CurrentUser) {
                $script:CurrentUser = (GraphGetMe).me.userprincipalname
            }

            $error.clear()

            $x = (GraphGetUserProperties $script:CurrentUser)

            if (($x.error -eq $false) -and ($x.properties.id)) {
                $AADProps = $x.properties
                $ADPropsCurrentUser = [PSCustomObject]@{}

                foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                    $z = $AADProps

                    foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                        $z = $z.$y
                    }

                    $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z
                }

                $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $script:CurrentUser).photo
                $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'manager' -Value (GraphGetUserManager $script:CurrentUser).properties.userprincipalname
            } else {
                Write-Host "      Problem getting data for '$($script:CurrentUser)' from Microsoft Graph. Exit." -ForegroundColor Red
                $error[0]
                exit 1
            }
        } else {
            Write-Host '      Problem connecting to Microsoft Graph. Exit.' -ForegroundColor Red
            $GraphToken.error
            exit 1
        }
    }

    if ($ADPropsCurrentUser.distinguishedname) {
        Write-Host "    $($ADPropsCurrentUser.distinguishedname)"
    } elseif ($ADPropsCurrentUser.userprincipalname) {
        Write-Host "    $($ADPropsCurrentUser.userprincipalname.tolower())"
    } elseif ($ADPropsCurrentUser.mail) {
        Write-Host "    $($ADPropsCurrentUser.mail.tolower())"
    }

    Write-Verbose "    distinguishedname: $($ADPropsCurrentUser.distinguishedname)"
    Write-Verbose "    userprincipalname: $($ADPropsCurrentUser.userprincipalname)"
    Write-Verbose "    mail: $($ADPropsCurrentUser.mail)"

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

                    $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z
                }

                $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsCurrentUserManager.userprincipalname).photo
                $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name 'manager' -Value $null
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
                    $ADPropsCurrentUserManager = $Search.FindOne().Properties
                    $ADPropsCurrentUserManager = [hashtable]::new($ADPropsCurrentUserManager, [StringComparer]::OrdinalIgnoreCase)


                    $Search.SearchRoot = "LDAP://$(($ADPropsCurrentUser.manager -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$($ADPropsCurrentUser.manager)))"
                    $ADPropsCurrentUserManagerLdap = $Search.FindOne().Properties
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
        if ($ADPropsCurrentUserManager.distinguishedname) {
            Write-Host "    $($ADPropsCurrentUserManager.distinguishedname)"
        } elseif ($ADPropsCurrentUserManager.userprincipalname) {
            Write-Host "    $($ADPropsCurrentUserManager.userprincipalname)"
        } elseif ($ADPropsCurrentUserManager.mail) {
            Write-Host "    $($ADPropsCurrentUserManager.mail)"
        }

        Write-Verbose "    distinguishedname: $($ADPropsCurrentUserManager.distinguishedname)"
        Write-Verbose "    userprincipalname: $($ADPropsCurrentUserManager.userprincipalname)"
        Write-Verbose "    mail: $($ADPropsCurrentUserManager.mail)"
    } else {
        Write-Host '    No manager found'
    }


    Write-Host
    Write-Host "Get email addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $MailAddresses = @()
    $RegistryPaths = @()
    $LegacyExchangeDNs = @()

    if ($SimulateUser -and $SimulateMailboxes) {
        Write-Host '  Simulation mode enabled, use SimulateMailboxes as mailbox list' -ForegroundColor Yellow
        for ($i = 0; $i -lt $SimulateMailboxes.count; $i++) {
            $MailAddresses += $SimulateMailboxes[$i].ToLower()
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''
        }
    } elseif (($OutlookProfiles) -and ($OutlookUseNewOutlook -ne $true)) {
        Write-Host '  Get email addresses from Outlook'

        foreach ($OutlookProfile in $OutlookProfiles) {
            Write-Host "    Profile '$($OutlookProfile)'"
            foreach ($RegistryFolder in @(Get-ItemProperty "hkcu:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$($OutlookProfile)\9375CFF0413111d3B88A00104B2A6676\*" -ErrorAction SilentlyContinue | Where-Object { if ($OutlookFileVersion -ge '16.0.0.0') { ($_.'Account Name' -like '*@*.*') } else { (($_.'Account Name' -join ',') -like '*,64,*,46,*') } })) {
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
                if (-not $BenefactorCircleLicenseFile) {
                    Write-Host "    The 'SignaturesForAutomappedAndAdditionalMailboxes' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                    Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SignaturesForAutomappedAndAdditionalMailboxes()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '      Error finding automapped and additional mailboxes.' -ForegroundColor Yellow
                        Write-Host "      $FeatureResult" -ForegroundColor Yellow
                    }
                }
            }
        }
    } else {
        if ($OutlookUseNewOutlook -eq $true) {
            Write-Host '  Get email addresses from New Outlook and Outlook Web, as New Outlook is set as default'
        } else {
            Write-Host '  Get email addresses from Outlook Web'
        }

        $OutlookProfiles = @()
        $OutlookDefaultProfile = $null

        $script:CurrentUserDummyMailbox = $true

        if ($OutlookUseNewOutlook -eq $true) {
            $x = @(
                @((Get-Content -Path $(Join-Path -Path $env:LocalAppData -ChildPath '\Microsoft\Olk\UserSettings.json') -Force -Encoding utf8 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ConvertFrom-Json).Identities.IdentityMap.PSObject.Properties | Select-Object -Unique | Where-Object { $_.name -match '(\S+?)@(\S+?)\.(\S+?)' }) | ForEach-Object {
                    if ((Get-Content -Path $(Join-Path -Path $env:LocalAppData -ChildPath "\Microsoft\OneAuth\accounts\$($_.Value)") -Force -Encoding utf8 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ConvertFrom-Json).association_status -ilike '*"com.microsoft.Olk":"associated"*') {
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
            $script:CurrentUserDummyMailbox = $false
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

                    $script:WebServicesDllPath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))

                    try {
                        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\netstandard2.0\Microsoft.Exchange.WebServices.Data.dll')) -Destination $script:WebServicesDllPath -Force
                        Unblock-File -LiteralPath $script:WebServicesDllPath

                        #if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                        #    Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\netcore\Microsoft.Exchange.WebServices.Data.dll')) -Destination $script:WebServicesDllPath -Force
                        #    Unblock-File -LiteralPath $script:WebServicesDllPath
                        #} else {
                        #    Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\netframework\Microsoft.Exchange.WebServices.dll')) -Destination $script:WebServicesDllPath -Force
                        #    Unblock-File -LiteralPath $script:WebServicesDllPath
                        #}
                    } catch {
                        Write-Verbose $_
                    }
                }

                $error.clear()

                try {
                    Import-Module -Name $script:WebServicesDllPath -Force -ErrorAction Stop

                    $exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

                    Write-Host '    Connect to Outlook Web'
                    try {
                        Write-Verbose '      Try Integrated Windows Authentication'

                        if (
                            ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile -and $GraphToken.AppAccessTokenExo) -or
                            $GraphToken.AccessTokenExo
                        ) {
                            throw '        EXO access token available, skip Integrated Windows Authentication'
                        }

                        $exchService.UseDefaultCredentials = $true
                        $exchService.ImpersonatedUserId = $null
                        $exchService.AutodiscoverUrl($MailAddresses[0], { $true }) | Out-Null
                    } catch {
                        try {
                            Write-Verbose $_

                            Write-Verbose '      Try OAuth with Autodiscover'

                            $exchService.UseDefaultCredentials = $false

                            if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) {
                                $exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailAddresses[0])
                                $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AppAccessTokenExo)
                            } else {
                                $exchService.ImpersonatedUserId = $null
                                $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AccessTokenExo)
                            }

                            $exchService.AutodiscoverUrl($MailAddresses[0], { $true }) | Out-Null
                        } catch {
                            Write-Verbose $_

                            Write-Verbose '      Try OAuth with fixed URL'

                            $exchService.UseDefaultCredentials = $false

                            if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) {
                                $exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailAddresses[0])
                                $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AppAccessTokenExo)
                            } else {
                                $exchService.ImpersonatedUserId = $null
                                $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AccessTokenExo)
                            }

                            $exchService.Url = "$($CloudEnvironmentExchangeOnlineEndpoint)/EWS/Exchange.asmx"
                        }
                    }

                    $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchservice, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)

                    if ($Calendar.DisplayName) {
                        $error.clear()

                        if ($SignaturesForAutomappedAndAdditionalMailboxes) {
                            if (-not $BenefactorCircleLicenseFile) {
                                Write-Host "    The 'SignaturesForAutomappedAndAdditionalMailboxes' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                                Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                            } else {
                                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SignaturesForAutomappedAndAdditionalMailboxes()

                                if ($FeatureResult -ne 'true') {
                                    Write-Host '    Error finding automapped and additional mailboxes.' -ForegroundColor Yellow
                                    Write-Host "    $FeatureResult" -ForegroundColor Yellow
                                }
                            }
                        }
                    } else {
                        Write-Host '      Could not connect to Outlook Web, although the EWS DLL threw no error.' -ForegroundColor Red
                        throw
                    }
                } catch {
                    Write-Host "      Error connecting to Outlook Web: $_" -ForegroundColor Red
                }
            }
        }
    }

    if ((($SetCurrentUserOutlookWebSignature -eq $true) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses -inotcontains $ADPropsCurrentUser.mail)) {
        # OOF and/or Outlook web signature must be set, but user does not seem to have a mailbox in Outlook
        # Maybe this is a pure Outlook Web user, so we will add a helper entry
        # This entry fakes the users mailbox in his default Outlook profile, so it gets the highest priority later
        Write-Host "  User's mailbox not found in email address list, but Outlook Web signature and/or OOF message should be set. Add dummy mailbox entry." -ForegroundColor Yellow

        if ($ADPropsCurrentUser.mail) {
            $script:CurrentUserDummyMailbox = $true

            $SignaturePaths = @(((New-Item -ItemType Directory (Join-Path -Path $script:tempDir -ChildPath ((New-Guid).guid))).fullname)) + $SignaturePaths

            $MailAddresses = @($ADPropsCurrentUser.mail.tolower()) + $MailAddresses
            $RegistryPaths = @('') + $RegistryPaths
            $LegacyExchangeDNs = @('') + $LegacyExchangeDNs
        } else {
            Write-Host '      User does not have mail attribute configured' -ForegroundColor Yellow
            $script:CurrentUserDummyMailbox = $false
        }
    } else {
        $script:CurrentUserDummyMailbox = $false
    }

    if ($MailAddresses.count -eq 0) {
        Write-Host
        Write-Host 'No email addresses found, exiting.'
        Write-Host '  In simulation mode, this might be a permission problem.'
        exit 1
    }


    Write-Host
    Write-Host "Get properties of each mailbox @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $ADPropsMailboxes = @()
    $ADPropsMailboxesUserDomain = @()

    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        Write-Host "  $($MailAddresses[$AccountNumberRunning])"

        $UserDomain = ''
        $ADPropsMailboxes += $null
        $ADPropsMailboxesUserDomain += $null

        if ($AccountNumberRunning -eq $MailAddresses.IndexOf($MailAddresses[$AccountNumberRunning])) {
            if ((($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '') -or ($($MailAddresses[$AccountNumberRunning]) -ne ''))) {
                if ($null -ne $TrustsToCheckForGroups[0]) {
                    # Loop through domains until the first one knows the legacyExchangeDN or the proxy address
                    for ($DomainNumber = 0; (($DomainNumber -lt $TrustsToCheckForGroups.count) -and ($UserDomain -eq '')); $DomainNumber++) {
                        if (($TrustsToCheckForGroups[$DomainNumber] -ne '')) {
                            Write-Host "    Search for mailbox user object in domain/forest '$($TrustsToCheckForGroups[$DomainNumber])': " -NoNewline
                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustsToCheckForGroups[$DomainNumber])")
                            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                                $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(|(msexchrecipienttypedetails<=32)(msexchrecipienttypedetails>=2147483648))(msExchMailboxGuid=*)(|(legacyExchangeDN=$($LegacyExchangeDNs[$AccountNumberRunning]))(&(legacyExchangeDN=*)(proxyaddresses=x500:$($LegacyExchangeDNs[$AccountNumberRunning])))))"
                            } elseif (($($MailAddresses[$AccountNumberRunning]) -ne '')) {
                                $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(|(msexchrecipienttypedetails<=32)(msexchrecipienttypedetails>=2147483648))(msExchMailboxGuid=*)(legacyExchangeDN=*)(proxyaddresses=smtp:$($MailAddresses[$AccountNumberRunning])))"
                            }
                            $u = $Search.FindAll()
                            if ($u.count -eq 0) {
                                Write-Host 'Not found'
                            } elseif ($u.count -gt 1) {
                                Write-Host 'Ignore due to multiple matches' -ForegroundColor Red
                                foreach ($SingleU in $u) {
                                    Write-Host "      $($SingleU.path)" -ForegroundColor Yellow
                                }
                                $LegacyExchangeDNs[$AccountNumberRunning] = ''
                                $MailAddresses[$AccountNumberRunning] = ''
                                $UserDomain = $null
                            } else {
                                $Search.SearchRoot = "GC://$(($(([adsi]"$($u[0].path)").distinguishedname) -split ',DC=')[1..999] -join '.')"
                                $Search.Filter = "((distinguishedname=$(([adsi]"$($u[0].path)").distinguishedname)))"
                                $ADPropsMailboxes[$AccountNumberRunning] = $Search.FindOne().Properties
                                $ADPropsMailboxes[$AccountNumberRunning] = [hashtable]::new($ADPropsMailboxes[$AccountNumberRunning], [StringComparer]::OrdinalIgnoreCase)

                                $Search.SearchRoot = "LDAP://$(($(([adsi]"$($u[0].path)").distinguishedname) -split ',DC=')[1..999] -join '.')"
                                $Search.Filter = "((distinguishedname=$(([adsi]"$($u[0].path)").distinguishedname)))"
                                $tempLdap = $Search.FindOne().Properties
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
                                Write-Host 'Found'
                                Write-Host "      distinguishedname: $($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                            }
                        }
                    }

                    if (-not $ADPropsMailboxes[$AccountNumberRunning]) {
                        $LegacyExchangeDNs[$AccountNumberRunning] = ''
                        $UserDomain = $null
                    }
                } else {
                    $AADProps = (GraphGetUserProperties $($MailAddresses[$AccountNumberRunning])).properties

                    $ADPropsMailboxes[$AccountNumberRunning] = [PSCustomObject]@{}

                    if ($AADProps) {
                        foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                            $z = $AADProps

                            foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                                $z = $z.$y
                            }

                            $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z
                        }

                        $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsMailboxes[$AccountNumberRunning].userprincipalname).photo
                        $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'manager' -Value (GraphGetUserManager $ADPropsMailboxes[$AccountNumberRunning].userprincipalname).properties.userprincipalname

                        if (-not $LegacyExchangeDNs[$AccountNumberRunning]) {
                            $LegacyExchangeDNs[$AccountNumberRunning] = 'dummy'
                        }

                        $MailAddresses[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].mail.tolower()
                    } else {
                        $LegacyExchangeDNs[$AccountNumberRunning] = ''
                        $UserDomain = $null
                    }
                }
            } else {
                $ADPropsMailboxes[$AccountNumberRunning] = $null
            }
        } else {
            Write-Host '    Mailbox user object already searched before, using cached data'

            $ADPropsMailboxes[$AccountNumberRunning] = $ADPropsMailboxes[$MailAddresses.IndexOf($MailAddresses[$AccountNumberRunning])]
        }
    }


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
            Write-Host '    No matching mailbox found' -ForegroundColor Yellow
        }
    } else {
        Write-Host '  AD mail attribute of currently logged-in user is empty' -NoNewline

        if ($null -ne $TrustsToCheckForGroups[0]) {
            Write-Host ', searching msExchMasterAccountSid'
            # No mail attribute set, check for match(es) of user's objectSID and mailbox's msExchMasterAccountSid
            for ($i = 0; $i -lt $MailAddresses.count; $i++) {
                if ($ADPropsMailboxes[$i].msexchmasteraccountsid) {
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
                    if (($RegistryPaths[$i] -ilike "Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\*") -and ($i -ne $p)) {
                        $MailboxNewOrder += $i
                    }
                }
            }

        }
    } else {
        for ($i = 0; $i -lt $MailAddresses.Count; $i++) {
            if ($MailboxNewOrder -inotcontains $i ) {
                $MailboxNewOrder += $i
            }
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

    $TemplateFilesGroupSIDsOverall = @{}

    foreach ($SigOrOOF in ('signature', 'OOF')) {
        if (($SigOrOOF -eq 'OOF') -and ($SetCurrentUserOOFMessage -eq $false)) {
            break
        }

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
        $TemplateIniPath = Get-Variable -Name "$($SigOrOOF)IniPath" -ValueOnly
        $TemplateIniSettings = Get-Variable -Name "$($SigOrOOF)IniSettings" -ValueOnly

        $TemplateFiles = @((Get-ChildItem -LiteralPath $TemplateTemplatePath -File -Filter $(if ($UseHtmTemplates) { '*.htm' } else { '*.docx' })) | Sort-Object -Culture $TemplateFilesSortCulture)

        if ($TemplateIniPath -ne '') {
            Write-Host "  Compare $($SigOrOOF) ini entries and file system"
            foreach ($Enumerator in $TemplateIniSettings.GetEnumerator().name) {
                if ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>']) {
                    if (($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -ine '<Set-OutlookSignatures configuration>') -and ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -inotin $TemplateFiles.name)) {
                        Write-Host "    '$($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'])' ($($SigOrOOF) ini index #$($Enumerator)) found in ini but not in signature template path." -ForegroundColor Yellow
                    }

                    if (($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -ine '<Set-OutlookSignatures configuration>') -and ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -inotlike "*.$(if($UseHtmTemplates){'htm'}else{'docx'})")) {
                        Write-Host "    '$($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'])' ($($SigOrOOF) ini index #$($Enumerator)) has the wrong file extension (-UseHtmTemplates true allows .htm, else .docx)" -ForegroundColor Yellow
                    }
                }
            }

            $x = @(foreach ($Enumerator in $TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)]) { $Enumerator['<Set-OutlookSignatures template>'] })
            foreach ($TemplateFile in $TemplateFiles) {
                if ($TemplateFile.name -inotin $x) {
                    Write-Host "    '$($TemplateFile.name)' found in $($SigOrOOF) template path but not in ini." -ForegroundColor Yellow
                }
            }

            Write-Host '  Sort template files according to configuration'
            try {
                $TemplateFilesSortCulture = (@($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1)['SortCulture']
            } catch {
                $TemplateFilesSortCulture = $null
            }

            # Populate template files in the most complicated way first: SortOrder 'AsInThisFile'
            # This also considers that templates can be referenced multiple times in the INI file
            # If the setting in the ini file is different, we only need to sort $TemplateFiles
            $TemplateFilesExisting = @(foreach ($Enumerator in $TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)]) { $Enumerator['<Set-OutlookSignatures template>'] })
            $TemplateFiles = @($TemplateFiles | Where-Object { $_.name -iin $TemplateFilesExisting })
            $TemplateFiles | Add-Member -MemberType NoteProperty -Name TemplateIniSettingsIndex -Value $null
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
                            # same as 'ascending'
                            $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } })
                        }
                    }
                } else {
                    $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } })
                }
            }
        }

        foreach ($TemplateFile in $TemplateFiles) {
            $TemplateIniSettingsIndex = $TemplateFile.TemplateIniSettingsIndex
            $TemplateFileGroupSIDs = @{}
            Write-Host ("    '$($TemplateFile.Name)' ($($SigOrOOF) ini index #$($TemplateIniSettingsIndex))")

            if ($TemplateIniSettings[$TemplateIniSettingsIndex]['<Set-OutlookSignatures template>'] -ieq $TemplateFile.name) {
                $TemplateFilePart = ($TemplateIniSettings[$TemplateIniSettingsIndex].GetEnumerator().Name -join '] [')
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

            # time based template
            $TemplateFileTimeActive = $true
            if (($TemplateFilePart -imatch $TemplateFilePartRegexTimeAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexTimeDeny)) {
                Write-Host '      Time based template'
                if (-not $BenefactorCircleLicenseFile) {
                    Write-Host "        The 'time based template' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                    Write-Host "        Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
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

                            Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '') = " -NoNewline
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
                                    Write-Host "$($TemplateFileGroupSIDs[$TemplateFilePartTag] -ireplace '(?i)^(-:|-CURRENTUSER:|CURRENTUSER:|)', '')"
                                    $TemplateFilesGroupFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesGroupFilePart[$TemplateIniSettingsIndex] + '[' + $TemplateFileGroupSIDs[$TemplateFilePartTag] + ']')
                                } else {
                                    Write-Host 'Not found' -ForegroundColor Yellow
                                }
                            } else {
                                Write-Host 'Not found' -ForegroundColor Yellow
                                $TemplateFilesGroupSIDsOverall.add($($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '$1$3'), $null)
                            }
                        }
                    }
                }

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
            $tempVerbosePreference = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'
            $script:COMWordDummy = New-Object -ComObject Word.Application
            $VerbosePreference = $tempVerbosePreference
            $script:COMWordDummy.Visible = $false

            if ($script:COMWordDummy) {
                # Set Word process priority
                $script:COMWordDummyCaption = $script:COMWordDummy.Caption
                $script:COMWordDummy.Caption = "Set-OutlookSignatures $([guid]::NewGuid())"
                $script:COMWordDummyHWND = [Win32Api]::FindWindow( 'OpusApp', $($script:COMWordDummy.Caption) )
                $script:COMWordDummyPid = [IntPtr]::Zero
                $null = [Win32Api]::GetWindowThreadProcessId( $script:COMWordDummyHWND, [ref] $script:COMWordDummyPid );
                $script:COMWordDummy.Caption = $script:COMWordDummyCaption
                $null = Get-CimInstance Win32_process -Filter "ProcessId = ""$script:COMWordDummyPid""" | Invoke-CimMethod -Name SetPriority -Arguments @{Priority = $WordProcessPriority }
            }

            $tempVerbosePreference = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'
            $script:COMWord = New-Object -ComObject Word.Application
            $VerbosePreference = $tempVerbosePreference
            $script:COMWord.Visible = $false

            if ($script:COMWord) {
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

            Add-Type -Path (Get-ChildItem -LiteralPath ((Join-Path -Path ($env:SystemRoot) -ChildPath 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word')) -Filter 'Microsoft.Office.Interop.Word.dll' -Recurse | Select-Object -ExpandProperty FullName -Last 1)
        } catch {
            Write-Host '  Word not installed or not working correctly. Exit.' -ForegroundColor Red
            $error[0]
            exit 1
        }
    }


    # Process each email address only once
    $script:SignatureFilesDone = @()
    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        if (($AccountNumberRunning -eq $MailAddresses.IndexOf($MailAddresses[$AccountNumberRunning])) -and ($($MailAddresses[$AccountNumberRunning]) -like '*@*')) {
            Write-Host
            Write-Host "Mailbox $($MailAddresses[$AccountNumberRunning]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

            $UserDomain = ''

            Write-Host "  Get group membership of mailbox @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            if ($($ADPropsMailboxesUserDomain[$AccountNumberRunning])) {
                Write-Host "    $($ADPropsMailboxesUserDomain[$AccountNumberRunning]) (mailbox home domain/forest) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            }

            $GroupsSIDs = @()
            $ADPropsCurrentMailbox = @()
            $ADPropsCurrentMailboxManager = @()

            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                $ADPropsCurrentMailbox = $ADPropsMailboxes[$AccountNumberRunning]

                if ($null -ne $TrustsToCheckForGroups[0]) {
                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")

                    try {
                        $Search.filter = "(distinguishedname=$($ADPropsCurrentMailbox.manager))"
                        $ADPropsCurrentMailboxManager = ([ADSI]"$(($Search.FindOne()).path)").Properties

                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")
                        $Search.filter = "(distinguishedname=$($ADPropsCurrentMailbox.manager))"
                        $ADPropsCurrentMailboxManagerLdap = ([ADSI]"$(($Search.FindOne()).path)").Properties

                        foreach ($keyName in @($ADPropsCurrentMailboxManagerLdap.Keys)) {
                            if (
                                ($keyName -inotin $ADPropsCurrentMailboxManager.Keys) -or
                                (-not ($ADPropsCurrentMailboxManager[$keyName]) -and ($ADPropsCurrentMailboxManagerLdap[$keyName]))
                            ) {
                                $ADPropsCurrentMailboxManager[$keyName] = $ADPropsCurrentMailboxManagerLdap[$keyName]
                            }
                        }
                    } catch {
                        $ADPropsCurrentMailboxManager = @()
                    }

                    $UserDomain = $ADPropsMailboxesUserDomain[$AccountNumberRunning]
                    $SIDsToCheckInTrusts = @()

                    if ($ADPropsCurrentMailbox.objectsid) {
                        $SIDsToCheckInTrusts += (New-Object System.Security.Principal.SecurityIdentifier $($ADPropsCurrentMailbox.objectsid), 0).value
                    }

                    foreach ($SidHistorySid in @($ADPropsCurrentMailbox.sidhistory | Where-Object { $_ })) {
                        $SIDsToCheckInTrusts += (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                    }

                    try {
                        # Security groups and distribution groups, global and universal, forest-wide
                        Write-Verbose "      LDAP query of tokengroupsglobalanduniversal attribute (security and distribution groups, global and universal, forest-wide) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                        $UserAccount = [ADSI]"LDAP://$($ADPropsCurrentMailbox.distinguishedname)"
                        $UserAccount.GetInfoEx(@('tokengroupsglobalanduniversal'), 0)
                        foreach ($sidBytes in $UserAccount.Properties.tokengroupsglobalanduniversal) {
                            $sid = (New-Object System.Security.Principal.SecurityIdentifier($sidbytes, 0)).value
                            Write-Verbose "        $sid"
                            $GroupsSIDs += $sid
                            $SIDsToCheckInTrusts += $sid
                        }

                        <#
                        # No longer needed due to switching to tokengroupsglobalanduniversal from tokengroups, and querying local groups only via parameter command
                        # Distribution groups (static only)
                        Write-Verbose "      GC query for static distribution groups (global and universal, forest-wide) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$(($($ADPropsCurrentMailbox.distinguishedname) -split ',DC=')[1..999] -join '.')")
                        $Search.filter = "(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648))(member:1.2.840.113556.1.4.1941:=$($ADPropsCurrentMailbox.distinguishedname)))"
                        foreach ($DistributionGroup in $search.findall()) {
                            if ($DistributionGroup.properties.objectsid) {
                                $sid = (New-Object System.Security.Principal.SecurityIdentifier $($DistributionGroup.properties.objectsid), 0).value
                                Write-Verbose "        $sid"
                                $GroupsSIDs += $sid
                                $SIDsToCheckInTrusts += $sid
                            }

                            foreach ($SidHistorySid in @($DistributionGroup.properties.sidhistory | Where-Object { $_ })) {
                                $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                                Write-Verbose "        $sid"
                                $GroupsSIDs += $sid
                                $SIDsToCheckInTrusts += $sid
                            }
                        }
                        #>

                        # Domain local groups
                        if ($IncludeMailboxForestDomainLocalGroups -eq $true) {
                            Write-Verbose "      LDAP query for domain local groups (security and distribution, one query per domain) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            foreach ($DomainToCheckForDomainLocalGroups in @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ieq $LookupDomainsToTrusts[$(($($ADPropsCurrentMailbox.distinguishedname) -split ',DC=')[1..999] -join '.')] }).name)) {
                                Write-Verbose "        $($DomainToCheckForDomainLocalGroups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DomainToCheckForDomainLocalGroups)")
                                $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($ADPropsCurrentMailbox.distinguishedname)))"
                                foreach ($LocalGroup in $search.findall()) {
                                    if ($LocalGroup.properties.objectsid) {
                                        $sid = (New-Object System.Security.Principal.SecurityIdentifier $($LocalGroup.properties.objectsid), 0).value
                                        Write-Verbose "          $sid"
                                        $GroupsSIDs += $sid
                                        $SIDsToCheckInTrusts += $sid
                                    }

                                    foreach ($SidHistorySid in @($LocalGroup.properties.sidhistory | Where-Object { $_ })) {
                                        $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                                        Write-Verbose "          $sid"
                                        $GroupsSIDs += $sid
                                        $SIDsToCheckInTrusts += $sid
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Host "      Error getting group information from $((($ADPropsCurrentMailbox.distinguishedname) -split ',DC=')[1..999] -join '.'), check firewalls, DNS and AD trust" -ForegroundColor Red
                        $error[0]
                    }

                    $GroupsSIDs = @($GroupsSIDs | Select-Object -Unique)

                    # Loop through all domains outside the mailbox account's home forest to check if the mailbox account has a group membership there
                    # Across a trust, a user can only be added to a domain local group.
                    # Domain local groups can not be used outside their own domain, so we don't need to query recursively
                    # But when it's a cross-forest trust, we need to query every every domain on that other side of the trust
                    #   This is handled before by adding every single domain of a cross-forest trusted forest to $TrustsToCheckForGroups
                    if ($SIDsToCheckInTrusts.count -gt 0) {
                        $SIDsToCheckInTrusts = @($SIDsToCheckInTrusts | Select-Object -Unique)
                        $LdapFilterSIDs = '(|'

                        foreach ($SidToCheckInTrusts in $SIDsToCheckInTrusts) {
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
                                Write-Host '      Error creating LDAP filter for search across trusts.' -ForegroundColor Red
                                $error[0]
                            }
                        }
                        $LdapFilterSIDs += ')'
                    } else {
                        $LdapFilterSIDs = ''
                    }

                    if ($LdapFilterSids -ilike '*(objectsid=*') {
                        # Across each trust, search for all Foreign Security Principals matching a SID from our list
                        foreach ($TrustToCheckForFSPs in @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $LookupDomainsToTrusts[$(($($ADPropsCurrentMailbox.distinguishedname) -split ',DC=')[1..999] -join '.')] }).value | Select-Object -Unique)) {
                            Write-Host "    $($TrustToCheckForFSPs) (trusted domain/forest of mailbox home forest) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustToCheckForFSPs)")
                            $Search.filter = "(&(objectclass=foreignsecurityprincipal)$LdapFilterSIDs)"

                            $fsps = $Search.FindAll()

                            if ($fsps.count -gt 0) {
                                foreach ($fsp in $fsps) {
                                    if (($fsp.path -ne '') -and ($null -ne $fsp.path)) {
                                        # A Foreign Security Principal (FSP) is created in each (sub)domain in which it is granted permissions
                                        # A FSP it can only be member of a domain local group - so we set the searchroot to the (sub)domain of the Foreign Security Principal
                                        # FSPs have no tokengroups or tokengroupsglobalanduniversal attribute, which would not contain domain local groups anyhow
                                        # member:1.2.840.113556.1.4.1941:= (LDAP_MATCHING_RULE_IN_CHAIN) returns groups containing a specific DN as member, incl. nesting
                                        Write-Verbose "      Found ForeignSecurityPrincipal $($fsp.properties.cn) in $((($fsp.path -split ',DC=')[1..999] -join '.'))"

                                        if ($((($fsp.path -split ',DC=')[1..999] -join '.')) -iin @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $LookupDomainsToTrusts[$(($($ADPropsCurrentMailbox.distinguishedname) -split ',DC=')[1..999] -join '.')] }).name)) {
                                            try {
                                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$((($fsp.path -split ',DC=')[1..999] -join '.'))")
                                                $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($fsp.Properties.distinguishedname)))"

                                                $fspGroups = $Search.FindAll()

                                                if ($fspGroups.count -gt 0) {
                                                    foreach ($group in $fspgroups) {
                                                        $sid = (New-Object System.Security.Principal.SecurityIdentifier $($group.properties.objectsid), 0).value
                                                        Write-Verbose "        $sid"
                                                        $GroupsSIDs += $sid

                                                        foreach ($SidHistorySid in @($group.properties.sidhistory | Where-Object { $_ })) {
                                                            $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                                                            Write-Verbose "        $sid"
                                                            $GroupsSIDs += $sid
                                                        }
                                                    }
                                                } else {
                                                    Write-Verbose '        FSP is not member of any group'
                                                }
                                            } catch {
                                                Write-Host "        Error: $($error[0].exception)" -ForegroundColor red
                                            }
                                        } else {
                                            Write-Verbose "        Ignoring, because '$($fsp.path)' is not part of a trust in TrustsToCheckForGroups."
                                        }
                                    }
                                }
                            } else {
                                Write-Verbose '      No ForeignSecurityPrincipal(s) found'
                            }
                        }
                    }
                } else {
                    try {
                        $AADProps = $null
                        if ($ADPropsCurrentMailbox.manager) {
                            $AADProps = (GraphGetUserProperties $ADPropsCurrentMailbox.manager).properties

                            $ADPropsCurrentMailboxManager = [PSCustomObject]@{}

                            foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                                $z = $AADProps

                                foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                                    $z = $z.$y
                                }

                                $ADPropsCurrentMailboxManager | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z
                            }

                            $ADPropsCurrentMailboxManager | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsCurrentMailboxManager.userprincipalname).photo
                            $ADPropsCurrentMailboxManager | Add-Member -MemberType NoteProperty -Name 'manager' -Value $null
                        }
                        Write-Host '    Microsoft Graph'
                        foreach ($sid in @((GraphGetUserTransitiveMemberOf $ADPropsCurrentMailbox.userPrincipalName).memberof.value.securityidentifier | Where-Object { $_ })) {
                            $GroupsSIDs += $sid
                            Write-Verbose "      $sid"
                        }
                    } catch {
                        $ADPropsCurrentMailboxManager = @()
                        Write-Host '    Skipping, mailbox not in Microsoft Graph.' -ForegroundColor yellow
                    }
                }
            } else {
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox' -ForegroundColor yellow
            }

            $ADPropsCurrentMailbox | Add-Member -MemberType NoteProperty -Name 'GroupsSIDs' -Value $GroupsSIDs

            if ($MailAddresses[$AccountNumberRunning] -ieq $PrimaryMailboxAddress) {
                $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'GroupsSIDs' -Value $GroupsSIDs
            }

            Write-Host "  Get SMTP addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
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
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox' -ForegroundColor Yellow
                Write-Host '    Use mailbox name as single known SMTP address' -ForegroundColor Yellow
            }

            Write-Host "  Get data for replacement variables @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            $ReplaceHash = @{}

            if (Test-Path -Path $ReplacementVariableConfigFile -PathType Leaf) {
                try {
                    Write-Host "    Execute config file '$ReplacementVariableConfigFile'"
                    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $ReplacementVariableConfigFile -Encoding UTF8 -Raw)))
                } catch {
                    Write-Host "    Problem executing content of '$ReplacementVariableConfigFile'. Exit." -ForegroundColor Red
                    $error[0]
                    exit 1
                }
            } else {
                Write-Host "    Problem connecting to or reading from file '$ReplacementVariableConfigFile'. Exit." -ForegroundColor Red
                exit 1
            }

            foreach ($replaceKey in @($replaceHash.Keys | Sort-Object -Culture $TemplateFilesSortCulture)) {
                if ($replaceKey -notin ('$CurrentMailboxManagerPhoto$', '$CurrentMailboxPhoto$', '$CurrentUserManagerPhoto$', '$CurrentUserPhoto$', '$CurrentMailboxManagerPhotodeleteempty$', '$CurrentMailboxPhotodeleteempty$', '$CurrentUserManagerPhotodeleteempty$', '$CurrentUserPhotodeleteempty$')) {
                    if ($($replaceHash[$replaceKey])) {
                        Write-Verbose "    $($replaceKey): $($replaceHash[$replaceKey])"
                    }
                } else {
                    if ($null -ne $($replaceHash[$replaceKey])) {
                        Write-Verbose "    $($replaceKey): Photo available"
                    }
                }
            }

            # Export pictures if available
            $CurrentMailboxManagerPhotoGuid = (New-Guid).guid
            $CurrentMailboxPhotoGuid = (New-Guid).guid
            $CurrentUserManagerPhotoGuid = (New-Guid).guid
            $CurrentUserPhotoGuid = (New-Guid).guid

            foreach ($VariableName in (('$CurrentMailboxManagerPhoto$', $CurrentMailboxManagerPhotoGuid) , ('$CurrentMailboxPhoto$', $CurrentMailboxPhotoguid), ('$CurrentUserManagerPhoto$', $CurrentUserManagerPhotoGuid), ('$CurrentUserPhoto$', $CurrentUserPhotoGuid))) {
                if ($null -ne $ReplaceHash[$VariableName[0]]) {
                    if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                        $ReplaceHash[$VariableName[0]] | Set-Content -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))) -AsByteStream -Force
                    } else {
                        $ReplaceHash[$VariableName[0]] | Set-Content -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))) -Encoding Byte -Force
                    }
                }
            }

            if ($MirrorLocalSignaturesToCloud -eq $true) {
                if (-not $BenefactorCircleLicenseFile) {
                    Write-Host "    The 'MirrorLocalSignaturesToCloud' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                    Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesDownload()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '    Error downloading roaming signatures from the cloud.' -ForegroundColor Yellow
                        Write-Host "    $FeatureResult" -ForegroundColor Yellow
                    }
                }
            }


            EvaluateAndSetSignatures


            # Delete photos from file system
            foreach ($VariableName in (('$CurrentMailboxManagerPhoto$', $CurrentMailboxManagerPhotoGuid) , ('$CurrentMailboxPhoto$', $CurrentMailboxPhotoguid), ('$CurrentUserManagerPhoto$', $CurrentUserManagerPhotoGuid), ('$CurrentUserPhoto$', $CurrentUserPhotoGuid))) {
                Remove-Item -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))) -Force -ErrorAction SilentlyContinue
                $ReplaceHash.Remove($VariableName[0])
                $ReplaceHash.Remove(($VariableName[0][-999..-2] -join '') + 'DELETEEMPTY$')
            }


            # Set OOF message and Outlook Web signature
            if (((($SetCurrentUserOutlookWebSignature -eq $true)) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses[$AccountNumberRunning] -ieq $PrimaryMailboxAddress)) {
                if ((-not $SimulateUser)) {
                    if (-not $script:WebServicesDllPath) {
                        Write-Host "  Set up environment for connection to Outlook Web @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                        $script:WebServicesDllPath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))
                        try {
                            Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\netstandard2.0\Microsoft.Exchange.WebServices.Data.dll')) -Destination $script:WebServicesDllPath -Force
                            Unblock-File -LiteralPath $script:WebServicesDllPath

                            #if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                            #    Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\netcore\Microsoft.Exchange.WebServices.Data.dll')) -Destination $script:WebServicesDllPath -Force
                            #    Unblock-File -LiteralPath $script:WebServicesDllPath
                            #} else {
                            #    Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\netframework\Microsoft.Exchange.WebServices.dll')) -Destination $script:WebServicesDllPath -Force
                            #    Unblock-File -LiteralPath $script:WebServicesDllPath
                            #}
                        } catch {
                        }
                    }

                    $error.clear()

                    try {
                        Import-Module -Name $script:WebServicesDllPath -Force -ErrorAction Stop

                        $exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

                        Write-Host '  Connect to Outlook Web'
                        try {
                            Write-Verbose '    Try Integrated Windows Authentication'

                            if (
                                ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile -and $GraphToken.AppAccessTokenExo) -or
                                $GraphToken.AccessTokenExo
                            ) {
                                throw '      EXO access token available, skip Integrated Windows Authentication'
                            }

                            $exchService.UseDefaultCredentials = $true
                            $exchService.ImpersonatedUserId = $null
                            $exchService.AutodiscoverUrl($PrimaryMailboxAddress, { $true }) | Out-Null
                        } catch {
                            try {
                                Write-Verbose $_

                                Write-Verbose '    Try OAuth with Autodiscover'

                                $exchService.UseDefaultCredentials = $false

                                if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) {
                                    $exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $PrimaryMailboxAddress)
                                    $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AppAccessTokenExo)
                                } else {
                                    $exchService.ImpersonatedUserId = $null
                                    $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AccessTokenExo)
                                }

                                $exchService.AutodiscoverUrl($PrimaryMailboxAddress, { $true }) | Out-Null
                            } catch {
                                Write-Verbose $_

                                Write-Verbose '    Try OAuth with fixed URL'

                                $exchService.UseDefaultCredentials = $false

                                if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) {
                                    $exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $PrimaryMailboxAddress)
                                    $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AppAccessTokenExo)
                                } else {
                                    $exchService.ImpersonatedUserId = $null
                                    $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($GraphToken.AccessTokenExo)
                                }

                                $exchService.Url = "$($CloudEnvironmentExchangeOnlineEndpoint)/EWS/Exchange.asmx"
                            }
                        }

                        $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchservice, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)

                        if ($Calendar.DisplayName) {
                            $error.clear()
                        } else {
                            Write-Host '    Could not connect to Outlook Web, although the EWS DLL threw no error.' -ForegroundColor Red
                            throw
                        }
                    } catch {
                        Write-Host "    Error connecting to Outlook Web: $_" -ForegroundColor Red

                        if ($SetCurrentUserOutlookWebSignature) {
                            Write-Host '    Outlook Web signature can not be set' -ForegroundColor Red
                            $SetCurrentUserOutlookWebSignature = $false
                        }

                        if ($SetCurrentUserOOFMessage -and (($null -ne $TrustsToCheckForGroups[0]) -and ($ADPropsCurrentMailbox.msexchrecipienttypedetails -lt 2147483648))) {
                            Write-Host '   out of office (OOF) message(s) can not be set' -ForegroundColor Red
                            $SetCurrentUserOOFMessage = $false
                        }
                    }
                } else {
                    $error.Clear()
                }

                if ($SetCurrentUserOutlookWebSignature) {
                    Write-Host "  Set classic Outlook Web signature @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                    if ($SimulateUser -and (-not $SimulateAndDeploy)) {
                        Write-Host '    Simulation mode enabled, skip task' -ForegroundColor Yellow
                    } else {
                        if (-not $BenefactorCircleLicenseFile) {
                            Write-Host "    The 'SetCurrentUserOutlookWebSignature' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                            Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                        } else {
                            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SetCurrentUserOutlookWebSignature()

                            if ($FeatureResult -ne 'true') {
                                Write-Host '    Error setting current user Outlook web signature.' -ForegroundColor Yellow
                                Write-Host "    $FeatureResult" -ForegroundColor Yellow
                            }
                        }

                        if ($MirrorLocalSignaturesToCloud -eq $true) {
                            Write-Host "  Set roaming Outlook Web signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            if (-not $BenefactorCircleLicenseFile) {
                                Write-Host "    The 'MirrorLocalSignaturesToCloud' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                                Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                            } else {
                                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesSetDefaults()

                                if ($FeatureResult -ne 'true') {
                                    Write-Host '    Error setting default roaming signatures in the cloud.' -ForegroundColor Yellow
                                    Write-Host "    $FeatureResult" -ForegroundColor Yellow
                                }
                            }
                        }
                    }
                }

                if ($SetCurrentUserOOFMessage) {
                    Write-Host "  Process out of office (OOF) auto replies @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                    if (-not $BenefactorCircleLicenseFile) {
                        Write-Host "    The 'SetCurrentUserOOFMessage' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                        Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                    } else {
                        $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SetCurrentUserOOFMessage()

                        if ($FeatureResult -ne 'true') {
                            Write-Host '    Error setting current user out of office message.' -ForegroundColor Yellow
                            Write-Host "    $FeatureResult" -ForegroundColor Yellow
                        }
                    }
                }
            }
        }
    }


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


    # Delete old signatures created by this script, which are no longer available in $SignatureTemplatePath
    # We check all local signatures for a specific marker in HTML code, so we don't touch user-created signatures
    if ($DeleteScriptCreatedSignaturesWithoutTemplate -eq $true) {
        Write-Host
        Write-Host "Remove old signatures created by this script, which are no longer centrally available @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        if (-not $BenefactorCircleLicenseFile) {
            Write-Host "  The 'DeleteScriptCreatedSignaturesWithoutTemplate' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
            Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
        } else {
            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DeleteScriptCreatedSignaturesWithoutTemplate()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error deleting script created signature which no longer have a corresponding template.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    }

    # Delete user-created signatures if $DeleteUserCreatedSignatures -eq $true
    if ($DeleteUserCreatedSignatures -eq $true) {
        Write-Host
        Write-Host "Remove user-created signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        if (-not $BenefactorCircleLicenseFile) {
            Write-Host "  The 'DeleteUserCreatedSignatures' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
            Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
        } else {
            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DeleteUserCreatedSignatures()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error removing user-created signatures.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    }

    # Upload local signatures to Exchange Online as roaming signatures
    if ($MirrorLocalSignaturesToCloudDoUpload -eq $true) {
        Write-Host
        Write-Host "Uploading local signatures to Exchange Online as roaming signatures for current user @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        if (-not $BenefactorCircleLicenseFile) {
            Write-Host "  The 'MirrorLocalSignaturesToCloud' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
            Write-Host "  Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
        } else {
            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesUpload()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error uploading roaming signatures to the cloud.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    }

    # Copy signatures to additional path if $AdditionalSignaturePath is set
    if ($AdditionalSignaturePath) {
        Write-Host
        Write-Host "Copy signatures to AdditionalSignaturePath @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        Write-Host "  '$AdditionalSignaturePath'"

        if ($SimulateUser) {
            Write-Host '    Simulation mode enabled, AdditionalSignaturePath already used as output directory' -ForegroundColor Yellow
        } else {
            if (-not $BenefactorCircleLicenseFile) {
                Write-Host "    The 'AdditionalSignaturePath' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                Write-Host "    Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
            } else {
                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::AdditionalSignaturePath()

                if ($FeatureResult -ne 'true') {
                    Write-Host '    Error copying signatures to additional signature path.' -ForegroundColor Yellow
                    Write-Host "    $FeatureResult" -ForegroundColor Yellow
                }
            }
        }
    }

    if (
        ($script:CurrentUserDummyMailbox -eq $true) -or
        ($OutlookUseNewOutlook -eq $true)
    ) {
        RemoveItemAlternativeRecurse $SignaturePaths[0] -SkipFolder
    }
}


function ResolveToSid($string) {
    # Find the last ':', use everything right from it and remove surrounding whitespace
    $string = (($string -split ':')[-1]).trim()

    if ($string.contains('\')) {
        # is already a NT4 style format
        $local:NTName = $string
    } elseif ($string.contains(' ')) {
        # make it a NT4 style format
        $local:NTName = ([regex]' ').replace($string, '\', 1)
    } else {
        # Invalid
        return $null
    }

    if (($null -ne $TrustsToCheckForGroups[0]) -and (-not ($local:NTName -imatch '^(AzureAD\\|EntraID\\)'))) {
        try {
            $local:x = (New-Object System.Security.Principal.NTAccount($local:NTName)).Translate([System.Security.Principal.SecurityIdentifier]).value

            if ($local:x) {
                return $local:x
            }
        } catch {
            try {
                # No group with this sAMAccountName found. Interpreting it as a display name.

                $objTrans = New-Object -ComObject 'NameTranslate'
                $objNT = $objTrans.GetType()
                $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($local:NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (4, ($local:NTName -split '\\')[1])) # 4 = ADS_NAME_TYPE_DISPLAY

                $local:x = $(((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
                Remove-Variable -Name 'objTrans'
                Remove-Variable -Name 'objNT'

                if ($local:x) {
                    return $local:x
                }
            } catch {
                try {
                    # Let the API guess what it is

                    $objTrans = New-Object -ComObject 'NameTranslate'
                    $objNT = $objTrans.GetType()
                    $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($local:NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                    $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, ($local:NTName -split '\\')[1])) # 8 = ADS_NAME_TYPE_UNKNOWN

                    $local:x = $(((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)

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
        $tempFilterOrder = @(
            "((onPremisesNetBiosName eq '$($local:NTName.Split('\')[0])') and (onPremisesSamAccountName eq '$($local:NTName.Split('\')[1])'))"
            "((onPremisesNetBiosName eq '$($local:NTName.Split('\')[0])') and (displayName eq '$($local:NTName.Split('\')[1])'))"
            "(proxyAddresses/any(x:x eq 'smtp:$($local:NTName.Split('\')[1])'))"
            "(mailNickname eq '$($local:NTName.Split('\')[1])')"
            "(displayName eq '$($local:NTName.Split('\')[1])')"
        )

        # Search Graph for groups
        ForEach ($tempFilter in $tempFilterOrder) {
            $tempResults = (GraphFilterGroups $tempFilter)
            if (($tempResults.error -eq $false) -and ($tempResults.groups.count -eq 1) -and $($tempResults.groups[0].value)) {
                if ($($tempResults.groups[0].value.securityidentifier)) {
                    return $($tempResults.groups[0].value.securityidentifier)
                }
            }
        }

        # Search Graph for users
        ForEach ($tempFilter in $tempFilterOrder) {
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
                $runtimeAssembly = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($file)
            } Catch {
                $runtimeAssembly = $null
            }

            Try {
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


Function ConvertToSingleFileHTML([string]$inputfile, [string]$outputfile) {
    $tempFileContent = Get-Content -LiteralPath $inputfile -Encoding UTF8 -Raw

    $src = @()

    foreach ($regex in @(([regex]'(?i)src="(.*?)"').Matches($tempFileContent))) {
        $src += $regex.Groups[0].Value

        if ($regex.Groups[0].Value.StartsWith('src="data:')) {
            $src += ''
        } else {
            $src += (Join-Path -Path (Split-Path -Path ($inputfile) -Parent) -ChildPath ([uri]::UnEscapeDataString($regex.Groups[1].Value)))
        }
    }

    for ($x = 0; $x -lt $src.count; $x = $x + 2) {
        if ($src[$x].StartsWith('src="data:')) {
            # nothing to do
        } elseif (Test-Path -LiteralPath $src[$x + 1] -PathType leaf) {
            $fmt = $null

            switch ((Get-ChildItem -LiteralPath $src[$x + 1]).Extension) {
                '.apng' { $fmt = 'data:image/apng;base64,'; break }
                '.avif' { $fmt = 'data:image/avif;base64,'; break }
                '.gif' { $fmt = 'data:image/gif;base64,'; break }
                '.jpg' { $fmt = 'data:image/jpeg;base64,'; break }
                '.jpeg' { $fmt = 'data:image/jpeg;base64,'; break }
                '.jfif' { $fmt = 'data:image/jpeg;base64,'; break }
                '.pjpeg' { $fmt = 'data:image/jpeg;base64,'; break }
                '.pjp' { $fmt = 'data:image/jpeg;base64,'; break }
                '.png' { $fmt = 'data:image/png;base64,'; break }
                '.png' { $fmt = 'data:image/svg+xml;base64,'; break }
                '.webp' { $fmt = 'data:image/webp;base64,'; break }
                '.css' { $fmt = 'data:text/css;base64,'; break }
                '.less' { $fmt = 'data:text/css;base64,'; break }
                '.js' { $fmt = 'data:text/javascript;base64,'; break }
                '.otf' { $fmt = 'data:font/otf;base64,'; break }
                '.sfnt' { $fmt = 'data:font/sfnt;base64,'; break }
                '.ttf' { $fmt = 'data:font/ttf;base64,'; break }
                '.woff' { $fmt = 'data:font/woff;base64,'; break }
                '.woff2' { $fmt = 'data:font/woff2;base64,'; break }
                default { $fmt = 'data:;base64,' }
            }

            if ($fmt) {
                $tempFileContent = $tempFileContent -ireplace [Regex]::Escape($src[$x]), ('src="' + $fmt + [Convert]::ToBase64String(([System.IO.File]::ReadAllBytes(((Resolve-Path $src[$x + 1]).Path)))) + '"')
            }
        }
    }

    if ((Test-Path -LiteralPath $outputfile) -and (Get-ChildItem -LiteralPath $outputfile).IsReadOnly) {
        Remove-Item -LiteralPath $outputfile -Force -Recurse -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        [System.IO.File]::WriteAllLines($outputfile, $tempFileContent, (New-Object System.Text.UTF8Encoding($False)))
        Set-ItemProperty -LiteralPath $outputfile -Name IsReadOnly -Value $true
    } else {
        Remove-Item -LiteralPath $outputfile -Force -Recurse -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        [System.IO.File]::WriteAllLines($outputfile, $tempFileContent, (New-Object System.Text.UTF8Encoding($False)))
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
        Write-Host "$Indent  Process $TemplateGroup $(if($TemplateGroup -iin ('group', 'mailbox', 'replacementvariable')){'specific '})templates @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        if (-not (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly -ErrorAction SilentlyContinue)) {
            continue
        }

        for ($TemplateFileIndex = 0; $TemplateFileIndex -lt (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly).count; $TemplateFileIndex++) {
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


            # check for allow entries
            Write-Host "$Indent        Allows"
            if ($TemplateGroup -ieq 'common') {
                $TemplateAllowed = $true
                Write-Host "$Indent          Common: Template is classified as common template valid for all mailboxes"
            } elseif ($TemplateGroup -ieq 'group') {
                $tempAllowCount = 0

                foreach ($GroupsSid in $GroupsSIDs) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($GroupsSid)``]*") {
                        $TemplateAllowed = $true
                        $tempAllowCount++
                        Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', '$1') -join '|') = $($GroupsSid) (current mailbox)"
                        break
                    }
                }

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
                $tempAllowCount = 0

                foreach ($CurrentMailboxSmtpaddress in $CurrentMailboxSmtpAddresses) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($CurrentMailboxSmtpAddress)``]*") {
                        $TemplateAllowed = $true
                        $tempAllowCount++
                        Write-Host "$Indent          First email address match: $($CurrentMailboxSmtpAddress) (current mailbox)"
                        break
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          No email address match for current mailbox, checking current user specific allows"

                    foreach ($CurrentUserSmtpaddress in $ADPropsCurrentUser.proxyaddresses) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[CURRENTUSER:$($CurrentUserSmtpAddress -ireplace '^smtp:', '')``]*") {
                            $TemplateAllowed = $true
                            $tempAllowCount++
                            Write-Host "$Indent          First email address match: $($CurrentUserSmtpAddress -ireplace '^smtp:', '') (current user)"
                            break
                        }
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          Email address: Mailbox and current user do not have any allowed email address"
                }
            } elseif ($TemplateGroup -ieq 'replacementvariable') {
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
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Group: Mailbox and current user are not member of any denied group"
                }

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
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Email address: Mailbox and current user do not have any denied email address"
                }

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

            # result
            if ($Template -and ($TemplateAllowed -eq $true)) {
                Write-Host "$Indent        Use template as there is at least one allow and no deny"
                if ($ProcessOOF) {
                    if ($OOFFilesInternal.contains($TemplateIniSettingsIndex)) {
                        $OOFInternal = $Template
                    }

                    if ($OOFFilesExternal.contains($TemplateIniSettingsIndex)) {
                        $OOFExternal = $Template
                    }
                } else {
                    $Signature = $Template
                    SetSignatures -ProcessOOF:$ProcessOOF
                }
            } else {
                Write-Host "$Indent        Do not use template as there is no allow or at least one deny"
            }
        }
    }

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

            SetSignatures -ProcessOOF:$ProcessOOF

            if ($OOFExternal -eq $OOFInternal) {
                Copy-Item -Path (Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm") -Destination (Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")
            }
        }
    }

    # External OOF message
    if ($OOFExternal -and ($OOFExternal -ne $OOFInternal)) {
        $Signature = $OOFExternal

        Write-Host "$Indent    External OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        if ($UseHtmTemplates) {
            $Signature.value = "$OOFExternalGUID OOFExternal.htm"
        } else {
            $Signature.value = "$OOFExternalGUID OOFExternal.docx"
        }

        SetSignatures -ProcessOOF:$ProcessOOF
    }
}


function SetSignatures {
    Param(
        [switch]$ProcessOOF = $false
    )

    if ($ProcessOOF) {
        $Indent = '  '
    }

    if (-not $ProcessOOF) {
        Write-Host "      Outlook signature name: '$([System.IO.Path]::ChangeExtension($($Signature.value), $null) -ireplace '\.$')'"
    }

    if (-not $ProcessOOF) {
        $SignatureFileAlreadyDone = ($script:SignatureFilesDone -contains $TemplateIniSettingsIndex)

        if ($SignatureFileAlreadyDone) {
            Write-Host "$Indent      Template already processed before with higher priority, no need to update signature"
        } else {
            $script:SignatureFilesDone += $TemplateIniSettingsIndex
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

            $pathConnectedFolderNames += "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$($ConnectedFilesFolderName)"
            $pathConnectedFolderNames += [uri]::EscapeDataString($pathConnectedFolderNames[-1])
        }

        if ($UseHtmTemplates) {
            try {
                Copy-Item -LiteralPath $Signature.name -Destination $path

                foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                    if (Test-Path (Join-Path -Path (Split-Path $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName")) {
                        Copy-Item (Join-Path -Path (Split-Path $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName") (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files") -Recurse -Force
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
                Copy-Item -LiteralPath $Signature.Name -Destination $path -Force
            } catch {
                Write-Host "$Indent        Error copying file. Skip template." -ForegroundColor Red
                continue
            }
        }

        $Signature.value = $([System.IO.Path]::ChangeExtension($($Signature.value), '.htm'))

        if (-not $ProcessOOF) {
            $script:SignatureFilesDone += $Signature.Value

            if ($OutlookDisableRoamingSignatures -eq 0) {
                $script:SignatureFilesDone += ($Signature.Value -ireplace '\.htm$', " ($($PrimaryMailboxAddress)).htm")
                #$script:SignatureFilesDone += ($Signature.Value -ireplace '\.htm$', " ($($MailAddresses[$AccountNumberRunning])).htm")
            }
        }

        if ($UseHtmTemplates) {
            Write-Host "$Indent      Replace picture variables"

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
                $tempImageIsDeleted = $false

                if (($image.src -ilike '*$*$*') -or ($image.alt -ilike '*$*$*')) {
                    # Mailbox photos
                    foreach ($VariableName in (('$CurrentMailboxManagerPhoto$', $CurrentMailboxManagerPhotoGuid) , ('$CurrentMailboxPhoto$', $CurrentMailboxPhotoguid), ('$CurrentUserManagerPhoto$', $CurrentUserManagerPhotoGuid), ('$CurrentUserPhoto$', $CurrentUserPhotoGuid))) {
                        $tempImageVariableString = $Variablename[0] -ireplace '\$$', 'DELETEEMPTY$'

                        if (($image.src -ilike "*$($VariableName[0])*") -or ($image.alt -ilike "*$($VariableName[0])*")) {
                            if ($ReplaceHash[$VariableName[0]]) {
                                if ($EmbedImagesInHtml -eq $false) {
                                    Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -ireplace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                    Copy-Item (Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')) (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$($VariableName[0]).jpeg") -Force
                                    $image.src = [System.Web.HttpUtility]::UrlDecode("$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$($VariableName[0]).jpeg")

                                    if ($image.alt) {
                                        $image.alt = $($image.alt) -ireplace [Regex]::Escape($VariableName[0]), ''
                                    }
                                } else {
                                    $image.src = ('data:image/jpeg;base64,' + [Convert]::ToBase64String([System.IO.File]::ReadAllBytes(((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg'))))))
                                }
                            } else {
                                $image.src = "$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -ireplace '^about:', '')))))"
                            }
                        } elseif (($image.src -ilike "*$($tempImageVariableString)*") -or ($image.alt -ilike "*$($tempImageVariableString)*")) {
                            if ($ReplaceHash[$VariableName[0]]) {
                                if ($EmbedImagesInHtml -eq $false) {
                                    Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -ireplace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                    Copy-Item (Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')) (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$($VariableName[0]).jpeg") -Force
                                    $image.src = [System.Web.HttpUtility]::UrlDecode("$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$($VariableName[0]).jpeg")

                                    if ($image.alt) {
                                        $image.alt = $($image.alt) -ireplace [Regex]::Escape($tempImageVariableString), ''
                                    }
                                } else {
                                    $image.src = ('data:image/jpeg;base64,' + [Convert]::ToBase64String([System.IO.File]::ReadAllBytes(((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg'))))))
                                }
                            } else {
                                Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -ireplace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                $image.removenode() | Out-Null
                                $tempImageIsDeleted = $true
                                break
                            }
                        }

                        if ((-not $tempImageIsDeleted) -and ($image.alt)) {
                            $image.alt = $($image.alt) -ireplace [Regex]::Escape($VariableName[0]), ''
                            $image.alt = $($image.alt) -ireplace [Regex]::Escape($tempImageVariableString), ''
                        }
                    }

                    if ($tempImageIsDeleted) {
                        continue
                    }
                }

                # Other images
                if (($image.src -ilike '*$*DELETEEMPTY$*') -or ($image.alt -ilike '*$*DELETEEMPTY$*')) {
                    foreach ($VariableName in @(@($ReplaceHash.Keys) | Where-Object { $_ -inotin @('$CurrentMailboxPhoto$', '$CurrentMailboxManagerPhoto$', '$CurrentUserPhoto$', '$CurrentUserManagerPhoto$') })) {
                        $tempImageVariableString = $Variablename -ireplace '\$$', 'DELETEEMPTY$'

                        if (($image.src -ilike "*$($tempImageVariableString)*") -or ($image.alt -ilike "*$($tempImageVariableString)*")) {
                            if ($ReplaceHash[$VariableName]) {
                                if ($image.alt) {
                                    $image.alt = $($image.alt) -ireplace [Regex]::Escape($tempImageVariableString), ''
                                }
                            } else {
                                Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -ireplace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                $image.removenode() | Out-Null
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

            Write-Host "$Indent      Replace non-picture variables"
            $tempFileContent = $html.documentelement.outerhtml

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($html) | Out-Null
            Remove-Variable -Name 'html'

            foreach ($replaceKey in @($replaceHash.Keys | Where-Object { $_ -inotin ('$CurrentMailboxManagerPhoto$', '$CurrentMailboxPhoto$', '$CurrentUserManagerPhoto$', '$CurrentUserPhoto$', '$CurrentMailboxManagerPhotodeleteempty$', '$CurrentMailboxPhotodeleteempty$', '$CurrentUserManagerPhotodeleteempty$', '$CurrentUserPhotodeleteempty$') } | Sort-Object -Culture $TemplateFilesSortCulture)) {
                $tempFileContent = $tempFileContent -ireplace [Regex]::Escape($replacekey), $replaceHash.$replaceKey
            }

            Write-Host "$Indent      Export to HTM format"
            $tempFileContent | Out-File -LiteralPath $path -Encoding UTF8 -Force
        } else {
            $script:COMWord.Documents.Open($path, $false) | Out-Null

            Write-Host "$Indent      Replace picture variables"
            if ($script:COMWord.ActiveDocument.Shapes.Count -gt 0) {
                Write-Host "$Indent        Warning: Template contains $($script:COMWord.ActiveDocument.Shapes.Count) images configured as non-inline shapes." -ForegroundColor Yellow
                Write-Host "$Indent        Outlook does not support all formatting options of these images (e.g., behind the text)." -ForegroundColor Yellow
            }

            try {
                foreach ($image in @($script:COMWord.ActiveDocument.Shapes + $script:COMWord.ActiveDocument.InlineShapes)) {
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
                        foreach ($Variablename in (('$CurrentMailboxManagerPhoto$', $CurrentMailboxManagerPhotoGuid) , ('$CurrentMailboxPhoto$', $CurrentMailboxPhotoguid), ('$CurrentUserManagerPhoto$', $CurrentUserManagerPhotoGuid), ('$CurrentUserPhoto$', $CurrentUserPhotoGuid))) {
                            if (
                                $(if ($tempImageSourceFullname) { ((Split-Path $tempImageSourceFullname -Leaf) -ilike "*$($Variablename[0])*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($Variablename[0])*") })
                            ) {
                                if ($null -ne $ReplaceHash[$Variablename[0]]) {
                                    $tempImageSourceFullname = (Join-Path -Path $script:tempDir -ChildPath ($Variablename[0] + $Variablename[1] + '.jpeg'))
                                }
                            } elseif (
                                $(if ($tempImageSourceFullname) { ((Split-Path $tempImageSourceFullname -Leaf) -ilike "*$($Variablename[0] -ireplace '\$$', 'DELETEEMPTY$')*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($Variablename[0] -ireplace '\$$', 'DELETEEMPTY$')*") })
                            ) {
                                if ($null -ne $ReplaceHash[$Variablename[0]]) {
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

                    # Other images
                    if (
                        $(if ($tempImageSourceFullname) { ((Split-Path $tempImageSourceFullname -Leaf) -ilike '*$*DELETEEMPTY$*') }) -or
                        $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike '*$*DELETEEMPTY$*') })
                    ) {
                        foreach ($Variablename in  @(@($ReplaceHash.Keys) | Where-Object { $_ -inotin @('$CurrentMailboxPhoto$', '$CurrentMailboxManagerPhoto$', '$CurrentUserPhoto$', '$CurrentUserManagerPhoto$') })) {
                            $tempImageVariableString = $Variablename -ireplace '\$$', 'DELETEEMPTY$'

                            if (
                                $(if ($tempImageSourceFullname) { ((Split-Path $tempImageSourceFullname -Leaf) -ilike "*$($tempImageVariableString)*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($tempImageVariableString)*") })
                            ) {
                                if ($ReplaceHash[$Variablename]) {
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

                    foreach ($replaceKey in @($replaceHash.Keys | Where-Object { $_ -inotin ('$CurrentMailboxManagerPhoto$', '$CurrentMailboxPhoto$', '$CurrentUserManagerPhoto$', '$CurrentUserPhoto$', '$CurrentMailboxManagerPhotodeleteempty$', '$CurrentMailboxPhotodeleteempty$', '$CurrentUserManagerPhotodeleteempty$', '$CurrentUserPhotodeleteempty$') } | Sort-Object -Culture $TemplateFilesSortCulture)) {
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
                    ($tempImageSourceFullname) -and
                    ($image.linkformat.sourcefullname) -and
                    ($tempImageSourceFullname -ine $image.linkformat.sourcefullname)
                    ) {
                        $image.linkformat.sourcefullname = $tempImageSourceFullname
                    }

                    if ($null -ne $tempImageAlternativeText) {
                        $image.AlternativeText = $tempImageAlternativeText
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
                Write-Host "$Indent        Error replacing picture variables in Word. Exit." -ForegroundColor Red
                Write-Host "$Indent        If the error says 'Access denied', your environment may require to assign a Microsoft Purview Information Protection sensitivity label to your DOCX templates." -ForegroundColor Red
                $error[0]
                exit 1
            }

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

                foreach ($replaceKey in @($replaceHash.Keys | Where-Object { ($_ -inotin ('$CurrentMailboxManagerPhoto$', '$CurrentMailboxPhoto$', '$CurrentUserManagerPhoto$', '$CurrentUserPhoto$', '$CurrentMailboxManagerPhotodeleteempty$', '$CurrentMailboxPhotodeleteempty$', '$CurrentUserManagerPhotodeleteempty$', '$CurrentUserPhotodeleteempty$')) -and ($tempWordText -imatch [regex]::escape($_)) } | Sort-Object -Culture $TemplateFilesSortCulture )) {
                    $script:COMWord.Selection.Find.Execute($replaceKey, $MatchCase, $MatchWholeWord, `
                            $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, `
                            $Wrap, $Format, $(($replaceHash.$replaceKey -ireplace "`r`n", '^p') -ireplace "`n", '^l'), $ReplaceAll) | Out-Null
                }

                # Restore original view
                $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal

                $tempWordText = $null

                # Replace in field codes
                foreach ($field in $script:COMWord.ActiveDocument.Fields) {
                    $tempWordFieldCodeOriginal = $field.Code.Text
                    $tempWordFieldCodeNew = $tempWordFieldCodeOriginal

                    foreach ($replaceKey in @($replaceHash.Keys | Where-Object { ($_ -inotin ('$CurrentMailboxManagerPhoto$', '$CurrentMailboxPhoto$', '$CurrentUserManagerPhoto$', '$CurrentUserPhoto$', '$CurrentMailboxManagerPhotodeleteempty$', '$CurrentMailboxPhotodeleteempty$', '$CurrentUserManagerPhotodeleteempty$', '$CurrentUserPhotodeleteempty$')) } | Sort-Object -Culture $TemplateFilesSortCulture )) {
                        $tempWordFieldCodeNew = $tempWordFieldCodeNew -ireplace [regex]::escape($replaceKey), $($replaceHash.$replaceKey)
                    }

                    if ($tempWordFieldCodeOriginal -ne $tempWordFieldCodeNew) {
                        $field.Code.Text = $tempWordFieldCodeNew
                    }
                }
            } catch {
                Write-Host "$Indent        Error replacing non-picture variables in Word. Exit." -ForegroundColor Red
                Write-Host "$Indent        If the error says 'Access denied', your environment may require to assign a Microsoft Purview Information Protection sensitivity label to your DOCX templates." -ForegroundColor Red
                $error[0]
                exit 1
            }


            # Save changed document, it's later used for export to .htm, .rtf and .txt
            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatDocumentDefault')

            try {
                # Overcome Word security warning when export contains embedded pictures
                if ((Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security\DisableWarningOnIncludeFieldsUpdate") -eq $false) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                # Save
                $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)

                # Restore original security setting
                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                } else {
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }
            } catch {
                # Restore original security setting after error
                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                } else {
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }

                Start-Sleep -Seconds 2

                # Overcome Word security warning when export contains embedded pictures
                if ((Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security\DisableWarningOnIncludeFieldsUpdate") -eq $false) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                # Save
                $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)

                # Restore original security setting
                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                } else {
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }
            }

            # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
            # Close the document to remove in-memory references to already deleted images
            $script:COMWord.ActiveDocument.Saved = $true
            $script:COMWord.ActiveDocument.Close($false)

            # Export to .htm
            Write-Host "$Indent      Export to HTM format"
            $path = $([System.IO.Path]::ChangeExtension($path, '.docx'))
            $script:COMWord.Documents.Open($path, $false) | Out-Null

            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatFilteredHTML')
            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

            $script:WordWebOptions = $script:COMWord.ActiveDocument.WebOptions

            $script:COMWord.ActiveDocument.WebOptions.TargetBrowser = 4 # IE6, which is the maximum
            $script:COMWord.ActiveDocument.WebOptions.BrowserLevel = 2 # IE6, which is the maximum
            $script:COMWord.ActiveDocument.WebOptions.AllowPNG = $true
            $script:COMWord.ActiveDocument.WebOptions.OptimizeForBrowser = $true
            $script:COMWord.ActiveDocument.WebOptions.RelyOnCSS = $true
            $script:COMWord.ActiveDocument.WebOptions.RelyOnVML = $true
            $script:COMWord.ActiveDocument.WebOptions.Encoding = 65001 # Outlook uses 65001 (UTF8) for .htm, but 1200 (UTF16LE a.k.a Unicode) for .txt
            $script:COMWord.ActiveDocument.WebOptions.OrganizeInFolder = $true
            $script:COMWord.ActiveDocument.WebOptions.PixelsPerInch = 96
            $script:COMWord.ActiveDocument.WebOptions.ScreenSize = 10 # 1920x1200
            $script:COMWord.ActiveDocument.WebOptions.UseLongFileNames = $true

            $script:COMWord.ActiveDocument.WebOptions.UseDefaultFolderSuffix()
            $pathHtmlFolderSuffix = $script:COMWord.ActiveDocument.WebOptions.FolderSuffix

            try {
                # Overcome Word security warning when export contains embedded pictures
                if ((Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security\DisableWarningOnIncludeFieldsUpdate") -eq $false) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                # Save
                $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)

                # Restore original security setting
                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                } else {
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }
            } catch {
                # Restore original security setting after error
                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                } else {
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }

                Start-Sleep -Seconds 2

                # Overcome Word security warning when export contains embedded pictures
                if ((Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security\DisableWarningOnIncludeFieldsUpdate") -eq $false) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                # Save
                $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)

                # Restore original security setting
                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                } else {
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }
            }

            # Restore original WebOptions
            try {
                if ($script:WordWebOptions) {
                    foreach ($property in @('TargetBrowser', 'BrowserLevel', 'AllowPNG', 'OptimizeForBrowser', 'RelyOnCSS', 'RelyOnVML', 'Encoding', 'OrganizeInFolder', 'PixelsPerInch', 'ScreenSize', 'UseLongFileNames')) {
                        $script:COMWord.ActiveDocument.WebOptions.$property = $script:WordWebOptions.$property
                    }
                }
            } catch {}

            # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
            $script:COMWord.ActiveDocument.Saved = $true

            if ($DocxHighResImageConversion) {
                Write-Host "$Indent        Export high-res images"
                if (-not $BenefactorCircleLicenseFile) {
                    $script:COMWord.ActiveDocument.Close($false)

                    Write-Host "$Indent          The 'DocxHighResImageConversion' feature is reserved for Benefactor Circle members." -ForegroundColor Yellow
                    Write-Host "$Indent          Find out details in '.\docs\Benefactor Circle'." -ForegroundColor Yellow
                } else {
                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DocxHighResImageConversion()

                    if ($FeatureResult -ne 'true') {
                        try {
                            $script:COMWord.ActiveDocument.Close($false)
                        } catch {
                        }
                        Write-Host "$Indent          Error converting high resolution images from DOCX template." -ForegroundColor Yellow
                        Write-Host "$Indent          $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                $script:COMWord.ActiveDocument.Close($false)
            }
        }

        Write-Host "$Indent        Copy HTM image width and height attributes to style attribute"
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

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

        $tempFileContent = $html.documentelement.outerhtml
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($html) | Out-Null
        Remove-Variable -Name 'html'
        $tempFileContent | Out-File -LiteralPath $path -Encoding UTF8 -Force

        if ($MoveCSSInline) {
            Write-Host "$Indent        Move CSS inline"

            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
            $tempFileContent = Get-Content -LiteralPath $path -Encoding UTF8 -Raw

            # Use a separate runspace for PreMailer.Net, as there are DLL conflicts in PowerShell 5.x with Invoke-RestMethod
            # Do not use jobs, as they fall back to Constrained Language Mode in secured environments, which makes Import-Module fail
            [System.IO.File]::WriteAllText($path, $(MoveCssInline($tempFileContent)), (New-Object System.Text.UTF8Encoding($False)))
        }

        Write-Host "$Indent        Add marker to final HTM file"
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
        $tempFileContent = Get-Content -LiteralPath $path -Encoding UTF8 -Raw

        if ($tempFileContent -inotlike "*$HTMLMarkerTag*") {
            if ($tempFileContent -ilike '*<head>*') {
                $tempFileContent = $tempFileContent -ireplace '<HEAD>', "<HEAD> $($HTMLMarkerTag)"
            } else {
                $tempFileContent = $tempFileContent -ireplace '<HTML>', "<HTML><HEAD> $($HTMLMarkerTag) </HEAD>"
            }
        }

        Write-Host "$Indent        Modify connected folder name"

        foreach ($pathConnectedFolderName in $pathConnectedFolderNames) {
            $tempFileContent = $tempFileContent -ireplace ('(\s*src=")(' + [regex]::escape($pathConnectedFolderName) + '\/)'), ('$1' + "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.value)).files/")
            Rename-Item (Join-Path -Path (Split-Path $path) -ChildPath $($pathConnectedFolderName)) $([System.IO.Path]::GetFileNameWithoutExtension($Signature.value) + '.files') -ErrorAction SilentlyContinue
        }

        [System.IO.File]::WriteAllText($path, $tempFileContent, (New-Object System.Text.UTF8Encoding($False)))


        if (-not $ProcessOOF) {
            if ($EmbedImagesInHtml) {
                Write-Host "$Indent        Embed local images"

                [System.IO.File]::WriteAllText($path, $tempFileContent, (New-Object System.Text.UTF8Encoding($False)))

                ConvertToSingleFileHTML $path $path
            }
        } else {
            ConvertToSingleFileHTML $path ((Join-Path -Path $script:tempDir -ChildPath $Signature.Value))
        }

        if (-not $ProcessOOF) {
            if ($CreateRtfSignatures) {
                Write-Host "$Indent      Export to RTF format"
                # If possible, use .docx file to avoid problems with MS Information Protection
                if ($UseHtmTemplates) {
                    $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
                    $script:COMWord.Documents.Open($path, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, 65001) | Out-Null
                } else {
                    $path = $([System.IO.Path]::ChangeExtension($path, '.docx'))
                    $script:COMWord.Documents.Open($path, $false) | Out-Null
                }


                $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatRTF')
                $path = $([System.IO.Path]::ChangeExtension($path, '.rtf'))

                try {
                    # Overcome Word security warning when export contains embedded pictures
                    if ((Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security\DisableWarningOnIncludeFieldsUpdate") -eq $false) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                    }

                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore

                    if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                    }

                    # Save
                    $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)

                    # Restore original security setting
                    if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                        Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                    } else {
                        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                    }
                } catch {
                    # Restore original security setting after error
                    if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                        Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                    } else {
                        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                    }

                    Start-Sleep -Seconds 2

                    # Overcome Word security warning when export contains embedded pictures
                    if ((Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security\DisableWarningOnIncludeFieldsUpdate") -eq $false) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                    }

                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore

                    if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                    }

                    # Save
                    $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)

                    # Restore original security setting
                    if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                        Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                    } else {
                        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                    }
                }

                # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
                # Close the document as conversion to .rtf happens from .htm
                $script:COMWord.ActiveDocument.Saved = $true
                $script:COMWord.ActiveDocument.Close($false)

                Write-Host "$Indent        Shrink RTF file"
                $((Get-Content -LiteralPath $path -Raw -Encoding Ascii) -ireplace '\{\\nonshppict[\s\S]*?\}\}', '') | Set-Content -LiteralPath $path -Encoding Ascii
            }

            if ($CreateTxtSignatures) {
                Write-Host "$Indent      Export to TXT format"

                $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

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

                $path = $([System.IO.Path]::ChangeExtension($path, '.txt'))

                $html.body.innerText | Out-File -LiteralPath $path -Encoding Unicode -Force # Outlook uses 65001 (UTF8) for .htm, but 1200 (UTF16LE a.k.a Unicode) for .txt
            }
        }

        if (-not $ProcessOOF) {
            foreach ($SignaturePath in $SignaturePaths) {
                Write-Host "$Indent      Copy signature files to '$SignaturePath'"

                RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.htm')))

                foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                    RemoveItemAlternativeRecurse -LiteralPath ((Join-Path -Path $SignaturePath -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.value))") + $ConnectedFilesFolderName)
                }

                Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.htm')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.htm')))) -Force

                if ($EmbedImagesInHtml -eq $false) {
                    if (Test-Path (Join-Path -Path (Split-Path $path) -ChildPath "$([System.IO.Path]::ChangeExtension($Signature.value, '.files'))")) {
                        Copy-Item -LiteralPath (Join-Path -Path (Split-Path $path) -ChildPath "$([System.IO.Path]::ChangeExtension($Signature.value, '.files'))") -Destination $SignaturePath -Force -Recurse
                    }
                }

                if ($CreateRtfSignatures -eq $true) {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))
                    Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.rtf')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))) -Force
                } else {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))
                }

                if ($CreateTxtSignatures -eq $true) {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))
                    Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.txt')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))) -Force
                } else {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))
                }


                if ($SignatureFilesWriteProtect.containskey($TemplateIniSettingsIndex)) {
                    Write-Host "$Indent      Write protect signature files"
                    @('.htm', '.rtf', '.txt') | ForEach-Object {
                        $file = Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, $_))
                        if (Test-Path -Path $file -PathType Leaf) {
                            (Get-Item $file -Force).Attributes += 'ReadOnly'
                        }
                    }
                }
            }
        }

        Write-Host "$Indent      Remove temporary files"
        foreach ($extension in ('.docx', '.htm', '.rtf', '.txt')) {
            Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, $extension)) -ErrorAction SilentlyContinue | Out-Null
            if ($pathHighResHtml) {
                Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($pathHighResHtml, $extension)) -ErrorAction SilentlyContinue | Out-Null
            }
        }

        Foreach ($file in @(Get-ChildItem -Path ("$($script:tempDir)\*" + [System.IO.Path]::GetFileNameWithoutExtension($path) + '*') -Directory).FullName) {
            Remove-Item -LiteralPath $file -Force -Recurse -ErrorAction SilentlyContinue
        }

        if ($pathHighResHtml) {
            Foreach ($file in @(Get-ChildItem -Path ("$($script:tempDir)\*" + [System.IO.Path]::GetFileNameWithoutExtension($pathHighResHtml) + '*') -Directory).FullName) {
                Remove-Item -LiteralPath $file -Force -Recurse -ErrorAction SilentlyContinue
            }
        }

        Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath $([System.IO.Path]::ChangeExtension($signature.value, '.files'))) -Force -Recurse -ErrorAction SilentlyContinue
    }

    if ((-not $ProcessOOF)) {
        # Set default signature for new emails
        if ($SignatureFilesDefaultNew.containskey($TemplateIniSettingsIndex)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    if (-not $SimulateUser) {
                        if ($RegistryPaths[$j] -ilike '*\9375CFF0413111d3B88A00104B2A6676\*') {
                            Write-Host "$Indent      Set signature as default for new messages (Outlook profile '$(($RegistryPaths[$j] -split '\\')[8])')"

                            if (($script:CurrentUserDummyMailbox -ne $true)) {
                                if ($OutlookFileVersion -ge '16.0.0.0') {
                                    New-ItemProperty -Path $RegistryPaths[$j] -Name 'New Signature' -PropertyType String -Value ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')) -Force | Out-Null
                                } else {
                                    New-ItemProperty -Path $RegistryPaths[$j] -Name 'New Signature' -PropertyType Binary -Value ([byte[]](([System.Text.Encoding]::Unicode.GetBytes(((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')) + "`0")))) -Force | Out-Null
                                }
                            }
                        } else {
                            $script:CurrentUserDummyMailboxDefaultSigNew = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                        }
                    } else {
                        $WshShell = New-Object -ComObject WScript.Shell

                        @('htm', 'rtf', 'txt') | ForEach-Object {
                            if (Test-Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)"))) {
                                $script:CurrentUserDummyMailboxDefaultSigNew = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                                $ShortcutTempPath = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath "$((New-Guid).guid).lnk"

                                $Shortcut = $WshShell.CreateShortcut($ShortcutTempPath)
                                $Shortcut.TargetPath = (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)"))
                                $Shortcut.Save()

                                Copy-Item -LiteralPath $ShortcutTempPath -Destination $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultNew $($_).lnk")) -Force

                                Remove-Item $ShortcutTempPath -Force
                            }
                        }
                    }
                }
            }
        }

        # Set default signature for replies and forwarded emails
        if ($SignatureFilesDefaultReplyFwd.containskey($TemplateIniSettingsIndex)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    if (-not $SimulateUser) {
                        if ($RegistryPaths[$j] -ilike '*\9375CFF0413111d3B88A00104B2A6676\*') {
                            Write-Host "$Indent      Set signature as default for reply/forward messages (Outlook profile '$(($RegistryPaths[$j] -split '\\')[8])')"

                            if (($script:CurrentUserDummyMailbox -ne $true)) {
                                if ($OutlookFileVersion -ge '16.0.0.0') {
                                    New-ItemProperty -Path $RegistryPaths[$j] -Name 'Reply-Forward Signature' -PropertyType String -Value ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')) -Force | Out-Null
                                } else {
                                    New-ItemProperty -Path $RegistryPaths[$j] -Name 'Reply-Forward Signature' -PropertyType Binary -Value ([byte[]](([System.Text.Encoding]::Unicode.GetBytes(((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')) + "`0")))) -Force | Out-Null
                                }
                            }
                        } else {
                            $script:CurrentUserDummyMailboxDefaultSigReply = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                        }
                    } else {
                        $WshShell = New-Object -ComObject WScript.Shell

                        @('htm', 'rtf', 'txt') | ForEach-Object {
                            if (Test-Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)"))) {
                                $script:CurrentUserDummyMailboxDefaultSigReply = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                                $ShortcutTempPath = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath "$((New-Guid).guid).lnk"

                                $Shortcut = $WshShell.CreateShortcut($ShortcutTempPath)
                                $Shortcut.TargetPath = (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)"))
                                $Shortcut.Save()

                                Copy-Item -LiteralPath $ShortcutTempPath -Destination $((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultReplyFwd $($_).lnk")) -Force

                                Remove-Item $ShortcutTempPath -Force
                            }
                        }
                    }
                }
            }
        }
    }
}


function CheckADConnectivity {
    param (
        [array]$CheckDomains,
        [string]$CheckProtocolText,
        [string]$Indent
    )
    [void][runspacefactory]::CreateRunspacePool()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 25)
    $RunspacePool.Open()

    for ($DomainNumber = 0; $DomainNumber -lt $CheckDomains.count; $DomainNumber++) {
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
                    $null = ([ADSI]"$(($Search.FindOne()).path)")
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
        foreach ($job in $script:jobs) {
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
                        Write-Host "$Indent  If this error is permanent, check firewalls, DNS and AD trust. Consider parameter TrustsToCheckForGroups." -ForegroundColor Red

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
    }
    return $returnvalue
}


function MoveCssInline {
    param (
        $HtmlCode
    )
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
                Import-Module (Join-Path -Path $path -ChildPath 'PreMailer.Net.dll')
                Write-Debug $([PreMailer.Net.PreMailer]::MoveCssInline($HtmlCode).html)
            } catch {
                Write-Debug 'Failed'
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
        foreach ($job in $script:jobs) {
            if (($null -eq $job.StartTime) -and ($job.Powershell.Streams.Debug[0].Message -imatch 'Start')) {
                $StartTicks = $job.powershell.Streams.Debug[0].Message -ireplace '[^0-9]'
                $job.StartTime = [Datetime]::MinValue + [TimeSpan]::FromTicks($StartTicks)
            }

            if ($null -ne $job.StartTime) {
                if ((($job.handle.IsCompleted -eq $true) -and ($job.Done -eq $false)) -or (($job.Done -eq $false) -and ((New-TimeSpan -Start $job.StartTime -End (Get-Date)).TotalSeconds -ge 5))) {
                    $data = $job.Object[0..$(($job.object).count - 1)]
                    if ($job.Powershell.Streams.Debug[1].Message -imatch 'Failed') {
                        $returnvalue = $HtmlCode
                    } else {
                        $returnvalue = $job.Powershell.Streams.Debug[1].Message
                    }
                    $job.Done = $true
                }
            }
        }
    }
    return $returnvalue
}


function CheckPath([string]$path, [switch]$silent = $false, [switch]$create = $false) {
    if ($create -eq $false) {
        if (($path.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($path -ilike '*@ssl\*')) {
            $path = $path -ireplace '@ssl\\', '\'
            $path = ([uri]::UnescapeDataString($path) -ireplace ('https://', '\\'))
            $path = ([System.URI]$path).AbsoluteURI -ireplace 'file:\/\/(.*?)\/(.*)', '\\${1}@SSL\$2' -ireplace '/', '\'
            $path = [uri]::UnescapeDataString($path)
        } else {
            try {
                $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
                $path = ([System.URI]$path).absoluteuri -ireplace 'file:///', '' -ireplace 'file://', '\\' -ireplace '/', '\'
                $path = [uri]::UnescapeDataString($path)
            } catch {
                if ($silent -eq $false) {
                    Write-Host ': ' -NoNewline
                    Write-Host "Problem connecting to or reading from folder '$path'. Exit." -ForegroundColor Red
                    exit 1
                }
            }
        }

        if (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) {
            # Reconnect already connected network drives at the OS level
            # New-PSDrive is not enough for this
            foreach ($NetworkConnection in @(Get-CimInstance Win32_NetworkConnection)) {
                & net use $NetworkConnection.LocalName $NetworkConnection.RemoteName 2>&1 | Out-Null
            }

            if (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) {
                # Connect network drives
                '`r`n' | & net use "$path" 2>&1 | Out-Null
                try {
                    (Test-Path -LiteralPath $path -ErrorAction Stop) | Out-Null
                } catch {
                    if ($_.CategoryInfo.Category -eq 'PermissionDenied') {
                        & net use "$path" 2>&1
                    }
                }
                & net use "$path" /d 2>&1 | Out-Null
            }

            if (($path -ilike '*@ssl\*') -and (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue))) {
                if (((Get-Service -ServiceName 'WebClient' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Status -ieq 'Running') -eq $false) {
                    Write-Host
                    Write-Host 'WebClient service not running.' -ForegroundColor Red
                }

                Try {
                    # Add site to trusted sites in internet options
                    New-Item ('HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\' + (New-Object System.Uri -ArgumentList ($path -ireplace ('@SSL', ''))).Host) -Force | New-ItemProperty -Name * -Value 1 -Type DWORD -Force | Out-Null

                    # Open site in new IE process
                    $oIE = New-Object -com InternetExplorer.Application
                    $oIE.Visible = $false
                    $oIE.Navigate2('https://' + ((($path -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($oIE) | Out-Null
                    Remove-Variable -Name 'oIE'

                    # Wait until an IE tab with the corresponding URL is open
                    $app = New-Object -com shell.application
                    $i = 0
                    while ($i -lt 1) {
                        $i += @($app.windows() | Where-Object { $_.LocationURL -ilike ('*' + ([uri]::EscapeUriString(((($path -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))) + '*') }).count
                        Start-Sleep -Milliseconds 50
                    }

                    # Wait until the corresponding URL is fully loaded, then close the tab
                    foreach ($window in @($app.windows() | Where-Object { $_.LocationURL -ilike ('*' + ([uri]::EscapeUriString(((($path -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))) + '*') })) {
                        while ($window.busy) {
                            Start-Sleep -Milliseconds 50
                        }
                        $window.quit([ref]$false)
                    }

                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
                    Remove-Variable -Name 'app'
                } catch {
                }
            }
        }

        if ((Test-Path -LiteralPath $path) -eq $false) {
            if ($silent -eq $false) {
                Write-Host ': ' -NoNewline
                Write-Host "Problem connecting to or reading from folder '$path'. Exit." -ForegroundColor Red

                exit 1
            } else {
                return $false
            }
        } else {
            if ($silent -eq $false) {
                Write-Host
            } else {
                return $true
            }
        }
    } else {
        if ($path.StartsWith('https://', 'CurrentCultureIgnoreCase')) {
            $path = ((([uri]::UnescapeDataString($path) -ireplace ('https://', '\\')) -ireplace ('(.*?)/(.*)', '${1}@SSL\$2')) -ireplace ('/', '\'))
        }
        $pathTemp = $path
        for ($i = (($path.ToCharArray() | Where-Object { $_ -eq '\' } | Measure-Object).Count); $i -ge 0; $i--) {
            if ((CheckPath $pathTemp -Silent) -eq $true) {
                if (-not (Test-Path $pathTemp -PathType Container -ErrorAction SilentlyContinue)) {
                    Write-Host ': ' -NoNewline
                    Write-Host "'$pathTemp' is a file, '$path' not valid. Exit." -ForegroundColor Red
                    exit 1
                }

                if ($pathTemp -eq $path) {
                    break
                } else {
                    New-Item -ItemType Directory -Path $path -ErrorAction SilentlyContinue | Out-Null
                    if (Test-Path -Path $path -PathType Container) {
                        break
                    }
                }
            } else {
                $pathTemp = Split-Path ($pathTemp -ireplace '@SSL', '') -Parent
            }
        }

        if ((checkpath $path -silent) -ne $true) {
            Write-Host ': ' -NoNewline
            Write-Host "Problem connecting to or reading from folder '$path'. Exit." -ForegroundColor Red
            exit 1
        } else {
            Write-Host
        }
    }
}


function GraphGetToken {
    Write-Verbose '      Authentication'

    if ($SimulateAndDeployGraphCredentialFile) {
        Write-Verbose "        Via SimulateAndDeployGraphCredentialFile '$SimulateAndDeployGraphCredentialFile'"

        try {
            try {
                $auth = Import-Clixml -Path $SimulateAndDeployGraphCredentialFile
            } catch {
                Start-Sleep -Seconds 2
                $auth = Import-Clixml -Path $SimulateAndDeployGraphCredentialFile
            }

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
        $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $CloudEnvironmentEnvironmentName -TenantId $(if ($script:CurrentUser) { ($script:CurrentUser -split '@')[1] } else { 'organizations' }) -RedirectUri 'http://localhost' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

        try {
            Write-Verbose '        Via IntegratedWindowsAuth'
            $auth = $script:msalClientApp | Get-MsalToken -AzureCloudInstance $CloudEnvironmentEnvironmentName -LoginHint $(if ($script:CurrentUser) { $script:CurrentUser } else { '' }) -Scopes "$($CloudEnvironmentGraphApiEndpoint)/.default" -IntegratedWindowsAuth -Timeout (New-TimeSpan -Minutes 1)
        } catch {
            Write-Verbose $error[0]

            try {
                Write-Verbose '        Via Silent with LoginHint'
                $auth = $script:msalClientApp | Get-MsalToken -AzureCloudInstance $CloudEnvironmentEnvironmentName -LoginHint $(if ($script:CurrentUser) { $script:CurrentUser } else { '' }) -Scopes "$($CloudEnvironmentGraphApiEndpoint)/.default" -Silent -ForceRefresh -Timeout (New-TimeSpan -Minutes 1)
            } catch {
                Write-Verbose $error[0]

                try {
                    Write-Verbose '        Via Prompt with LoginHint and Timeout'

                    if (-not [string]::IsNullOrWhitespace($GraphHtmlMessageboxText)) {
                        Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

                        $window = New-Object System.Windows.Window
                        $window.Width = 1
                        $window.Height = 1
                        $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
                        $window.ShowActivated = $false
                        $window.Topmost = $true
                        $window.Show()
                        $window.Hide()
                        [System.Windows.MessageBox]::Show($window, "$($GraphHtmlMessageboxText)", $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }), [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information, [System.Windows.MessageBoxResult]::None)
                        $window.close()
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

                    $auth = $script:msalClientApp | Get-MsalToken -AzureCloudInstance $CloudEnvironmentEnvironmentName -LoginHint $(if ($script:CurrentUser) { $script:CurrentUser } else { '' }) -Scopes "$($CloudEnvironmentGraphApiEndpoint)/.default" -Interactive -Timeout (New-TimeSpan -Minutes 5) -Prompt 'NoPrompt' -UseEmbeddedWebView:$false @MsalInteractiveParams
                } catch {
                    Write-Verbose '        No authentication possible'
                    $auth = $null
                    return @{
                        error             = (($error[0] | Out-String) + @"
No Graph authentication possible.
1. Did you follow the Quick Start Guide in '.\docs\README' and configure the Entra ID/Azure AD app correctly?
2. If the "Via Prompt with LoginHint and Timeout" authentication message is diplayed:
   - Does a browser (the system default browser, if configured) open and ask for authentication?
     - Yes:
       - Check if the correct user account is selected/entered and if the authentication is successful
       - Check if authentication happens within two minutes
       - Ensure that access to 'http://localhost' is allowed ('https://localhost' is currently not technically feasible, see 'https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/System-Browser-on-.Net-Core' and 'https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/MSAL.NET-uses-web-browser' for details)
     - No:
       - Run Set-OutlookSignatures in a new PowerShell session
       - Check if a default browser is set and if "start https://github.com/Set-OutlookSignatures/Set-OutlookSignatures" opens it
       - Make sure that Set-OutlookSignatures is executed in the security context of the currently logged-in user
       - Make sure that the current PowerShell session allows TLS 1.2 (see https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/85 for details)
3. Run Set-OutlookSignatures with the "-Verbose" parameter and check for authentication messages
4. Delete the MSAL.PS Graph token cache file '$(try { [TokenCacheHelper]::CacheFilePath } catch { 'TokenCacheHelper error - file info not available!' } )'.
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

        if ($auth) {
            $authExo = $script:msalClientApp | Get-MsalToken -AzureCloudInstance $CloudEnvironmentEnvironmentName -LoginHint $(if ($script:CurrentUser) { $script:CurrentUser } else { '' }) -Scopes "$($CloudEnvironmentExchangeOnlineEndpoint)/.default" -Silent -ForceRefresh

            try {
                $script:AuthorizationHeader = @{
                    Authorization = $auth.CreateAuthorizationHeader()
                }

                $script:ExoAuthorizationHeader = @{
                    Authorization = $authExo.CreateAuthorizationHeader()
                }

                return @{
                    error             = $false
                    AccessToken       = $auth.AccessToken
                    AuthHeader        = $script:AuthorizationHeader
                    AccessTokenExo    = $authExo.AccessToken
                    AuthHeaderExo     = $authExo.createauthorizationheader()
                    AppAccessToken    = $null
                    AppAuthHeader     = $null
                    AppAccessTokenExo = $null
                    AppAuthHeaderExo  = $null
                }
            } catch {
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

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/me?`$select=" + [System.Web.HttpUtility]::UrlEncode(($GraphUserProperties -join ','))
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            $local:uri = $local:pagedResults.'@odata.nextlink'
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
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
            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            $local:uri = $local:pagedResults.'@odata.nextlink'
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

    $user = GraphGetUpnFromSmtp($user)

    if ($user.properties.value.userprincipalname) {
        try {
            $local:x = @($GraphUserProperties | Select-Object -Unique) -join ','

            $requestBody = @{
                Method      = 'Get'
                Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user.properties.value.userprincipalname)?`$select=" + [System.Web.HttpUtility]::UrlEncode($local:x)
                Headers     = $authHeader
                ContentType = 'Application/Json; charset=utf-8'
            }

            $OldProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            $local:x = @()
            $local:uri = $null

            do {
                if ($local:uri) {
                    $requestBody['Uri'] = $local:uri
                }

                $local:pagedResults = Invoke-RestMethod @requestBody
                $local:x += $local:pagedResults

                $local:uri = $local:pagedResults.'@odata.nextlink'
            } until (!($local:uri))


            if (($user.properties.value.userprincipalname -ieq $script:CurrentUser) -and ((-not $SimulateUser) -or ($SimulateUser -and $SimulateAndDeployGraphCredentialFile -and ($authHeader -eq $script:AppAuthorizationHeader))) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true) -or ($MirrorLocalSignaturesToCloud -eq $true))) {
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
                    $local:uri = $null

                    do {
                        if ($local:uri) {
                            $requestBody['Uri'] = $local:uri
                        }

                        $local:pagedResults = Invoke-RestMethod @requestBody
                        $local:y += $local:pagedResults

                        $local:uri = $local:pagedResults.'@odata.nextlink'
                    } until (!($local:uri))

                    $local:x | Add-Member -MemberType NoteProperty -Name 'mailboxSettings' -Value $local:y.mailboxSettings
                } catch {
                    Write-Host "      Problem getting mailboxSettings for '$($script:CurrentUser)' from Microsoft Graph." -ForegroundColor Yellow
                    $error[0]
                    Write-Host '      This is a Microsoft Graph API problem, which can only be solved by Microsoft itself.' -ForegroundColor Yellow
                    Write-Host '      To be able to continue, SetCurrentUserOutlookWebSignature and SetCurrentUserOOFMessage are now disabled.' -ForegroundColor Yellow

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

        if (($user.properties.value.userprincipalname -ieq $script:CurrentUser) -and ($SimulateUser -and $SimulateAndDeployGraphCredentialFile -and ($authHeader -eq $script:AuthorizationHeader))) {
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
            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            $local:uri = $local:pagedResults.'@odata.nextlink'
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
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
            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            $local:uri = $local:pagedResults.'@odata.nextlink'
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
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

        $local:x = [System.IO.File]::ReadAllBytes($local:tempFile)

        Remove-Item $local:tempFile -Force -ErrorAction SilentlyContinue
    } catch {
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

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/groups?`$select=securityidentifier&`$filter=" + [System.Web.HttpUtility]::UrlEncode($filter)
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            $local:uri = $local:pagedResults.'@odata.nextlink'
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
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

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users?`$select=securityidentifier&`$filter=" + [System.Web.HttpUtility]::UrlEncode($filter)
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            $local:uri = $local:pagedResults.'@odata.nextlink'
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
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


function ExoGenericQuery ([Parameter(Mandatory = $true)] [string]$method, [Parameter(Mandatory = $true)] [uri]$uri, [Parameter(Mandatory = $true)] [AllowEmptyString()] [string]$body, [Parameter(Mandatory = $true)] [bool]$isLargeSetting ) {
    $error.clear()
    try {
        $requestBody = @{
            Method      = $method
            Uri         = $uri
            Headers     = $script:ExoAuthorizationHeader
            ContentType = 'application/json; charset=utf-8'
        }

        if ($body) {
            $requestBody['Body'] = $body
        }

        if ($isLargeSetting -eq $true) {
            $requestBody['Headers']['x-islargesetting'] = 'true'
        } else {
            $requestBody['Headers']['x-islargesetting'] = 'false'
        }

        if ($PrimaryMailboxAddress) {
            $requestBody['Headers']['x-anchormailbox'] = "SMTP:$($PrimaryMailboxAddress)"
        }

        $requestBody['Headers']['x-overridetimestamp'] = 'true'

        $requestBody['Headers']['content-type'] = 'Application/Json; charset=utf-8'

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            $local:uri = $local:pagedResults.'@odata.nextlink'
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
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


function GetIniContent ($filePath) {
    $local:ini = [ordered]@{}
    $local:SectionIndex = -1
    if ($filePath -ne '') {
        try {
            Write-Verbose '    Original ini content'

            foreach ($line in @(Get-Content -LiteralPath $FilePath -Encoding UTF8 -ErrorAction Stop)) {
                Write-Verbose "      $line"
                switch -regex ($line) {
                    # Comments starting with ; or #, or empty line, whitespace(s) before are ignored
                    '(^\s*(;|#))|(^\s*$)' { continue }

                    # Section in square brackets, whitespace(s) before and after brackets are ignored
                    '^\s*\[(.+)\]\s*' {
                        $local:section = ($matches[1]).trim().trim('"').trim('''')
                        if ($null -ne $local:section) {
                            $local:SectionIndex++
                            $local:ini["$($local:SectionIndex)"] = @{ '<Set-OutlookSignatures template>' = $local:section }
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
            Write-Host
            Write-Host "Error accessing '$FilePath'. Exit." -ForegroundColor red
            $Error[0]
            exit 1
        }
    }
    return $local:ini
}


function ConvertPath ([ref]$path) {
    if ($path) {
        if (($path.value.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($path.value -ilike '*@ssl\*')) {
            $path.value = $path.value -ireplace '@ssl\\', '\'
            $path.value = ([uri]::UnescapeDataString($path.value) -ireplace ('https://', '\\'))
            $path.value = ([System.URI]$path.value).AbsoluteURI -ireplace 'file:\/\/(.*?)\/(.*)', '\\${1}@SSL\$2' -ireplace '/', '\'
            $path.value = [uri]::UnescapeDataString($path.value)
        } else {
            $path.value = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path.value)
            $path.value = ([System.URI]$path.value).absoluteuri -ireplace 'file:///', '' -ireplace 'file://', '\\' -ireplace '/', '\'
            $path.value = [uri]::UnescapeDataString($path.value)
        }
    }
}


function RemoveItemAlternativeRecurse {
    # Function to avoid problems with OneDrive throwing "Access to the cloud file is denied"

    param(
        [alias('LiteralPath')][string] $Path,
        [switch] $SkipFolder # when $Path is a folder, do not delete $path, only it's content
    )

    $local:ToDelete = @()

    if (Test-Path -LiteralPath $path) {
        foreach ($SinglePath in @(Get-Item -LiteralPath $Path)) {
            if (Test-Path -LiteralPath $SinglePath -PathType Container) {
                if (-not $SkipFolder) {
                    $local:ToDelete += @(Get-ChildItem -LiteralPath $SinglePath -Recurse -Force | Sort-Object -Culture $TemplateFilesSortCulture -Property PSIsContainer, @{expression = { $_.FullName.split('\').count }; descending = $true }, fullname)
                    $local:ToDelete += @(Get-Item -LiteralPath $SinglePath -Force)
                } else {
                    $local:ToDelete += @(Get-ChildItem -LiteralPath $SinglePath -Recurse -Force | Sort-Object -Culture $TemplateFilesSortCulture -Property PSIsContainer, @{expression = { $_.FullName.split('\').count }; descending = $true }, fullname)
                }
            } elseif (Test-Path -LiteralPath $SinglePath -PathType Leaf) {
                $local:ToDelete += (Get-Item -LiteralPath $SinglePath -Force)
            }
        }
    } else {
        # Item to delete does not exist, nothing to do
    }

    foreach ($SingleItemToDelete in $local:ToDelete) {
        try {
            Remove-Item $SingleItemToDelete.FullName -Force -Recurse
        } catch {
            Write-Verbose "Could not delete $($SingleItemToDelete.FullName), error: $($_.Exception.Message)"
            Write-Verbose $_
        }
    }
}


function ParseJwtToken {
    # Idea for this code: https://www.michev.info/blog/post/2140/decode-jwt-access-and-id-tokens-via-powershell

    [cmdletbinding()]
    param([Parameter(Mandatory = $true)][string]$token)

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
        $tokenHeader = [System.Text.Encoding]::ASCII.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json

        # Payload
        $tokenPayload = $token.Split('.')[1].Replace('-', '+').Replace('_', '/')

        # Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
        while ($tokenPayload.Length % 4) { $tokenPayload += '=' }

        # Convert to Byte array
        $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)

        # Convert to string array
        $tokenArray = [System.Text.Encoding]::ASCII.GetString($tokenByteArray)

        # Convert from JSON to PSObject
        $tokenPayload = $tokenArray | ConvertFrom-Json

        return @{
            error   = $false
            header  = $tokenHeader
            payload = $tokenPayload
        }
    }
}


#
# All functions have been defined above
# Initially executed code starts here
#

try {
    Write-Host
    Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    $ScriptPassedParameters = $MyInvocation.Line

    main
} catch {
    Write-Host
    Write-Host 'Unexpected error. Exit.' -ForegroundColor red
    $Error[0]
    exit 1
} finally {
    Write-Host
    Write-Host "Clean-up @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    # Restore original security setting
    if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
        Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction SilentlyContinue | Out-Null
    } else {
        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$($WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction SilentlyContinue | Out-Null
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

        $script:COMWord.Quit([ref]$false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWord) | Out-Null
        Remove-Variable -Name 'COMWord' -Scope 'script'
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

    if ($script:ScriptProcessPriorityOriginal) {
        $null = Get-CimInstance Win32_process -Filter "ProcessId = ""$PID""" | Invoke-CimMethod -Name SetPriority -Arguments @{Priority = $script:ScriptProcessPriorityOriginal }
    }

    Write-Host
    Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
}
