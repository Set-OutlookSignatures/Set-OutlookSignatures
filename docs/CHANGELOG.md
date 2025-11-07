<!-- omit in toc -->
## **<a href="https://set-outlooksignatures.com" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Set-OutlookSignatures%20Logo.png" width="250px" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a>**<br>Email signatures and out-of-office replies for Exchange and Outlook.<br>Full-featured, cost-effective, unsurpassed data privacy.<br><br><a href="https://set-outlooksignatures.com" target="_blank"><img src="https://img.shields.io/github/license/Set-OutlookSignatures/Set-OutlookSignatures?label=License&labelColor=black&color=informational" alt="License: EUPL 1.2"></a><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational?labelColor=black" alt="this release"></a> XXXRemoveWhenBuildingXXX--> <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/tag/Set-OutlookSignatures/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=Latest%20release&color=informational&labelColor=black" alt="Latest release" data-external="1"></a> <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/downloads/set-outlooksignatures/set-outlooksignatures/total?label=Downloads&labelColor=black" alt="Downloads" data-external="1"></a> <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/Set-OutlookSignatures/Set-OutlookSignatures?label=Issues&labelColor=black" alt="Issues" data-external="1"></a> <a href="https://set-outlooksignatures.com/faq/#44-what-can-i-learn-from-the-code-of-set-outlooksignatures" target="_blank"><img src="https://img.shields.io/badge/Behind%20the%20scenes-Learn%20from%20the%20code-lawngreen?labelColor=black" alt="Behind the scenes: Learn from the code"></a>

# Changelog
<!--
Sample changelog entry
Update .\docs\releases.txt

## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/vX.X.X" target="_blank">vX.X.X</a> - YYYY-MM-DD
_Put Notice here_
_**Breaking:** <Present tense verb> XXX_  
### Changed
- **Breaking:** <Present tense verb> XXX
- <Active present tense verb> XXX
### Added
- <Active present tense verb> XXX
### Removed
- <Active present tense verb> XXX
### Fixed
- <Active present tense verb> XXX
-->


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.23.0" target="_blank">v4.23.0</a> - 2025-10-18
### Changed
- Update dependencies
  - @azure/msal-browser v4.25.1
  - HtmlAgilityPack v1.12.4
  - MSAL.Net v4.77.1
  - QRCoder v1.7.0
- Reduce output of Outlook add-in taskpane when not being viewed in Outlook.
- Change included sample QR code from MeCard to the more frequently used vCard format.
- Update sample templates to reflect QR code and phone number changes as described in the '`Added`' section.
### Added
- Add check for PowerShell Full Language Mode to all admin scripts, including sample code.
- Add a phone number formatter based on Googles libphonenumber library.<br>The code '`FormatPhoneNumber`' in '`.\config\default replacement variables.ps1`' is used to create the new additional default replacement variables '`$Current[user|Manager|Mailbox|MailboxManager][Telephone|Fax|Mobile]-[E164|INTERNATIONAL|NATIONAL|RFC3966]$`'. The sample code takes a phone number and optionally a country, and converts them to standardized predefined formats for easy human readability or for use in '`tel:`' links. A custom formatting option is also available.
- Add an address formatter.<br>The AddressFormatter module is based on the great work OpenCage GmbH shares on <a href="https://github.com/opencagedata/address-formatting">their GitHub repo</a>.<br>The code in '`.\config\default replacement variables.ps1`' used to create the new '`$Current[user|Manager|Mailbox|MailboxManager]PostalAddress$`' replacement variables can easily be adapted to create other postal addresses matching the formatting rules of the defined for a user or mailbox.
- Add FAQ '[What can I learn from the code of Set-OutlookSignatures?](https://set-outlooksignatures.com/faq/#44-what-can-i-learn-from-the-code-of-set-outlooksignatures)'. It gives an overview which scripting techniques you can learn from Set-OutlookSignatures.
- Add '`.\sample code\Start-IfADAvailable.ps1`', showing how to run Set-OutlookSignatures (and other software) only when a working connection to Active Directory is available.
- Add a workaround for Windows PowerShell 5.1 sometimes not being able to access variables with a 'script' scope and a name of 'IsWindows', 'IsLinux', or 'IsMacOS'.
### Removed
### Fixed
- Fix the Sort-Object parameter to use -Culture in the Intune remediation sample script, restoring culture-aware sorting. (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/pull/145" target="_blank">#145</a>) (Thanks <a href="https://github.com/rene-schwabe" target="_blank">@rene-schwabe</a>!)
- Include relative path information when downloading files from SharePoint Online. This makes sure that images in subfolders are correctly downloaded when using HTM templates.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.22.0" target="_blank">v4.22.0</a> - 2025-09-05
### Changed
- Update Outlook add-in dependency @azure-msalbrowser to v4.22.1.
- Change Outlook add-in variable names for clarity:
  - Use '`customRulesResultSignatureName`' instead of '`customRulesPropertiesResult`'. The old variable name can still be used but the new one overrides it.
- Force the Outlook add-in to use the Graph API when the mailbox is hosted in Exchange Online, as Microsoft will be turning off EWS for Exchange Online soon and as the Outlook add-in supports the Graph API since it has been released over a year ago.
- Optimize the performance accessing SharePoint Online via Graph API by using different endpoints and exactly defining the requested attributes. This greatly enhances performance when the content used for Set-OutlookSignatures is not hosted in a separate document library, but in a document library with many other unrelated items.
### Added
- Allow Outlook add-in custom rules code not only to choose from a set of signatures previously deployed using Set-OutlookSignatures, but to set a completely customized signature using the '`customRulesResultSignatureBody`' property. This makes the Outlook add-in the most flexible signature add-in on the market as known of today.
  - The '`customRulesResultSignatureBody`' proprty is only available for mailboxes hosted in Exchange Online and requires setting two additional permissions in the Entra ID app it uses.
  - See '`.\sample code\CustomRulesCode.js`' for details about this new feature
- Add the following additional new features to Outlook add-in custom rules (see '`.\sample code\CustomRulesCode.js`' in the Outlook add-in folder for details):
  - The '`customRulesProperties.notificationCurrent`' property contains the current text used for the Outlook notification shown after the signature has been set, '`customRulesProperties.notificationOriginal`' contains the value defined originally for the add-in.
  - Define the text used for the Outlook notification shown after the signature has been set by setting the new '`customRulesResultNotification`' variable.
  - Add boolean properties '`customRulesProperties.itemIsReply`' and '`customRulesProperties.itemIsForward`'.
  - Add boolean properties '`customRulesProperties.itemIsHtml`' and '`customRulesProperties.itemIsText`'.
  - Add string properties '`graphAccessToken`' and '`graphApiEndpoint`', which can be used to query the Graph API for use in custom rules.
  - Add sample code showing how to query the Graph API (convert sender SMTP email address to UPN, get basic properties of a user, get transitive group membership of a user).
- Allow to define a custom ID for the Outlook add-in. This allows using multiple independent configurations and versions of the Outlook add-in in the same environment.
- Add a workaround in the Outlook add-in for the Microsoft problem with Outlook Web and New Outlook for Windows not replacing signatures but adding them in appointments, resulting in multiple signatures.
- Add comments to clarify mailbox requirements and permissions in the '`manifest.xml`' file of the Outlook add-in.
- Consider the names of the default signatures defined in Exchange Online when Set-OutlookSignatures runs the Benefactor Circle add-on specific preparations for the Outlook add-in and when the registry does not contain any information about default signatures. This is helpful in scenarios in which Set-OutlookSignatures is used to deploy signatures, but not to define a default signatures.
- Update '`.\sample code\Create-EntraApp.ps1`' to include the additional Graph scopes needed for the Outlook add-in feature '`customRulesResultSignatureBody`'.
- Add additional error handling to the Outlook add-in file '`run_before_deployment.ps1`'.
### Removed
### Fixed
- Make sure the Outlook add-in correctly reports '`customRulesProperties.itemIsNew`' and '`customRulesProperties.itemIsReplyForward`' on all platforms and for all item save and sync states when using custom rules code.
- Make sure the Outlook add-in does not re-use '`customRulesPropertiesResult`' between runs, no matter if triggered manually via the taskpane or automatically via a launch event.
- Add an additional Graph scope to '`.\sample code\Create-EntraApp.ps1`' to avoid problems with granting admin consent for delegated permissions. (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/discussions/143" target="_blank">#143</a>) (Thanks <a href="https://github.com/wwa-jvteeffelen" target="_blank">@wwa-jvteeffelen</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.21.0" target="_blank">v4.21.0</a> - 2025-08-15
### Changed
- Update dependency MSAL.Net to v4.75.0.
- Update Outlook add-in dependency @azure-msalbrowser to v4.20.0.
### Added
- Add the following new features to the Outlook add-in:
  - React to launch events '`OnMessageRecipientsChanged`' and '`OnAppointmentAttendeesChanged`'.
  - Allow adding custom code to the add-in so you can directly influence which signature it will set.  
  For example, you can set a specific signature…
    - …when there are only internal recipients, or another signature when there are external recipients
    - …depending on the from email address
    - …when a specific customer is in the To field
    - …when the current item is a mail or an appointment
    - …when the current item is a new mail, or another signature when it is a reply or a forward
    - …depending on the subject
    - …or any other condition derived from the information available in the customRulesProperties object  

    See '`.\sample code\CustomRulesCode.js`' in the Outlook add-in folder for details.
### Removed
### Fixed
- Correct a problem reading email addresses from New Outlook for Windows. (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/140" target="_blank">#140</a>) (Thanks <a href="https://github.com/psic4t" target="_blank">@psic4t</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.20.4" target="_blank">v4.20.4</a> - 2025-08-12
### Changed
- Update dependency MSAL.Net to v4.74.1.
- Update dependency HtmlAgilityPack to v1.12.2.
- Update dependency UTF.Unknown to v2.6.0.
- Update Outlook add-in dependency @azure-msalbrowser to v4.19.0.
### Added
- Add workaround for Microsoft internal problems with roaming signatures API, which would result in the following error when uploading roaming signatures to mailboxes which are not the one of the logged-on user (especially in SimulateAndDeploy mode). The log would contain the error "The response status code for https://localhost:444/owa/service.svc?action=SanitizeHtml action \u0027SanitizeHtml\u0027 is InternalServerError. ReasonPhrase:Internal Server Error".
- Add hints about probable root causes of empty signatures lists in the log output of the Outlook add-in.
### Removed
### Fixed


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.20.3" target="_blank">v4.20.3</a> - 2025-07-17
### Changed
- Update dependency MSAL.Net to v4.74.0.
- Update Outlook add-in dependency @azure-msalbrowser to v4.15.0.
### Added
### Removed
### Fixed
- Add a workaround for PowerShell 5.1 not correctly converting arrays to JSON when passing them via the InputObject parameter. (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/136" target="_blank">#136</a>)
- Fix Set-OutlookSignatures quitting with exit code 14 when using it in SimulateAndDeploy mode for Exchange Online.
- Add a workaround for PowerShell not reliably resolving SMB paths with relative parts to absolute paths in the MSAL.PS module.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.20.2" target="_blank">v4.20.2</a> - 2025-07-07

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update dependency PreMailer.Net to v2.7.2.
- Update Outlook add-in dependency @azure-msalbrowser to v4.14.0.
### Added
### Removed
### Fixed
- Update handling of the '`@`' character in signatures names. Microsoft no longer allows '`@`' in roaming signature names as it is reserved for internal use, and Microsoft Outlook will follow soon.
  - Add '`@`' to the list of invalid characters for signatures names as documented in sample INI files. The new complete list of invalid characters is '`\/:"*?><,|@`'.
  - Convert all invalid characters to '`_`', but '`@`' to '`_at_`'.
  - Only keep the '`@`' character where it is required for Microsoft internal use in roaming signature names.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.20.1" target="_blank">v4.20.1</a> - 2025-06-26

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update dependency MSAL.Net to v4.73.1.
- Update Outlook add-in dependency @azure-msalbrowser to v4.13.2.
### Added
### Removed
### Fixed
- Fix a logical error in the code for the '`DeleteScriptCreatedSignaturesWithoutTemplate'` parameter which made the cleanup stop before any work was done. This error only occurred when the parameter was enabled, which it is by default.
- Fix a logical error in the code for the '`DeleteUserCreatedSignatures`' parameter which deleted all signatures instead of only user-created ones. This error only occurred when the parameters was enabled, which it is not by default.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.20.0" target="_blank">v4.20.0</a> - 2025-06-18

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update dependency MSAL.Net to v4.73.0.
- Update dependency PreMailer.Net to v2.7.0.
- Update Outlook add-in dependency @azure-msalbrowser to v4.13.1.
- Update dependency QRCoder to v1.6.0.
- Switch from dependencies AngleSharp and HTMLFile to HtmlAgilityPack v1.12.1.
- Remove '`Set-Location -LiteralPath $PSScriptRoot`' from scripts that do not require to know from which path they are run from.
- Reduce Entra ID throttling probability in '`.\sample code\SimulateAndDeploy.ps1`' by applying the following changes:
  - Do not refresh the Graph token every time a new job is started, but only when:
    - Token is expired.
    - The remaining token lifetime is less than or equal the job timeout.
    - At least half of the token lifetime has already passed.
  - In case of an error, wait 70 seconds between the three retries.
- Show all replacement variables in verbose output, not just the ones that have values. This helps detect null or empty values.
- Rewrite and radically simplify the [Quick Start Guide](https://set-outlooksignatures.com/quickstart).
- Reduce repository and release size by deleting unused image files and by outsourcing the detailed documentation to [set-outlooksignatures.com](https://set-outlooksignatures.com).
### Added
- Support for cross-tenant access and Multitenant Organizations. This allows to deploy signatures for mailboxes that are not hosted in the users home tenant, with all the properties, replacement variables and Benefactor Circle features being fully available. See the description of the parameter '`GraphClientID`' in the [parameter documentation](https://set-outlooksignatures.com/parameters) for details.
- Detect the encoding of HTML template files automatically, using the same logic as browsers do. See the FAQ '`Should I use .docx or .htm as file format for templates?`' for details.
- Convert HTML files to UTF8 when uploading or downloading roaming signatures, with the same logic browsers use. This is another fix for the problems of Outlook's own signature sync mechanism, and also solves issues with signatures that have been created manually in Non-UTF-8 encoding.
- Check if the automatic variable '`$PSScriptRoot`' is available when scripts require to know from which path they are run from. This avoids follow-up problems when scripts are not run as file, but as selection of text in code editors.
- Delete local copies of roaming signatures when cleaning up if they have previously been downloaded from other mailboxes (not the mailbox of the current user) and if these mailboxes are no longer available in Outlook.
- Allow defining the text of the notification banner that is shown after a signature has been added to an email or an appointment by the Outlook add-in. The parameter is called '`NOTIFICATION_TEXT`' and can be passed to '`run_before_deployment.ps1`'.
- Add '`Does it support cross-tenant access and Multitenant Organizations?`' to the [FAQ section](https://set-outlooksignatures.com/faq).
### Removed
### Fixed
- Switch from '`Set-Location $PSScriptRoot`' to '`Set-Location -LiteralPath $PSScriptRoot`' in scripts that do require to know from which path they are run from. This allows the use of paths containing characters with a specific meaning in PowerShell (square brackets, for example).
- Fix a typo in source code of FAQ '`How can I start the software only when there is a connection to the Active Directory on-prem?`' in the [FAQ section](https://set-outlooksignatures.com/faq), and add checks for the forest root domain and it's child domains.
- Make sure that DOCX template files are not opened in "Read" view in Microsoft Word, as search and replace is not available in this view.
- Make sure that temporary copies of template files are not write protected and do not have the "mark of the web", as both may lead to problems with search and replace or file type conversion in Microsoft Word.
- Add support for Outlook profile names with special characters, which Classic Outlook for Windows saves in the registry using a completely unknown and not publicly documented encoding. Set-OutlookSignatures can not display the special characters correctly due to the unknown encoding, but can now read the contents without errors.
- Make sure that leading whitespace in signature lines is correctly represented in the email draft containing all signatures (parameter '`SignatureCollectionInDrafts`' and signature preview in the taskpane of the Outlook add-in).
- Correctly handle HTML template files referencing the same image file multiple times but with different image replacement variables such as '`$CurrentUserPhoto$`': Do not overwrite image files which are referenced multiple times, create a new GUID named file instead.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.19.0" target="_blank">v4.19.0</a> - 2025-04-16

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update dependency MSAL.Net to v4.70.2.
- Update Outlook add-in dependency @azure/msal-browser to v4.11.0.
- Make sure that all kinds of sort operations use the invariant culture 127 if documentation does not contain a specific sort order culture.
- Update '`How can I start the software only when there is a connection to the Active Directory on-prem?`' in the [FAQ section](https://set-outlooksignatures.com/faq).
- Format code blocks in documentation files using syntax highlighting where possible (GitHub only).
### Added
- Add Option '`CurrentUserOnly`' to parameter '`MirrorCloudSignatures`'. With this setting, only roaming signatures from the mailbox of the current user are downloaded, not from other mailboxes the user has full acccess to. This option does not have any effect on uploading roaming signatures.
- Add support for the taskpane of the Outlook add-in to be shown in read mode for messages. This makes it easier to check if the add-in is deployed correctly, and if it can access signatures. This is especially useful on mobile devices, in situations where enabling the debug mode is not wanted, and for basic tests when launch events are not triggered by Outlook.
- Add support for parameters to '`run_before_deployment.ps1`' of the Outlook add-in. This makes updates easier as you no longer have to modify the script itself, but can call it with parameters to set options specific for your environment.
- Check for missing dependencies before using authentication broker authentication on Linux.
- Add an option to '`.\sample code\Create-EntraApp.ps1`' that allows to create the Entra ID app required for the Outlook add-in that comes with the Benefactor Circle add-on.
- Add code to '`.\sample code\Create-EntraApp.ps1`' that automatically grants admin consent for a newly created Entra ID app.
- Support the '`ScriptProcessPriority`' parameter on Linux and macOS.
- Allow defining the HTML code of templates at script runtime using replacement variables. This is enabled by replacing non-picture variables directly in the source code of HTM template files instead of their parsed DOM OuterHtml.
### Removed
### Fixed
- Change command invocation logic in '`.\sample code\SimulateAndDeploy.ps1`' from a custom method to native splatting as this makes it easier to correctly handle strings, arrays, and parameters expecting boolean values. If you did not copy and modify '`.\sample code\SimulateAndDeploy.ps1`' but call it directly with parameters, you need to remove additional quotes in string parameter values.
- Fix an issue in '`.\sample code\Create-EntraApp.ps1`' that could lead to missing permissions when creating a new app for SimulateAndDeploy in certain situations.
- Make sure that '`run_before_deployment.ps1`' (part of the Outlook add-in) correctly handles empty arrays.
- Implement workarounds to reduce possible error sources when replacing variables in DOCX templates. These workarounds work in most cases, but not in all. To be on the safe side, make sure to use Word version 2502 (Build 18526.20144) or newer, or lower than 2408 (Build 17928.20468).
- Fix a typo in the default value for '`OOFIniFile`' on Linux and macOS, which leads to not finding the sample OOF INI file due to case sensitivity.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.18.3" target="_blank">v4.18.3</a> - 2025-04-03

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update Outlook add-in dependency @azure/msal-browser to v4.9.1.
- Make sure that signatures are available for the Outlook add-in even while "Prepare data for Outlook add-in" is running. No matter how long this preparation task takes, the switch from old to new signatures now only takes a second for the add-in.
### Added
### Removed
### Fixed
- Include line breaks when showing the signature preview for non-HTML signatures in the taskpane of the Outlook add-in.
- Do not detect the default console color, as this fails on many Linux terminals.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.18.2" target="_blank">v4.18.2</a> - 2025-03-19

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update dependency MSAL.Net to v4.70.0.
- Simplify the ['The Outlook Add-in'](https://set-outlooksignatures.com/outlookaddin) chapter and include the latest Microsoft clarifications for their Outlook APIs.
### Added
### Removed
### Fixed
- Add a workaround to a bug in the Microsoft API for Classic Outlook for Windows (and probably others), not fully garbage collecting variables. This bug results in signatures not updating when the OnMessageFromChanged or OnAppointmentFromChanged launch event is triggered. This bug may also lead to the launch events not being triggered at all, only Microsoft can solve this.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.18.1" target="_blank">v4.18.1</a> - 2025-03-14

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update Outlook add-in dependency @azure/msal-browser to v4.7.0.
### Added
### Removed
### Fixed
- Fix handling of existing signature files in Classic Outlook for Windows when they are write protected or hidden.
- Make the code deciding if a SharePoint Online element is a file or a folder work correctly even when the Graph API erronously returns the contentType in a non-English language.
- Fix a race condition in the Outlook add-in that may lead to an empty signature list in the taskpane when legacy Exchange Online tokens have been disabled (see [here](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/faq-nested-app-auth-outlook-legacy-tokens) for details).


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.18.0" target="_blank">v4.18.0</a> - 2025-03-06

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update Outlook add-in dependency @azure/msal-browser to v4.5.1.
- Update dependency MSAL.Net to v4.69.1.
- Copy MSAL.Net runtimes only for the current operating system.
- Reduce the timeout value for autodiscover via Exchange Web Services from 100 to 25 seconds, which matches the Outlook default value.
### Added
- Convert more code from Exchange Web Services (EWS) to Graph, as new APIs have been released. To be able to use this, add the '`Mail.ReadWrite`' permission to your Entra ID app as documented in '`.\config\default graph config.psq`', '`.\sample code\Create-EntraApp.ps1`' and '`.\sample code\SimulateAndDeploy.ps1`': Delegated for Set-OutlookSignatures, delegated and application for SimulateAndDeploy. If you do not do this, Set-OutlookSignatures will fall back to using EWS, but this is only considered a workaround from now on. There is no need to change anything for mailboxes hosted on-prem.
  - Adding '`Mail.ReadWrite`' actually does not add a permission, as they are already included in the existing '`EWS.AccessAsUser.All`' permission, but '`Mail.ReadWrite`' is required when doing things with Graph instead of EWS. Do not remove '`EWS.AccessAsUser.All`', as it is still needed for some tasks that cannot be done via Graph.
  - Also see the new FAQ '`What about Microsoft turning off Exchange Web Services for Exchange Online?`' in the [FAQ section](https://set-outlooksignatures.com/faq) to learn how Set-OutlookSignatures will be affected when Microsoft turns off EWS for Exchange Online.
- Remove empty CSS properties and resolve multiple assignments in style attributes when creating HTML files. This corrects several errors in Word which affect Outlook (which uses Word as HTML renderer): Using the no longer supported text-autospace CSS property in HTML exports, setting an invalid null value for it, and not interpreting the null value as default 'none' value when rendering.
- Work around a problem in PowerShell 7 showing wrong number of files in progress bars for copy and delete operations, and not removing these progress bars from screen.
- Show a warning when Graph authentication is using the Entra ID app provided by the developers of Set-OutlookSignatures. It is recommended to create and use your own Entra ID app.
- Detect when Set-OutlookSignatures has already been run in the current PowerShell session and exit when this is the case. This is the only way to prevent problems caused by .Net caching DLL files in memory.
- Require a key press to continue when loading the Benefactor Circle license file results in a warning. Set-OutlookSignatures continues automatically after 120 seconds.
- Add authentication broker support for Linux. No new redirect URI for the Entra ID app is required.
- Add new FAQ '`Roaming signatures in Classic Outlook for Windows look different`' to the [FAQ section](https://set-outlooksignatures.com/faq).
### Removed
- Remove comment-based help content from Set-OutlookSignatures.ps1 and refer to [the online help](https://set-outlooksignatures.com/help) instead.
- Remove testing the connectivity to endpoints during Graph authentication, as Microsoft's servers suddenly randomly incorrectly interpret the HTTP 'Expect' header, returning HTTP error status code 417 instead of a 2xx success code. Rely on integrated tests in Microsoft's authentication library instead. (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/133" target="_blank">#133</a>)
### Fixed
- Correct the Entra ID app configuration requirement for broker authentication in '`.\config\default graph config.ps1`'. Use the application ID, not the Object ID.
- Fix a race condition when evaluating OOF INI files. OOF INI entries without both the tags '`Internal`' and '`External`' now always correctly default to both, and not sometimes to only '`Internal`'.
- Write all signature HTML files without a byte order mark (BOM) for consistency between PowerShell 5 and PowerShell 7+.
- Correctly replace the connected folder path for images in signatures whose name starts with a number. This is a workaround for a problem in PowerShell's implementation of regular expressions.
- Fix a bug where Active Directory is queried for mailbox properties instead of Graph. This leads to multiple follow-up errors, including the license group membership check in the Benefactor Circle add-on. The error occurs when all of the following conditions are met: The user has a connection to the on-prem Active Directory, GraphOnly is not enabled, and a Graph connection must be enforced as other parameters require it (the conditions are logged in the script output).
- Fix the bug which made the Outlook add-in script '`run_before_deployment.ps1`' work in PowerShell 7.4 and newer only, but not in Windows PowerShell.
- Fix the bug that simulation mode did process out-of-office templates but did not include the results in the output.
- Fix the bug in '`SimulateAndDeploy.ps1`' that only running jobs instead of all jobs were considered for timeout cancellation.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.17.0" target="_blank">v4.17.0</a> - 2025-02-14

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Enable '`Ignore host and platform`' per default when using the taskpane of the Outlook add-in.
- Update dependency MSAL.Net to v4.68.0.
- Update Outlook add-in dependency @azure/msal-browser to v4.2.1.
### Added
- Add the new silent authentication mode '`Silent via Authentication Broker without login hint`'. This allows for silent login to Graph in environments where Integrated Windows Authentication does not work. See the '`Authentication`' chapter in the [technical details documentation](https://set-outlooksignatures.com/details) for details.
- Connecting to Exchange Online now goes through the same authentication steps as connecting to Graph.
- Add virtual mailboxes and dynamically create signature and out-of-office INI lines through code
  - Add parameter '`VirtualMailboxConfigFile`'. Virtual mailboxes are mailboxes that are not available in Outlook but are treated by Set-OutlookSignatures as if they were. This is an option for scenarios where you want to deploy signatures and out-of-office replies with not only the '`$CurrentUser...$`' but also '`$CurrentMailbox...$`' replacement variables for mailboxes that have not been added to Outlook, such as in Send As or Send On Behalf scenarios, where users often only change the from address but do not add the mailbox to Outlook. See the [parameters documentation](https://set-outlooksignatures.com/parameters) for details about the new parameter and sample code. This feature requires a Benefactor Circle license. It is best used together with [Export-RecipientPermissions](https://github.com/Export-RecipientPermissions).
  - Add a way to dynamically define signature INI and out-of-office INI file entries via the '`VirtualMailboxConfigFile`' parameter. This enables multiple scenarios for advanced usage - for example you can automate INI file entries for delegate scenarios in combination with [Export-RecipientPermissions](https://github.com/Export-RecipientPermissions). See the [parameters documentation](https://set-outlooksignatures.com/parameters) for details about the new parameter and sample code.
- Add a signature preview to the taskpane of the Outlook add-in.
- Add info to '`Authentication`' chapter in the [technical details documentation](https://set-outlooksignatures.com/details):
  - Integrated Windows Authentication only works for federated users in domains with an authentication type of "federated".
  - Description of 'Silent via Authentication Broker without login hint'.
- Add more workarounds for WebDAV client or .Net returning file and folder names with trailing NULL characters as well as other invalid file and folder names.
- Support SharePoint paths where the actual name of the document library is not reflected in the URL. This typically happens when renaming a document library.
- Allow setting '`$VersionToUse`' in '`.\sample code\Intune-setOutlookSignatures-Remediate.ps1`' to an empty value, so that no download occurs and you can use a Set-OutlookSignatures instance previously deployed with an other mechanism.
- Add to '`.\sample code\SimulateAndDeploy.ps1`': The verbose mode setting of SimulateAndDeploy.ps1 is now passed on to Set-OutlookSignatures.ps1 per default.
- Make signatures more compatible with Non-Outlook clients when using DOCX templates.
- Add new FAQ '`Empty lines contain an underlined space character`' to the [FAQ section](https://set-outlooksignatures.com/faq).
- Show exit code and, if available, the specific exit code description in script output. The exit code is 0 for success and a number between 1 and 255 in case of an error.
- Show the source for each security identifier in verbose output. This helps troubleshooting as it shows to which type of group a security identifier belongs to. Values for On-prem: Global or universal group, domain local group, static distribution group, domain local group across trust (including a hint to sIDHistory where possible). Values for cloud: Entra ID group, on-prem group.
- Include onPremisesSecurityIdentifier SIDs when getting group membership for mailboxes.
- Add support for .Net 9 and PowerShell 7.5.0, which is based on .Net 9.
### Fixed
- Check INI file paths for the correct type ('Leaf' instead of 'File'). This is a pure optical change.
- Update '`$SoftwarePath`' in '`.\sample code\Intune-SetoutlookSignatures-Remediate.ps1`', so that the resulting folder structure is easier to understand. This is a pure optical change.
- Avoid race conditions that can lead to a .Net Framework pop-up error ('`The pipeline has been stopped`') when detecting termination signals that allow for a graceful exit.
- Fixed '`.\sample code\Intune-setOutlookSignatures-Remediate.ps1`' trying to stop a non-existing transcript.
- Fix '`DeleteScriptCreatedSignaturesWithoutTemplate`' not deleting outdated signatures if they were downloaded from Exchange Online.
- Detect default Linux keyring instead of relying on a fixed name when checking availability and lock state of keyring.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.16.1" target="_blank">v4.16.1</a> - 2024-12-05
_**Attention, Exchange Online admins**_  
_See '`What about the roaming signatures feature in Exchange Online?`' in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works.<br>Set-OutlookSignatures supports cloud roaming signatures - see the [parameters documentation](https://set-outlooksignatures.com/parameters)._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Fixed
- Fix '`.\sample code\Create-EntraApp.ps1`' so that the redirect URI for broker authentication contains the Application ID and not the Object ID of the newly created app. When you have already used the sample code, make sure to manually change the redirect URI of the Entra ID app from the Object ID to the Application ID.
- Correct the text color in the sample templates to match the latest accessibility recommendations for contrast.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.16.0" target="_blank">v4.16.0</a> - 2024-12-02
_**Attention, Exchange Online admins**_  
_See '`What about the roaming signatures feature in Exchange Online?`' in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works.<br>Set-OutlookSignatures supports cloud roaming signatures - see the [parameters documentation](https://set-outlooksignatures.com/parameters)._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- **Prefer an authentication broker over browser-based authentication (browser auth is still used as fallback and on non-supported systems). This helps overcome issues with Entra ID MFA re-authentication as well as browser authentication problems such as being denied access to http://localhost. Make sure to add the Redirect URI '`ms-appx-web://microsoft.aad.brokerplugin/<Application ID of your app>`' to your Set-OutlookSignatures Entra ID app.** Make sure to use the Application ID and not the Object ID. The Entra ID app provided by the developers already has the additional Redirect URI set ('`ms-appx-web://microsoft.aad.brokerplugin/beea8249-8c98-4c76-92f6-ce3c468a61e6`').
- Change the path of the Graph token cache file to '`$(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Set-OutlookSignatures\MSAL.PS\MSAL.PS.msalcache.bin3')`' on all platforms. This change requires one-time re-authentication towards Graph on Windows, Linux and macOS when Integrated Windows Authentication does not work. The change is introduced to fix the following problems and to anticipate upcoming changes across all supported platforms:
  - On Windows, not only Set-OutlookSignatures uses the default MSAL.Net/MSAL.PS cache file path. This is a good idea but most software handles the cache as if it was application specific, replacing all other tokens with their own instead of sharing them.
  - On Linux (and macOS), MSAL.Net does not rely on .Net to determine the path for LocalApplicationData but uses own logic, which leads to inconsistent results on different Linux distributions and does not always match XDG specifications.
  - On macOS, .Net 8 returns a different path for LocalApplicationData than earlier versions, requiring a change anyhow.
- Simplify the taskpane user interface of the Outlook add-in
  - Show only default actions when opening the taskpane: A big button to set the signature, and a dropdown list to override the automatically chosen signature with a manual selection.
  - Show advanced options when scrolling down: Choosing a debug mode, an option to ignore host and platform, and a new textbox containing the log output of the add-in.
- Format plain text signatures in the system's monospace font when the '`-SignatureCollectionInDrafts true`' parameter is used.
- Speed up connecting to SharePoint Online paths by directly trying to access them via Graph when GraphClientID is available, instead of waiting for the Test-Path timeout.
- Change the encoding of all '.ps1' files which are intended to be executed directly by a user from UTF-8 to UTF-8 with BOM. This ensures that Unicode characters are correctly written to the console not only on PowerShell 7+, but also on Windows PowerShell. This is a pure optical fix without any functional changes and has no impact on the encoding of custom configuration and template files.
- Decode the Graph token as UTF-8 instead of ASCII to ensure that verbose output of the token correctly displays Unicode characters. This is a pure optical change without any functional impact.
- Update the sample code for Intune detect and remediate scripts to use the log file of the new logging feature of Set-OutlookSignatures.
- Update the README chapter '`Group membership`' to better point out that group membership includes transitive/nested/indirect membership.
- Update the sample code in FAQ '`How can I start the software only when there is a connection to the Active Directory on-prem?`'.
- Update Outlook add-in dependency @azure/msal-browser to v3.27.0.
- Update dependency MSAL.Net to v4.66.2.
### Added
- Log every run of Set-OutlookSignatures. Logs are saved in the folder '`$(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Set-OutlookSignatures\Logs')`', the files follow the naming scheme '`$("Set-OutlookSignatures_Log_yyyyMMddTHHmmssffff.txt")`', and files older than 14 days are deleted with every run.
- Check connectivity to the Graph authentication endpoint before trying to access Graph. This not only catches connection errors as early as possible, but also avoids prompting users with authentication pop-ups in scenarios where they are offline or access to Graph is blocked at the firewall or proxy level.
- Allow using Active Directory DNS domain names when assigning templates to groups in INI files. As Microsoft Graph started exposing the corresponding attributes for groups, the DNS domain name format now not only works on-prem but also in hybrid environments. This means that there is no longer anything in Set-OutlookSignatures for which a NetBIOS domain name is mandatory.
- Allow assignment of templates to Entra ID groups by their Object ID and their securityIdentifier, in addition to existing properties such as email address, mailNickname and displayName.
- Display hints if the search for the properties of a mailbox via Graph returns nothing or more than one result, analogous to the case where the same search in Active Directory on-prem fails.
- Reduce the number of Graph queries required to find the security identifier of a group defined in a template tag. This is done by a simple check of domain name and group name format, and avoiding queries which would not work anyhow (for example, querying Graph for a NetBIOS domain name when the given format is a DNS domain name).
- Prevent the system from going to sleep right before the first template file is opened, and allow sleep again as soon as the process running Set-OutlookSignatures ends.
- Detect termination signals that allow for a graceful exit, and start clean-up tasks as soon as possible. Around 500 exit points are defined in the code. On Linux and macOS, signals such as SIGINT, SIGTERM, SIGQUIT and SIGHUP are recognized; on Windows, an imminent logout, shutdown or restart is also recognized.
- Added sample code showing how to selectively enable verbose logging for specific users or computers to FAQ '`How can I log the software output?`' in the [FAQ section](https://set-outlooksignatures.com/faq).
- Animate the company logo in signature sample templates.
- Add a workaround that avoids errors when using DOCX templates with older versions of Word detecting that they are not the default program for all the file types they feel responsible for. Word displays an information dialog, the error message is '`Call was rejected by callee, 0x80010001 (RPC_E_CALL_REJECTED)`'.
- Add information to the '`Authentication`' chapter in the [technical details documentation](https://set-outlooksignatures.com/details).
- Add support for unblocking files on macOS using the PowerShell cmdlet 'Unblock-File'. Unblock-File is supported on Windows and macOS.
### Removed
- Removed attributes from '`$GraphUserProperties`' array in '`.\config\default graph config.ps1`' which were not used for default replacement variables. There should not be any side effect for you custom Graph configuration. (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/130" target="_blank">#130</a>)
- Remove the QR code from signature sample templates. The feature to create QR codes according to your specifications is unchanged, as you can see in sample template 'Test all default replacement variables'.
### Fixed
- Fix a problem where roaming signatures downloaded from additional and automapped mailboxes overwrite identically signatures which were created by Set-OutlookSignatures for a mailbox with higher priority.
- Correctly handle paths to outlook.exe and winword.exe stored in the registry in enclosing quotes.
- Update login hint detection logic for Graph authentication to handle cases where the UPN used to log on does not match the UPN in Entra ID.
- Add more code to work around limitations for Outlook add-ins on Outlook Web on premises. Microsoft has silently removed on-prem support for several features in office.js in the last few weeks - if this goes on, on-prem signature add-ins will soon not be realizable any more.
- Fix the workaround to not show Word security warning when converting documents with linked images when using '`-CreateRTFSignatures true`'.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.15.0" target="_blank">v4.15.0</a> - 2024-09-27
_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works.<br>Set-OutlookSignatures supports cloud roaming signatures - see `MirrorCloudSignatures` in the [parameters documentation](https://set-outlooksignatures.com/parameters)._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- **Change the default value of the '`MirrorCloudSignatures`' parameter to 'true'.**<br>Although the Microsoft API is still not publicly available, it has been stable for more than two years and is being used by more and more Outlook editions.<br>This change only affects a small number of installations, as MirrorCloudSignatures is a Benefactor Circle exclusive feature that practically all clients with cloud mailboxes activate.
- Try Integrated Windows Authentication first when connecting to Exchange, even when a cloud access token is available.<br>This is slower but ensures maximum compatibility for all architectures and combinations of access from internal or external networks, mailbox in cloud or on-prem, hybrid mode enabled (classic minimal/express/full, modern minimal/full) or disabled, Hybrid Modern Authentication enabled or disabled, and typical configuration issues.
- Change sample templates font to Aptos, the new default font used by Microsoft.
- Update dependency MSAL.Net to v4.65.0.
- Update Outlook add-in dependency @azure/msal-browser to v3.24.0.
### Added
- Allow to select signature in the taskpane of the Outlook add-in. This is like having roaming signatures on-prem.
- Make data preparation for Outlook add-in compatible with on-prem mailboxes. This does not (yet) remove the limitation that the add-in only works with cloud mailboxes on Outlook for Android and Outlook for iOS.
- Add Outlook add-in support for images in signatures in Outlook Web on premises (will start working as soon as Microsoft fixes a bug in their office.js framework) 
- Connect to Graph if only one Benefactor Circle license group is defined and this license group is an Entra ID group, and show a warning when the Benefactor Circle license group for a mailbox is an Entra ID group but there is no connection to Graph.
- Add support for the new way of Exchange Online modifying the HTML code of roaming signatures to '`DeleteScriptCreatedSignaturesWithoutTemplate`' and '`DeleteuserCreatedSignatures`'.
- Add support for the `OnAppointmentFromChanged event` in the `Outlook add-in`, as Microsoft has updated the office.js library accordingly.
- Adapt the conversion of DOCX templates to HTML so that the different Outlook editions on different platforms render fonts and colors more consistently.
- Show more detailed troubleshooting hints when Autodiscover fails.
- Show if a mailbox is in a license group before mailbox specific Benefactor Circle features are run.
### Removed
### Fixed
- Fix the problem in the Outlook add-in that led to images not being shown in the signature on the first run of the add-in for a new appointment.
- Add more code to work around limitations for Outlook add-ins on Outlook Web on premises.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.14.2" target="_blank">v4.14.2</a> - 2024-09-07
_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works.<br>Set-OutlookSignatures supports cloud roaming signatures - see `MirrorCloudSignatures` in the [parameters documentation](https://set-outlooksignatures.com/parameters)._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update Outlook add-in dependency @azure/msal-browser to v3.23.0.
- Update FAQ '`How to disable the tagline in signatures?`' in the [FAQ section](https://set-outlooksignatures.com/faq) with lots of background information.
- Update font formatting and phrasing in sample template files.
- Switch from the SaveAs Word method to SaveAs2, as the first is no longer documented and the latter is supported since Word 2010.
### Added
- Add hints in documentation that not only primary SMTP addresses are supported, but also alias and secondary addresses.
- Add FAQ '`When should I refer on-prem groups and when Entra ID groups?`' to the [FAQ section](https://set-outlooksignatures.com/faq).
- Add FAQ '`Why are signatures and out-of-office replies recreated even when their content has not changed?`' to the [FAQ section](https://set-outlooksignatures.com/faq).
- Allow '`//`' as additional comment marker in INI files.
- Add more comments in sample code files based on client feedback.
### Removed
### Fixed
- Fix sample code '`Intune-SetOutlookSignatures-Remediate.ps1`' so that files and folders are extracted with the correct item type from the downloaded ZIP file. This avoids that some files are extracted as folders instead of files, leading to errors when Set-OutlookSignatures is started from the remediation script.
- Update the workarounds to overcome Microsoft limitations for Outlook add-ins.
- Fix handling of special characters when applying encoding corrections for Outlook signatures.
- Prohibit running Set-OutlookSignatures and sample code in PowerShell ISE as this environment misses required features.
- Add more information about deploying and debugging the Outlook add-in in the '`The Outlook add-in`' chapter in the [FAQ section](https://set-outlooksignatures.com/faq) and in the Outlook add-in configuration file.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.14.1" target="_blank">v4.14.1</a> - 2024-08-29
_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works.<br>Set-OutlookSignatures supports cloud roaming signatures - see `MirrorCloudSignatures` in the [parameters documentation](https://set-outlooksignatures.com/parameters)._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Tagline**_  
_Starting with this release, a tagline is added to each signature deployed for mailboxes without a [Benefactor Circle](https://set-outlooksignatures.com/benefactorcircle) license. See FAQ `How to disable the tagline in signatures?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details._

### Fixed
- Provide a workaround to avoid problems loading DLL dependencies on some systems with certain .Net patch levels.

## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.14.0" target="_blank">v4.14.0</a> - 2024-08-28
_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works.<br>Set-OutlookSignatures supports cloud roaming signatures - see `MirrorCloudSignatures` in the [parameters documentation](https://set-outlooksignatures.com/parameters)._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Tagline**_  
_Starting with this release, a tagline is added to each signature deployed for mailboxes without a [Benefactor Circle](https://set-outlooksignatures.com/benefactorcircle) license. See FAQ `How to disable the tagline in signatures?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details._

### Changed
- Make simulation mode results for defaultNew and defaultReplyFwd signatures cross-platform compatible by not creating shortcuts (Windows) or symbolic links (Linux, macOS), but actually copying the files.
- Update handling of SharePoint WebDAV authentication to reflect not yet documented Microsoft API changes.
- Update sample code `Create-EntraApp.ps1` for new `Files.Read.All` permission and add a hint to consider switching to `Files.SelectedOperations.Selected`.
- Remove parameters and SharePoint sharing hints ('/:u:/r', etc.) from URLs passed as paths for template and configuration files: 'https://YourTenant.sharepoint.com/:u:/r/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini?SomeParam1=1&SomeParam2=2' becomes 'https://yourtenant.sharepoint.com/sites/SomeSite/SomeLibrary/SomeFolder/SomeFile.ini'
- Update dependency MSAL.Net to v4.64.0.
### Added
- Add an Outlook add-in that can:  
  - Automatically add signatures to new emails and reply emails (including Outlook on iOS and Outlook on Android)
  - Automatically add signatures to new appointment invites

  See the [Outlook add-in documentation](https://set-outlooksignatures.com/outlookaddin) for details about features, usage, configuration, deployment and remarks.  
  The Outlook add-in is available as part of the [Benefactor Circle](https://set-outlooksignatures.com/benefactorcircle) license.
- Detect if a template or configuration file path is hosted on SharePoint Online and use Microsoft Graph to access it. This enables SharePoint Online access on Linux and macOS.  
  To use this, you need to update your Entra ID app with the delegated (not application) permission `Files.Read.All` (`Files.SelectedOperations.Selected` also works but requires additional configuration). You also need to set the new parameter `GraphClientID`, which is described below.  
  Access to SharePoint on-prem is still limited to Windows and WebDAV with an authentication cookie created in Internet Explorer.
- Add new parameter `GraphClientID`. Must be set when `GraphConfigFile` is hosted on SharePoint Online. See the [parameters documentation](https://set-outlooksignatures.com/parameters) for details.
- Show local and UTC time for SimulateTime parameter when in simulation mode.
- Add a tagline to each signature deployed for mailboxes without a [Benefactor Circle](https://set-outlooksignatures.com/benefactorcircle) license. See FAQ `How to disable the tagline in signatures?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details.
- Remove trailing null character from file names being enumerated in SharePoint folders. .Net and the WebDAV client sometimes add a null character, which is not allowed in file and path names.
- Detect corrupted Outlook on Windows installations by comparing the major part of the Outlook version information from the registry and from outlook.exe.
- Add `Which group naming format should I choose?` recommendation to the [FAQ section](https://set-outlooksignatures.com/faq) and INI files.
- Add FAQ `Why is the out-of-office assistant not activated automatically?` to the [FAQ section](https://set-outlooksignatures.com/faq).
### Fixed
- Do not allow parameter `SignatureCollectionInDrafts` in simulation mode unless `SimulateAndDeploy` is enabled.
- Fix a potential problem with paged Microsoft Graph queries which could lead to an infinite loop.
- Fix search for Entra ID groups by their display name. (Thanks <a href="https://github.com/CoreyS222" target="_blank">@CoreyS222</a>!)
- Fix on-prem group membership search not including non-security enabled distribution groups due to a regression.
- Fix not finding a site ID for SharePoint Online paths directly in the root default site collection and not in /sites or /teams.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.13.0" target="_blank">v4.13.0</a> - 2024-06-27
_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorCloudSignatures` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update dependency MSAL.Net to v4.61.3.
- Update code in MSAL.PS used to detect active window handle, required for interactive login to Graph.
- Do not add DOCX templates to the list of recent files in Word, and do not route the document to the next recipient if the document has a routing slip.
### Added
- Add a hint in INI files stating to not modify them directly but a copy of them, and to follow the README FAQ `What is the recommended folder structure for script, license, template and config files?`.
- Add a fallback mechanism for detecting the User Principal Name of the currently logged-in user on Windows. This solves the rare case with unknown root cause where this information is not available in the registry. (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/pull/116" target="_blank">#116</a>) (Thanks <a href="https://github.com/CarlInLV" target="_blank">@CarlInLV</a>!)
- Add new parameter `SignatureCollectionInDrafts`. Enabled per default, this creates and updates an email message with the subject 'My signatures, powered by Set-OutlookSignatures Benefactor Circle' in the drafts folder of the current user, containing all available signatures in HTML and plain text for easy access in mail clients that do not have a signatures API. See the [parameters documentation](https://set-outlooksignatures.com/parameters) for details.
### Removed
- Remove dependency to System.Web.dll for URL encode and decode operations.
### Fixed
- Truncate text signature for SignatureTextOnMobile when setting Outlook Web signature if it is longer than 512 characters, and show a warning message. SignatureTextOnMobile is only used when browsing Outlook Web on a mobile device. (Thanks <a href="https://github.com/bmartins-EMCDDA" target="_blank">@bmartins-EMCDDA</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.12.2" target="_blank">v4.12.2</a> - 2024-05-27
_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorCloudSignatures` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Fixed
- Convert account pictures from Active Directory for use in signatures without byte array conversion error.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.12.1" target="_blank">v4.12.1</a> - 2024-05-24
_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorCloudSignatures` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._

_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

### Changed
- Update dependency MSAL.Net to v4.61.1.
- Use pure .Net methods to create Windows shortcut files in sample code and in Set-OutlookSignatures, as Microsoft marked VBS (Visual Basic Script) as deprecated and will remove it in future versions of Windows. 
### Fixed
- Name temporary file names correctly when the `MailboxSpecificSignatureNames` parameter is enabled, so that out-of-office replies can be set in this mode.
- Show the 'MSAL.PS Graph token cache info' when run in Windows PowerShell (PowerShell 5.x).
- Fix simulation mode wrongly deleting all '___Mailbox *' folders in the output directory, but the last one.
- Benefactor Circle only: Search for license groups in Entra ID via on-prem SIDs may fail because of a regression in checking the SID format.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.12.0" target="_blank">v4.12.0</a> - 2024-05-07
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorCloudSignatures` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Changed
- Graph connectivity: Update dependency MSAL.Net to v4.60.3.
- The Quick Start Guide is now easier to follow and clearer on requirements.
- Updated sample templates (full barrier-free design, and other smaller changes).
- Updated `Authentication` chapter in the [technical details documentation](https://set-outlooksignatures.com/details).
- The API for saving photos and user-defined images has been changed, which significantly speeds up the creation of temporary files.
### Added
- Support for Linux and macOS. Some restrictions apply to Non-Windows platforms, see `Linux and macOS` in the `Requirements` chapter of the [technical details documentation](https://set-outlooksignatures.com/details) for details.
- Custom image replacement variables that you can fill yourself with a byte array: `$CurrentUserCustomImage[1..10]$`, `$CurrentUserManagerCustomImage[1..10]$`, `$CurrentMailboxCustomImage[1..10]$`, `$CurrentMailboxManagerCustomImage[1..10]$`. Use cases: Account pictures from a share, QR code vCard/URL/text/Twitter/X/Facebook/App stores/geo location/email, etc.
  - QR code vCard (MeCard) in custom image replacement variable `$Current[..]CustomImage1$`. See file `.\config\default replacement variables.ps1` for the easily customizable code used to create it. See the [technical details documentation](https://set-outlooksignatures.com/details) for details.
- Support for maximum barrier-free accessibility with screen readers and comparable tools. Use Word ScreenTips and HTML titles for hyperlinks, and alt text for images, replacement variables are supported. All sample templates have been updated accordingly. Just hover your mouse pointer over a hyperlink or image to see additional information.
- Show a warning when authentication to Outlook Web (no matter if on-prem or in the cloud) is not possible via Autodiscover, as this means that Autodiscover is not configured correctly.
- The `MirrorLocalSignaturesToCloud` parameter is now also available under the name `MirrorCloudSignatures`. The old name may be removed with the next major release (v5.0.0).
### Fixed
- When using HTM templates, image paths were not correctly modified when containing one of the reserved characters defined in RFC3986 (`:/?#[]@!$&'()*+,;=`, Uri.EscapeUriString vs. Uri.EscapeDataString).


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.11.0" target="_blank">v4.11.0</a> - 2024-03-26
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Changed
- Updated sample templates
- Authentication against SharePoint document libraries containing templates or INI files has been adapted to Microsoft API changes not yet documented
### Added
- New parameter `MailboxSpecificSignatureNames`. By setting the `MailboxSpecificSignatureNames` parameter to `true`, the email address of the current mailbox is added to the name of the signature - instead of a single `Signature A` file, Set-OutlookSignatures can create a separate signature file for each mailbox: `Signature A (user.a@example.com)`, `Signature A (mailbox.b@example.net)`, etc. See the [parameters documentation](https://set-outlooksignatures.com/parameters) for details.
- When `MirrorLocalSignaturesToCloud` is enabled,
  - download signatures not only from the current user's mailbox, but from all mailboxes where possible (full access with current user's credentials, due to Microsoft restrictions)
  - upload personal signatures to Exchange Online mailbox of logged-on user right after they have been created and before they are even copied to a local signature path, avoiding a race condition with the internal download schedule of Outlook
- Always use the most recent TLS version available and supported for communication with Microsoft services
### Fixed
- DeleteUserCreatedSignatures, DeleteScriptCreatedSignaturesWithoutTemplate and AdditionalSignaturePath now handle folder and file names with special characters ($'‘’‚‛) correctly
- Check for PowerShell Full Language mode before checking for new versions to avoid an error message without a hint to the root cause
- When an Outlook profile is available, set default signatures no matter if current mailbox is a dummy mailbox or not
- Correctly handle INI files not containing SortOrder or SortCulture information


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.10.1" target="_blank">v4.10.1</a> - 2024-02-06
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Changed
- Graph connectivity: Update dependency MSAL.Net to v4.59.0
- Updated FAQ `How can I deploy and run Set-OutlookSignatures using Microsoft Intune?` in the [FAQ section](https://set-outlooksignatures.com/faq), and moved sample code from FAQ to `.\sample code` directory
- Active Directory attribute names are no longer case sensitive, making it easier creating own replacement variables and adding custom schema extensions
- Only use Integrated Windows Authentication for connection to Outlook Web when no cloud token is available
- If the variable replacement in Word fails, an additional note is displayed about the possible mandatory use of Microsoft Purview Information Protection.
- Updated sample templates
### Added
- New FAQs `Keep users from adding, editing and removing signatures` and `What is the recommended folder structure for script, license, template and config files?` in the [FAQ section](https://set-outlooksignatures.com/faq).
### Fixed
- Workaround for a Microsoft Graph API problem, which returns a HTTP 403 error when querying the settings of some mailboxes. This query now only happens when absolutely necessary. When this query fails, `SetCurrentUserOutlookWebSignature` and `SetCurrentUserOOFMessage` are disabled. Only Microsoft can fix the root cause of this problem.
- More workarounds for timing problems with file operations in folders used by OneDrive
- Verbose output no longer logs Graph tokens, only their properties (header and payload data only, no digital signature)
- When an existing signature is overwritten by a new signature and the two signature names only differ in upper and lower case ("signature a" and "Signature A", for example), always use the casing of the new signature
- Correctly detect automapped mailboxes in Outlook Web when New Outlook is not used
- Add scope 'Application.ReadWrite.All' to sample script `.\sample code\Create-EntraApp.ps1`


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.10.0" target="_blank">v4.10.0</a> - 2024-01-05
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Changed
- Graph connectivity: Update dependency MSAL.Net to v4.58.1
### Added
- Added to the description of each parameter in the script itself and the [parameters documentation](https://set-outlooksignatures.com/parameters):
  - Allowed values
  - Usage examples (PowerShell and Non-PowerShell)
  - Information when a feature requires a Benefactor Circle license
- Additional descriptions in template INI files
- A specific warning when the template INI file contains references to templates with a wrong file extension (for example, .html instead of .htm)
- New FAQs, see the [FAQ section](https://set-outlooksignatures.com/faq) for details:
  - `How can I deploy and run Set-OutlookSignatures using Microsoft Intune?`
  - `Why does Set-OutlookSignatures run slower sometimes?`
### Fixed
- Correctly handle empty Outlook profiles and a no longer existing default profile
- Graph authentication: Workaround for MSAL.Net "connection reset error" in browser. See `$GraphHtmlMessageSuccess` and `$GraphHtmlMessageError` in `.\config\default graph config.ps1` for details.
- Graph authentication: Workaround for MSAL.Net returning the access token in a different format in interactive and silent authentication


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.9.0" target="_blank">v4.9.0</a> - 2023-12-02
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Changed
- Word is no longer required to convert signatures to TXT format, which reduces resource consumption and execution time
- `MoveCssInline`: Update dependency PreMailer.Net to v2.5.0
- Graph connectivity: Update dependency MSAL.Net to v4.58.0
### Added
- Check for valid Windows/Outlook/Outlook Web signature names when using the `MirrorLocalSignaturesToCloud` parameter or the `OutlookSignatureName` INI parameter
- `MirrorLocalSignaturesToCloud`: Only download a roaming signature from Exchange Online when its local version is older or does not exist
- Sample code `SimulateAndDeploy.ps1`
  - Display info when a specific job ends, in addition to when there is an error
  - Display info at least once a minute, and make this update interval configurable
  - Separate logs for sample code output, successful jobs, jobs with errors and details of each job
- New parameter `ScriptProcessPriority`: Define the script process priority. With lower values, Set-OutlookSignatures runs longer but minimizes possible performance impact. See `README` for details.
### Fixed
- `DeleteScriptCreatedSignaturesWithoutTemplate` and `DeleteUserCreatedSignatures` did not remove all subfolders belonging to a signature
- When using HTM templates with images in a connected folder, the folder name was not corrected reliably, which resulted in missing images (as they pointed to a wrong path)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.8.1" target="_blank">v4.8.1</a> - 2023-11-24
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Fixed
- Embedding images and loading the Graph config file is now much faster because of switching from PowerShell cmdlets to .Net system calls for conversions to Base64 (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/pull/95" target="_blank">#95</a>)
- The manager of the current user was not correctly detected in certain hybrid scenarios (Thanks Thomas Müllerchen and Tommy Malodisdach!)
- The message informing the user before a new browser tab is opened for authentication now does not steal the focus in all combinations of PowerShell Core/Desktop and Windows Terminal/Classic Console Host


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.8.0" target="_blank">v4.8.0</a> - 2023-11-20
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Changed
- Update dependency MSAL.Net to v4.57.0.
- Allow more alternate names for cloud environments. See the [parameters documentation](https://set-outlooksignatures.com/parameters) for details about the `CloudEnvironment` parameter.
### Added
- Graph token cache now not only works in Windows PowerShell 5.1, but also in PowerShell 7. 
- A message box now informs the user before a new browser tab is opened for authentication, as Microsoft still does not show the Entra ID app name in the authentication prompt. The message text can be customized or disabled with the `$GraphHtmlMessageboxText` parameter in `.\config\default graph config.ps1`. See that file for details.
- The HTML message after a successful browser authentication can be customized with the `$GraphHtmlMessageSuccess` parameter in `.\config\default graph config.ps1`. See that file for details, and also consider `$GraphBrowserRedirectSuccess` for a redirection alternative.
- The HTML message after an unsuccessful browser authentication can be customized with the `$GraphHtmlMessageError` parameter in `.\config\default graph config.ps1`. See that file for details, and also consider `$GraphBrowserRedirectError` for a redirection alternative.
- New sample code `.\sample code\Create-EntraApp.ps1` automates the creation of the Entra ID app required to access Microsoft Graph.
- New FAQ `How do I alternate banners and other images in signatures?`. See the [FAQ section](https://set-outlooksignatures.com/faq) for details.
### Fixed
- MirrorLocalSignaturesToCloud now correctly detects cloud mailboxes in hybrid environments when a connection to the on-prem Active Directory is used
- Setting Word process priority no longer leads to an error in PowerShell 7
- Fix a regression introduced with the option to set the Word process priority. This fix avoids a rare problem where a manually started Word instance connects to the Word background process created by the software.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.7.0" target="_blank">v4.7.0</a> - 2023-10-29
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Changed
- Reduced minimum required Graph permissions: GroupMember.Read.All is now used instead of Group.Read.All, which reduces delegated app permissions.  
To compare the two permissions, see their description at [Microsoft Graph permission reference](https://learn.microsoft.com/en-us/graph/permissions-reference).
  - If you use an Entra ID app defined in your own tenant
    - If you want to use GroupMember.Read.All instead of Group.Read.All, you have to use at least v4.6.0 of Set-OutlookSignatures.  
    Remove admin consent for the Group.Read.All permission and remove it from the app, then add the delegated GroupMember.Read.All permission and grant admin consent for it.
    - If you do not want to use GroupMember.Read.All instead of Group.Read.All, there is nothing to do: v4.6.0 and up just request all permissions defined in the Entra ID app, versions before v4.6.0 explicitly request Group.Read.All.
  - If you use the Entra ID app provided by the developers (app ID 'beea8249-8c98-4c76-92f6-ce3c468a61e6')
    - If you want to use GroupMember.Read.All instead of Group.Read.All, you have to use at least v4.6.0 of Set-OutlookSignatures, and renew your admin consent for the new reduced permissions:
      1. Open a browser, preferably in a private window
      2. Open the URL 'https://login.microsoftonline.com/organizations/adminconsent?client_id=beea8249-8c98-4c76-92f6-ce3c468a61e6'
      3. Log on with a user that has Global Admin or Client Application Administrator rights in your tenant
      4. Accept the required permissions on behalf of your tenant. You can safely ignore the error message that the URL 'http://localhost/?admin_consent=True&tenant=[…]' could not be found or accessed.
    - If you want to use a version older than v4.6.0, you need to create your own app in your own Entra ID tenant as detailed in `.\config\default graph config.ps1`.
    - For security and maintenance reasons, it is recommended to create you own app in your own tenant.
### Added
- Support for all cloud environments via new parameter `CloudEnvironment`: Public (AzurePublic), US Government L4 (AzureUSGovernment), US Government L5 (AzureUSGovernment DoD), China (AzureChinaCloud operated by 21Vianet). See the [parameters documentation](https://set-outlooksignatures.com/parameters) for details.
- The software now shows a hint at startup when a newer release is available on GitHub.
### Fixed
- Implementation approach: Translated a sentence to English, which was only available in German (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/pull/89" target="_blank">#89</a>) (Thanks <a href="https://github.com/JeroenOortwijn" target="_blank">@JeroenOortwijn</a>!)
- Update dependency MSAL.PS so that process ID is correctly determined when run in Windows Terminal (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/88" target="_blank">#88</a>) (Thanks <a href="https://github.com/Ben-munich" target="_blank">@Ben-munich</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.6.1" target="_blank">v4.6.1</a> - 2023-10-27
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Fixed
- Correctly detect and restore Word registry key 'DisableWarningOnIncludeFieldsUpdate'
- Simulation mode: Show images in out-of-office replies, even though Exchange does not support them yet
- SimulateAndDeploy.ps1: Advanced error handling
- Implementation approach: Translated a sentence to English, which was only available in German (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/pull/89" target="_blank">#89</a>) (Thanks <a href="https://github.com/JeroenOortwijn" target="_blank">@JeroenOortwijn</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.6.0" target="_blank">v4.6.0</a> - 2023-10-23
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details on how this feature works. Set-OutlookSignatures supports cloud roaming signatures - see `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Added
- 'default graph config.ps1' now includes a description for each permission required by the Entra ID app for Graph access
- Additional documentation: Implementation approach
  - The content is based on real-life experiences implementing the software in multi-client environments with a five-digit number of mailboxes.
  - Proven procedures and recommendations for product managers, architects, operations managers, account managers, mail and client administrators. Suited for service providers as well as for clients.
  - It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.
  - Available in English and German.
- New and improved sample code
  - SimulateAndDeploy.ps1: Deploy signatures without end user interaction, running Set-OutlookSignatures on a server - including support for roaming signatures
  - Test-ADTrust.ps1: Detect why a client cannot query Active Directory information
  - SimulationModeHelper.ps1: Makes using simulation mode easier. An admin sets the parameters in the software, the content creators execute it and just have to enter the values required for simulation:
    - The user to simulate (mandatory)
    - The mailbox(es) to simulate (optional)
    - The time to simulate (optional)
    - The output path (optional)
- A basic configuration user interface with grouped parameter sets, just run `Show-Command .\Set-OutlookSignatures.ps1` in PowerShell.
- New parameter `WordProcessPriority`: Define the Word process priority. With lower values, Set-OutlookSignatures runs longer but minimizes possible performance impact. See `README` for details.
### Fixed
- On-prem only: Make sure that Active Directory attributes of the current user, the current mailbox and their managers are available in environments where not every domain controller is also a global catalog server
- Always connect to Entra ID/Graph when New Outlook is used 


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.5.0" target="_blank">v4.5.0</a> - 2023-09-29
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._  
_Set-OutlookSignatures can experimentally handle cloud roaming signatures since v4.0.0. See `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Changed
- Adapt program logic to cloud roaming signatures API changes introduced by Microsoft
- Updated FAQ `How can I log the software output?`. See `README` for details.
### Added
- New parameter `EmbedImagesInHtmlAdditionalSignaturePath`. See `README` for details.
- New FAQ `How can I start the software only when there is a connection to the Active Directory on-prem?`. See `README` for details.
### Fixed
- Variables in HTM templates have not been replaced with actual values because of a wrong RegEx syntax
- Content of path defined in `AdditionalSignaturePath` was not deleted before copy operations.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.4.0" target="_blank">v4.4.0</a> - 2023-09-20
_**Add features with the Benefactor Circle add-on and get professional support from ExplicIT Consulting**_  
_See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, Exchange Online admins**_  
_Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._  
_Set-OutlookSignatures can experimentally handle cloud roaming signatures since v4.0.0. See `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Added
- Support for New Outlook. Mailboxes are taken from the first matching source:
  1. Simulation mode is enabled: Mailboxes defined in SimulateMailboxes
  2. Outlook is installed and has profiles, and New Outlook is not set as default: Mailboxes from Outlook profiles
  3. New Outlook is installed: Mailboxes from New Outlook (including manually added and automapped mailboxes for the currently logged-in user)
  4. If none of the above matches: Mailboxes from Outlook Web (including manually added mailboxes, automapped mailboxes follow when Microsoft updates Outlook Web to match the New Outlook experience)
### Fixed
- Correctly handly write protected files when copying to AdditionalSignaturePath


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.3.0" target="_blank">v4.3.0</a> - 2023-09-08
_**Some features are exclusive to the commercial Benefactor Circle add-on.**_
- _See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, cloud mailbox users:**_
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._  
- _Set-OutlookSignatures can experimentally handle roaming signatures since v4.0.0. See `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
### Added
- When no mailboxes are configured in Outlook, additional mailboxes configured in Outlook Web are used. Thanks to our partner [ExplicIT Consulting](https://explicitconsulting.at) for donating this code, enabling another world-first feature and bringing us even closer to supporting the "new Outlook" client (codename "Monarch") in the future!
- Add hint to TLS 1.2 when Entra ID/Graph authentication is not successful (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/85" target="_blank">#85</a>) (Thanks <a href="https://github.com/halatovic" target="_blank">@halatovic</a>!)
- Update '`Quick Start Guide`' in '`README`' file with clearer instructions on how to register the Entra ID app required for hybrid and cloud-only environments


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.2.1" target="_blank">v4.2.1</a> - 2023-08-16
_**Some features are exclusive to the commercial Benefactor Circle add-on.** See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, cloud mailbox users:**_
- _**Set-OutlookSignatures can now experimentally handle roaming signatures!** See `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._
### Fixed
- MoveCSSInline may not find a dependent DLL on some systems (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/84" target="_blank">#84</a>) (Thanks <a href="https://github.com/panki27" target="_blank">@panki27</a>!)
- An error occurred when a trust of the forest root domain of an on-prem Active Directory to itself was detected (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/83" target="_blank">#83</a>) (Thanks <a href="https://github.com/panki27" target="_blank">@panki27</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.2.0" target="_blank">v4.2.0</a> - 2023-08-10
_**Some features are exclusive to the commercial Benefactor Circle add-on.** See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, cloud mailbox users:**_
- _**Set-OutlookSignatures can now experimentally handle roaming signatures!** See `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._
### Added
- New parameter `MoveCSSInline` to move CSS to inline style attributes, for maximum email client compatibility. This parameter is enabled per default, as a workaround to Microsoft's problem with formatting in Outlook Web (M365 roaming signatures and font sizes, especially).
### Fixed
- Set Word WebOptions in correct order, so that they do not overwrite each other


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.1.0" target="_blank">v4.1.0</a> - 2023-07-28
_**Some features are exclusive to the commercial Benefactor Circle add-on.** See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, cloud mailbox users:**_
- _**Set-OutlookSignatures can now experimentally handle roaming signatures!** See `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._
### Added
- Templates can now be **assigned to or excluded for specific replacement variables of the current user or the current mailbox**. Thanks to [ExplicIT Consulting](https://explicitconsulting.at) for donating this code!  
See `Template tags and INI files` in `README` for details and examples.  
Use cases:
  - Assign template to a specific mailbox or user, but only if user or mailbox is member in multiple groups at the same time.
  - Assign template to users or mailboxes which have or have not a value in a replacement variable.
  - Every replacement variable can be used: Current user and current mailbox, their managers, or tailored replacement variables defined in a custom replacement variable config file.
- Templates can now be **assigned to or excluded for specific email addresses or groups SIDs of the _mailbox of the current user_**. Thanks to [ExplicIT Consulting](https://explicitconsulting.at) for donating this code!  
See `Template tags and INI files` in `README` for details and examples.  
Use cases:
  - Assign template to a specific mailbox, but not if the _mailbox of the current user_ has a specific email address or is member of a specific group. It does not matter if this personal mailbox is added in Outlook or not.  
This is useful for delegate and boss-secretary scenarios - secretaries get specific delegate template for boss's mailbox, but the boss not. **Combine this with the feature that one template can be used multiple times in the INI file, and you basically only need one template file for all delegate combinations in the company!**
  - Assign a template to the mailbox of a specific logged-in user or deny a template for the mailbox of a specific user, no matter which mailboxes the user has added in Outlook.
- The attribute 'GroupsSIDs' is now also available in the `$AdPropsCurrentUser` replacement variable. It contains all the SIDs of the groups the mailbox of the current user is a member of, which allows for replacement variable content based on group membership, as well as assigning or denying templates for specific users. See `Delete images when attribute is empty, variable content based on group membership` in `README` for details and examples.
- Replacement variables are no longer case sensitive. This eliminates a common error source and makes replacement variables in template files easier to read.
- New chapter `Proposed template and signature naming convention` in `README` file. Thanks to [ExplicIT Consulting](https://explicitconsulting.at) for donating this piece of documentation!
### Fixed
- The attribute 'GroupsSIDs' is now reliably available in the `$AdPropsCurrentMailbox` replacement variable.
- Correctly log group and email address specific exclusions (only an optical issue, no technical one)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v4.0.0" target="_blank">v4.0.0</a> - 2023-07-12
_**Some features are exclusive to the commercial Benefactor Circle add-on.** See ['`Benefactor Circle add-on`'](https://set-outlooksignatures.com/benefactorcircle) for details about these features and how you can benefit from them with a Benefactor Circle license._

_**Attention, cloud mailbox users:**_
- _**Set-OutlookSignatures can now experimentally handle roaming signatures!** See `MirrorLocalSignaturesToCloud` in the [parameters documentation](https://set-outlooksignatures.com/parameters) for details._
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._
### Changed
- **Breaking:** Benefactor Circle members have exclusive access to the following features:
  - Prioritized support and feature requests
    - Issues are handled with priority via a Benefactor Circle exclusive email address and a callback option.
    - Requests for new features are checked for feasability with priority.
    - All release upgrades during the license period are for free, no matter if it is a patch, feature or major release.
  - Script features
    - Time-based campaigns by assigning time range constraints to templates
    - Signatures for automapped and additional mailboxes
    - Set current user Outlook Web signature (classic Outlook Web signature and roaming signature)
    - Download and upload roaming signatures
    - Set current user out-of-office replies
    - Delete signatures created by the software, for which the templates no longer exist or apply
    - Delete user-created signatures
    - Additional signature path (when used outside of simulation mode)
    - High resolution images from DOCX templates
- **Breaking:** The `CreateRTFSignatures` parameter now defaults to `false`, because the RTF format for emails is hardly used nowadays.
- **Breaking:** The `EmbedImagesInHtml` parameter now defaults to `false` because certain recent versions of Word (which Outlook uses as HTML renderer) incorrectly hande embedded images. See `Images in signatures have a different size than in templates, or a black background` in `README` for details.
When `EmbedimagesInHtml` is enabled, it now automatically enables the "Send pictures with document" Outlook registry key. 
- **Breaking:** New parameter `DisableRoamingSignatures` defaults to `true`. See `README` for details.
- Updated FAQ `Images in signatures have a different size than in templates`. See `README` for details.
### Added
- `Quick Start Guide`. See `README` for details.
- New parameter `MirrorLocalSignaturesToCloud` for experimentally handling roaming signatures. Disabled by default. See `README` for details.
- Thanks to our partnership with [ExplicIT Consulting](https://explicitconsulting.at), Set-OutlookSignatures and its components are digitally signed with an Extended Validation (EV) Code Signing Certificate (which is the highest code signing standard available).  
This is not only available for Benefactor Circle members, but also the Free and Open Source core version is code signed. Code signing makes it much easier to implement Set-OutlookSignatures in environments being locked down with AppLocker or comparable tools.
- All replacement variables now have the 'DELETEEMPTY' option, which allows for images to be kept only when an attribute has a value. See `Delete images when attribute is empty, variable content based on group membership` in `README` for details and examples.
- The attribute 'GroupsSIDs' is now available in the `$CurrentMailbox` variable for use with replacement variables. It contains all the SIDs of the groups the current mailbox is a member of, which allows for replacement variable content based on group membership. See `Delete images when attribute is empty, variable content based on group membership` in `README` for details and examples.
- A basic configuration user interface with grouped parameter sets, just run `Show-Command .\Set-OutlookSignatures.ps1` in PowerShell.
- The new template tag `WriteProtect` write protects individual signature files. See `README` for details and restrictions.
- The new script parameter `SimulateTime` allows to use a specific time when running simulation mode, which is handy for testing time-based templates.
- New parameter `DisableRoamingSignatures`. See `README` for details.
- New FAQ `Start Set-OutlookSignatures in hidden/invisible mode`. See `README` for details.
- Show a warning message when setting the Outlook Web signature is not possible because Outlook Web has not been initialized yet, making it impossible to set signature options in Outlook Web without breaking the first log in experience for this mailbox (getting asked for language, timezone, etc.)
- Copy HTM image width and height attributes to style attribute
- Show a warning when a template contains images formatted as non-inline shapes, as these image formatting options may not be supported by Outlook (e.g., behind the text)
- Support for mailboxes in the user's Entra ID tenant with different UPN/user ID and primary SMTP address
- The Word registry key `DontUseScreenDpiOnOpen` is set to `1` automatically, according to Microsoft documentation (see `README` for details). This helps avoid image sizing problems on devices with non-standard DPI settings.
### Fixed
- Simulation mode
  - Simulation mode partly returned data of the currently logged on user instead of data of the simulated user
  - Simulation mode did not prioritize mailbox list correctly
  - When SimulateUser is not defined, but other simulation parameters, the software exits
  - When multiple mailboxes to simulate were passed, only the first one was considered 
- Graph queries did not correctly handle paged results
- Restoring the Word setting 'ShowFieldCodes' now works correctly in more error scenarios
- Benefactor Circle members only: Additional and automapped mailboxes have not been detected reliably
- Categorizing template files is now much faster than before (two seconds instead of two minutes for 250 templates)
- Replacing variables in DOCX templates is now faster than before, as only variables actually being used in the document are replaced
- Realiably remove '$Current[…]Photo$' string from image alt text
- `SimulateAndDeploy.ps1`: Correctly convert HTML image tags with embedded images and additional options
- Display sort order for was not handled correctly when primary smtp address of a mailboxes has been changed after it was already added to Outlook


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.6.1" target="_blank">v3.6.1</a> - 2023-05-22
_Attention cloud mailbox users: Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._
### Fixed
- Signatures created with `DocxHighResImageConversion true` in combination with `EmbedImagesInHtml false` include high resolution image files, but these images could not be displayed because a wrong path was set in the HTM file


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.6.0" target="_blank">v3.6.0</a> - 2023-01-24
_Attention cloud mailbox users: Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details, known problems and workarounds._
### Changed
- Microsoft Information Protection sensitivity labels are now supported when using DOCX templates. See `How to make Set-OutlookSignatures work with Microsoft Information Protection?` in the [FAQ section](https://set-outlooksignatures.com/faq) for details.
- Shrinking RTF files is now compatible with Microsoft Information Protection sensitivity labels
- Updated chapter in the [technical details documentation](https://set-outlooksignatures.com/details): `Photos from Active Directory`
### Added
- New FAQ in the [FAQ section](https://set-outlooksignatures.com/faq): `How to make Set-OutlookSignatures work with Microsoft Information Protection?`
- New parameter `DocxHighResImageConversion`. Enabled by default, this parameter creates HTM signatures with high resolution images from DOCX templates. See the [parameters documentation](https://set-outlooksignatures.com/parameters) for details.
### Fixed
- User photo placeholders were not replaced when using HTM templates with images stored in connected sub-folders
- Sample template files `Test all default replacement variables.docx` and `Test all default replacement variables.htm` did not contain all default replacement variables
- Correctly handle empty `AdditionalSignaturePath` parameter


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.5.1" target="_blank">v3.5.1</a> - 2022-12-20
_Attention cloud mailbox users: Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `README` for details, known problems and workarounds._
### Fixed
- Use different code to determine Outlook and Word executable file bitness, as the .Net APIs used before seem to fail randomly with the latest Windows and .Net updates (especially when using 32-bit PowerShell on 64-bit Windows)
- Do not stop the software when `SignaturesForAutomappedAndAdditionalMailboxes` is enabled and the Outlook file path cannot be determined


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.5.0" target="_blank">v3.5.0</a> - 2022-12-19
_Attention cloud mailbox users: Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `README` for details, known problems and workarounds._
### Changed
- Mailbox prioritization: Within an Outlook profile, mailbox priority is now determined by the sort order shown in Outlook, no longer by the time a mailbox has been added to the profile. See `Signature and OOF application order` in `README` for more details about mailbox prioritization.
- `README`: Update FAQ `What about the roaming signatures feature in Exchange Online?`
### Fixed
- Mailbox priority list: Don't add automapped or additional mailboxes to the end of the mailbox priority list, but to the end of each Outlook profile in the mailbox priority list
- Mailbox priority list: When duplicates exist, only show the mail address with the highest priority. Verbose output contains additional information for each occurrence (Outlook profile name, registry path, legacyExchangeDN)
- Setting default signature: Show Outlook profile name, so that mailboxes that exist in multiple Outlook profiles can be distinguished
- When the logged in user's personal mailbox exists in multiple profiles, set the Outlook Web signature and OOF message for this mailbox only once
- When a mailbox exists in multiple Outlook profiles, only query AD/Graph at the first occurrence and use cached data on remaining occurrences


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.4.1" target="_blank">v3.4.1</a> - 2022-11-25
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Fixed
- Correctly handle logged in user with empty mail attribute
- Correctly enumerate SID and SidHistory when connected to a local Active Directory user with a mailbox in Exchange Online (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/59" target="_blank">#59</a>) (Thanks <a href="https://github.com/AnotherFranck" target="_blank">@AnotherFranck</a>!)
- Correctly handle empty templates, signatures and OOF messages


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.4.0" target="_blank">v3.4.0</a> - 2022-11-02
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Added
- New parameter `IncludeMailboxForestDomainLocalGroups`, see `README` for details
- LDAP and Global Catalog connectivity is now additionally checked for every child domain of the current user's Active Directory forest and every child domain of cross-forest trusts
- Consider SID history of groups in trusted domains/forests
- New FAQs in the [FAQ section](https://set-outlooksignatures.com/faq): `What if Outlook is not installed at all?` and `What if a user has no Outlook profile or is prohibited from starting Outlook?`
### Fixed
- Correctly calculate mailbox priority when simulation mode is enabled and/or the email address is a secondary address
- On-prem: Membership in domain local groups is now recognized if the group is in a child domain of a forest connected with a cross-forest trust
- Only consider mailboxes as additional mailboxes when they appear in Outlook's list in the email navigation pane. This avoids falsely adding shared calendars as additional mailboxes.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.3.0" target="_blank">v3.3.0</a> - 2022-09-05
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Changed
- Use different method to delete files to avoid occassional OneDrive error "access to the cloud file is denied"
- Update logo and icon
### Added
- the software now detects not only primary mailboxes configured in Outlook, but also automapped and additional mailboxes. This behavior can be disabled with the new parameter `SignaturesForAutomappedAndAdditionalMailboxes`. See `README` for details.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.2.2" target="_blank">v3.2.2</a> - 2022-08-12
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Fixed
- When the `EmbedImagesInHtml` parameter is set to `false`, correctly handle a certain file system condition instead of stopping processing the current template after the 'Embed local files in HTM format and add marker' step


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.2.1" target="_blank">v3.2.1</a> - 2022-08-04
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Added
- New FAQ: Why does the text color of my signature change sometimes?
### Fixed
- The permission check no longer takes more time than necessary by showing all allow or deny reasons, only the first match. Denies are only evaluated when an allow match has been found before.
- Template file categorization time no longer grows exponentially with each template appearing multiple times in an INI file
- Handle nested attribute names in graph config file correctly ('onPremisesExtensionAttributes.extensionAttribute1' et al.) (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/41" target="_blank">#41</a>) (Thanks <a href="https://github.com/dakolta" target="_blank">@dakolta</a>!)
- Handle INI files with only one section correctly (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/42" target="_blank">#42</a>) (Thanks <a href="https://github.com/dakolta" target="_blank">@dakolta</a>!)
- Include 'state' in list of default replacement variables (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/44" target="_blank">#44</a>) (Thanks <a href="https://github.com/dakolta" target="_blank">@dakolta</a>!)
- The code detecting Outlook and Word registry version, file version and bitness has been corrected


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.2.0" target="_blank">v3.2.0</a> - 2022-07-19
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Added
- Workaround for Word ignoring manual line breaks (`` `n ``) and paragraph marks (`` `r`n ``) in replacement variables when converting a template to a signature in RTF format (signatures in HTM and TXT formats are not affected).
- Sample script `Test-ADTrust.ps1` to test the connection to all Domain Controllers and Global Catalog server of a trusted domain


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.1.0" target="_blank">v3.1.0</a> - 2022-06-26
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Changed
- Each template reference in an INI file is now considered individually, not just the last entry. See 'How to work with INI files' in README for a use case example.
- Additional output is now fully available in the verbose stream, and no longer scattered around the debug and the verbose streams
- Rewrite FAQ "Why is dynamic group membership not considered on premises?" to reflect recent substantial changes in Microsoft Graph, which make Set-OutlookSignatures automatically support dynamic groups in the cloud. See the FAQ in README for more details and the reason why dynamic groups are not supported on premises.
- Extend FAQ "How to avoid blank lines when replacement variables return an empty string?" with new examples and sample code that automatically differentiates between DOCX and HTM templates
- Optimized format of 'hashes.txt'
### Fixed
- Convert SharePoint document library paths to a PowerShell compatible format before accessing them. (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/discussions/38" target="_blank">#38</a>) (Thanks <a href="https://github.com/Johan-Claesson" target="_blank">@Johan-Claesson</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v3.0.0" target="_blank">v3.0.0</a> - 2022-04-20
_This major release brings several changes which can make it incompatible with previous versions. Pay special attention to the changes marked '**Breaking:**' to find out if your environment is affected and what to do._
### Added
- New FAQ: How can I get more script output for troubleshooting?
### Changed
- **Breaking:** All input files of type .htm, .ini and .ps1 are now expected to be UTF8 encoded.  
If you copied and/or modified the sample files delivered with earlier versions of Set-OutlookSignatures, no changes should be necessary as these were delivered in UTF8 already. Please check the encoding anyway.
- The following data of the currently processed mailbox is no longer displayed in the standard output stream but in the verbose output stream: 
  - List of group membership security identifiers (SIDs)
  - List of SMTP addresses
  - Final data of replacement variables
- Update documentation to make clear that 'DNS or NetBIOS name of AD domain' and 'Example' are just examples which need to be replaced with actual AD domain names, but 'EntraID' and 'AzureAD' are not examples
### Removed
- **Breaking:** File name based tags are no longer supported. Use INI files instead.  
This change has been announced with the release of v2.5.0 on 2022-01-14.
- **Breaking:** Parameter AdditionalSignaturePathFolder is no longer supported. Just append the folder to the AdditionalSignaturePath parameter.
- All sample files with tags based on file names have been removed


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.5.2" target="_blank">v2.5.2</a> - 2022-02-09
### Fixed
- Use another Windows API to get the Active Directory object of the logged in user. This API also works when 'CN=Computers,DC=[…]' does not exist or the logged in user does not have read access to it. (Thanks <a href="https://www.linkedin.com/in/mariandanisek/" target="_blank">Marián Daníšek</a>!)
- Correct handle objectSid and SidHistory returned from Graph. The format is no longer a byte array as from on-prem Active Directory, but a list of clear text strings ('S-1-[…]').
- Validate SimulateUser and SimulateMailboxes input


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.5.1" target="_blank">v2.5.1</a> - 2022-01-20
### Fixed
- Fix search for mailbox user object across trusts


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.5.0" target="_blank">v2.5.0</a> - 2022-01-14
### Added
- New parameter DeleteScriptCreatedSignaturesWithoutTemplate, see README for details
- New parameter EmbedImagesInHtml, see README for details
- Tags can now not only be used to allow access to a template, but also to deny access. Denies are available for time, group and email based tags. See README for details.
- Consider distribution group membership in addition to security group membership
- Consider sIDHistory in searches across trusts and when comparing msExchMasterAccountSid, which adds support for scenarios in which a mailbox or a linked account has been migrated between Active Directory domains/forests
- Show matching allow and deny tags for each mailbox-template-combination. This makes it easy to find out why a certain template is applied for a certain mailbox and why not.
- Show which tags lead to a classification as time based, common, group based or email address specific template
- New FAQ: Why is membership in dynamic distribution groups and dynamic security groups not considered?
- New FAQ: Why is no admin or user GUI available?
### Fixed
- Don't throw an error when UseHtmTemplates is set to true and OOFIniFile is used, but there is no \*.htm file in OOFTemplatePath
- Correct mapping of Graph businessPhones attribute, so the replacement variable `$Current[…]Telephone$` is populated (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/26" target="_blank">#26</a>)  (Thanks <a href="https://github.com/vitorpereira" target="_blank">@vitorpereira</a>!)
- Fix Outlook 2013 registry key handling and temporary folder handling in environments without Outlook or Outlook profile (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/27" target="_blank">#27</a>)  (Thanks <a href="https://github.com/Imaginos" target="_blank">@Imaginos</a>!)
### Changed
- Cache group SIDs across all types of templates to reduce network load and increase script speed
- Deprecate file name based tags. They work as-is, no new features will be added and support for file name based tags will be removed completely in the next months. Please switch to INI files, see README for details.
- Update usage examples in script


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.4.0" target="_blank">v2.4.0</a> - 2021-12-10
### Changed
- Documentation updates
- Updated FAQ: What about the new signature roaming feature Microsoft announced?
- When connecting to Microsoft Graph, the TenantID is no longer set to 'organizations', but extracted from the logged in or simulated user name
- Reduce number of required Graph authentication prompts by using a token cache file
- Switching to the EWS Managed API .Net Standard port from https://github.com/ststeiger/RedmineMailService (the official Microsoft DLL is used with Windows PowerShell, ststeiger's port when run in PowerShell 7+) (Thanks <a href="https://github.com/ststeiger" target="_blank">@ststeiger</a>!)
- When saving a document in Word fails, wait for two seconds and retry saving to avoid problems with virus scanners
### Added
- Added sample code files ('.\sample code'), including a wrapper script for central creation and deployment of signatures and OOF messages without end user or client involvement
- New default replacement variables for displayName and mailNickname (a.k.a. alias)
- New parameter GraphOnly: Try to connect to Microsoft Graph only, ignoring any local Active Directory. The default behavior without GraphOnly is unchanged (try Active Directory first, fall back to Graph).
- New parameters CreateRtfSignatures and CreateTxtSignatures allow to disable RTF/TXT signature creation
- New parameter SimulateAndDeployGraphCredentialFile
- New FAQ: How to deploy signatures for "Send As", "Send On Behalf" etc.?
- New FAQ: Can I centrally manage and deploy Outook stationery with this script?
- Report templates that are mentioned in the INI file but do not exist in the file system, and vice versa
### Fixed
- Do not ignore remote mailboxes when searching mailboxes in Active Directory (Thanks <a href="https://www.linkedin.com/in/lwhdk/" target="_blank">Lars Würtz Hammer</a>!)
- Correctly handle hybrid scenarios with basic auth disabled in the cloud (Thanks <a href="https://www.linkedin.com/in/lwhdk/" target="_blank">Lars Würtz Hammer</a>!)
- Correctly handle time based tags, so they are not checked twice (the first check is positive, the second one returns 'unknown tag')


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.3.1" target="_blank">v2.3.1</a> - 2021-11-05
### Fixed
- Ignore mail-enabled users an mailbox search to avoid binding to the wrong Exchange object in migration scenarios (which would lead to wrong replacement variable data and group membership)
- When connecting to Exchange Online, check for valid mailbox in addition to valid credentials
- Clarify port requirements and group membership evaluation in documentation


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.3.0" target="_blank">v2.3.0</a> - 2021-10-08
### Changed
- The parameter DomainsToCheckForGroups is also available under the more descriptive name TrustsToCheckForGroups. Both names can be used, functionality is unchanged.
- Contribution opportunities in '.\docs\CONTRIBUTING'
### Added
- Support for mailboxes in Microsoft 365, including hybrid and cloud only scenarios (see '.\config\default graph config.ps1' for details)
- Possibility to use INI files instead of file name tags, including settings for template sort order, sort culture, and custom Outlook signature names (see parameters 'SignatureIniFile' and 'OOFIniFile' for details)
- New default replacement variables `$Current[…]Office$` and `$Current[…]Company$`, including updated templates
- Enterprise ready workaround for Word security warning when converting documents with linked images
- FAQ: the software hangs at HTM/RTF export, Word shows a security warning!?
- FAQ: Isn't a plural noun in the software name against PowerShell best practices?
- FAQ: How to avoid empty lines when replacement variables return an empty string?
- FAQ: Is there a roadmap for future versions?
- Code of Conduct (see '.\docs\CODE_OF_CONDUCT' for details)
### Fixed
- User could connect to hidden Word instance used for conversion of DOCX templates
- Do no classify templates with unknown tags as common templates 
- Word settings temporarily changed by the software are now also restored to their original values when the software ends due to an unexpected error
- Do not try to change read-only Word attributes \<image>.hyperlink.name and \<image>.hyperlink.addressold (regression bug)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.2.1" target="_blank">v2.2.1</a> - 2021-09-15
### Fixed
- Allow multi-relative paths (Example: 'c:\a\b\x\\..\c\y\z\\..\\..\d' -> 'c:\a\b\c\d')


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.2.0" target="_blank">v2.2.0</a> - 2021-09-15
### Changed
- Make script compatible with PowerShell Core 7.x on Windows (Linux and MacOS are not supported yet)
- Reduce and speed up Active Directory queries by only accepting input in the 'Domain\User' or UPN (User Principal Name) format for the 'SimulateUser' parameter
- Reduce and speed up Active Directory queries by only accepting email addresses as input for the 'SimulateMailboxes' parameter
- Revise repository structure, as well as the process for development, build and release
### Added
- Full support for Exchange mailboxes added in Outlook as POP3 or IMAP4 accounts
- Add FAQs: "Where can I find the changelog?", "How can I contribute, propose a new feature or file a bug?", "What is the recommended approach for custom configuration files?"
- Add file hash of build artifacts to release information and hashes.txt 
- Add dark mode support and badges to documentation files
### Fixed
- Do not show an error message when no default Outlook profile is configured
- Avoid additional blank lines at the end of TXT signature files when DOCX templates are used (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/13" target="_blank">#13</a>)
- Detect user's domain correctly when user and computer belong to different AD forests
- Do not show an error message when only external or interal OOF message is set
- Set current user's OWA signature even when mailbox is in another Outlook profile than the default one


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.1.2" target="_blank">v2.1.2</a> - 2021-09-03
### Fixed
- Correct extension attributes being shown as empty in replacement variables (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/11" target="_blank">#11</a>) (Thanks <a href="https://github.com/goranko73" target="_blank">@goranko73</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.1.1" target="_blank">v2.1.1</a> - 2021-08-26
### Changed
- Disable positional binding of passed arguments for easier debugging.
- Rename '\bin\licenses.txt' to '\bin\LICENSE.txt'
### Added
- "implementation approach.html" describes the recommended approach for implementing the software, based on real-life experience implementing the software in multi-client environments with a five-digit number of mailboxes.
- New FAQ "How to create a shortcut to the software with parameters?"
- New FAQ "What is the recommended approach for implementing the software?"
- Add multi-client capability hint to script description and readme file


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.1.0" target="_blank">v2.1.0</a> - 2021-08-13
### Changed
- Enhance long file path handling
- Enhance FullLanguage mode detection
### Added
- FAQ: How do I start the software from the command line or a scheduled task?
- Added command line and task scheduler example to script
- Logo and icon files are now part of the download package


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.0.2" target="_blank">v2.0.2</a> - 2021-07-23
### Changed
- Inform the user when an Active Directory search returns less or more than one result
- Readme chapter "Simulation mode" updated
### Added
- Readme FAQ "Can multiple script instances run in parallel?"
### Fixed
- Readme link about MS Word ExportPictureWithMetafile registry key to avoid huge RTF files supplemented by alternate link to Internet Archive Wayback Machine (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/9" target="_blank">#9</a>) (Thanks <a href="https://github.com/nitishkanu820" target="_blank">@nitishkanu820</a>!)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.0.1" target="_blank">v2.0.1</a> - 2021-07-22
_Do not use this release. It was withdrawn due to a severe problem._


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v2.0.0" target="_blank">v2.0.0</a> - 2021-07-21
### Changed
- **Breaking:** The configuration file is no longer a plain text file, but a full PowerShell script. This allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
- When enumerating and categoring templates, category defining information is logged (assigned mail address, group name including SID)
- Advanced primary mailbox detection and sorting
- Easier readable output (s cript start and end times are shown, logical grouping of tasks (Outlook tasks first, basic AD tasks second, template enumeration third, then signature handling tasks), vertical whitespace between main tasks makes output easier consumable visually)
### Added
- Simulation mode
- Major script steps now show a timestamp
- Readme for "Simulation mode" and FAQ "How can I log the software output?"
### Fixed
- Templates with multiple groups or multiple mail addresses were not applied correctly and led to redundant Active Directory queries


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.6.1" target="_blank">v1.6.1</a> - 2021-06-30
### Fixed
- Empty AdditionalSignaturePath leads to error (<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/8" target="_blank">#8</a>)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.6.0" target="_blank">v1.6.0</a> - 2021-06-26
### Changed
- Change template path structure
- Update readme
- Update templates
### Added
- Add support for HTML template files (.htm)
- Add parameter UseHtmlTemplates to switch from DOCX to HTML template handling
- Consider images in .docx templates with different text wrapping setting (Shapes for "in line with text" and InlineShapes for all other text wrapping settings)
### Fixed
- Check existence of signature file before trying to set Outlook web signature


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.5.4" target="_blank">v1.5.4</a> - 2021-06-24
### Added
- New FAQ: Why DOCX as template format and not HTML? Signatures in Outlook sometimes look different than my DOCX templates.
### Fixed
- Fix: Consider images with different text wrapping setting (Shapes for "in line with text" and InlineShapes for all other text wrapping settings).


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.5.3" target="_blank">v1.5.3</a> - 2021-06-23
### Fixed
- Fix problem connecting to SharePoint document libraries.
- Fix readme file to reflect enhanced SharePoint document library possibilities in path parameters.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.5.2" target="_blank">v1.5.2</a> - 2021-06-21
### Fixed
- Fix handling of Outlook Web connection error.
- Fix readme: OOF templates are not applied when currently active or scheduled.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.5.1" target="_blank">v1.5.1</a> - 2021-06-20
## Fixed
- Provide readme.html in releases, not markdown file
- Update link formatting in readme files
- Add attribution for logo source
- Update logo path and dependencies


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.5.0" target="_blank">v1.5.0</a> - 2021-06-18
### Added
- Add support for out-of-office replies
- New parameter SetCurrentUserOOFMessage
- New parameter OOFTemplatePath
- Add sample files for OOF templates '.\OOF templates'


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.4.0" target="_blank">v1.4.0</a> - 2021-06-17
### Added
- New parameter AdditionalSignaturePath


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.3.0" target="_blank">v1.3.0</a> - 2021-06-16
### Added
- New parameter DeleteUserCreatedSignatures
- New parameter SetCurrentUserOutlookWebSignature


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.2.1" target="_blank">v1.2.1</a>  2021-06-14
### Fixed
- Fix signature group name to SID mapping
- Make logo work with every background (transparency, white glow)


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.2.0" target="_blank">v1.2.0</a> - 2021-06-11
### Changed
- Reduce LDAP queries by getting replacement variable data per mailbox, not per signature file and mailbox
- Speed up variable replacement in image metadata
### Added
- Show replacement variable values in output
- Show variables and script root in output
- Show warning when replacement variable config file cannot be accessed
- Update signature template file 'Test all signature replacement variables.docx'
- Include info about case sensitivity in file 'default replacement variables.txt'

## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.1.0" target="_blank">v1.1.0</a> - 2021-06-10
### Changed
-Move all replacement variable definitions to './config/default replacement variables.txt'
-Update README: Add logo, modify chapter ordering, document parameter ReplacementVariableConfigFile
### Added
- Add Exchange Extension variables 1..15 to './config/default replacement variables.txt' and 'Test all signature replacement variables.docx'
- Create subdirectories for binaries and configurations, adapt script to work with new subdirectories
- Add a logo
- Adapt script to include script information (version and others) in code and output
### Fixed
- Modify license.txt to that GitHub recognizes the license type
- Add '.gitattributes' file to ignore '.git*' folders and README in relase
- Add 'readme.txt', a plain text version of README, which can be read on all systems with on-board tools.


## <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/tag/v1.0.0" target="_blank">v1.0.0</a> - 2021-06-01
_Initial release._


## v0.1.0 - 2021-04-21
_First lines of code were written as proof of concept, but never published._
