<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="../src/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a>**<br>Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages.<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Set-OutlookSignatures" alt=""></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Set-OutlookSignatures/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a>  

# Changelog
<!--
  Sample changelog entry
  Remove leading spaces after pasting
  ## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/vX.X.X" target="_blank">vX.X.X</a> - YYYY-MM-DD
  _Put Notice here_
  _**Breaking:** Notice about breaking change_  
  ### Changed
  - **Breaking:** XXX
  ### Added
  ### Removed
  ### Fixed
-->


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/vX.X.X" target="_blank">vX.X.X</a> - YYYY-MM-DD
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Fixed
- Mailbox priority list: Don't add automapped or additional mailboxes to the end of the mailbox priority list, but to the end of each Outlook profile in the mailbox priority list
- Mailbox priority list: When duplicates exist, only show the mail address with the highest priority. Verbose output contains additional information for each occurrence (Outlook profile name, registry path, legacyExchangeDN)
- Setting default signature: Show Outlook profile name, so that mailboxes that exist in multiple Outlook profiles can be distinguished
- When the logged in user's personal mailbox exists in multiple profiles, set the Outlook Web signature and OOF message only for this mailbox only once
- When a mailbox exists in multiple Outlook profiles, only query AD/Graph at the first occurrence and use cached data on remaining occurrences

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.4.1" target="_blank">v3.4.1</a> - 2022-11-25
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Fixed
- Correctly handle logged in user with empty mail attribute
- Correctly enumerate SID and SidHistory when connected to a local Active Directory user with a mailbox in Exchange Online (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/59" target="_blank">#59</a>) (Thanks <a href="https://github.com/AnotherFranck" target="_blank">@AnotherFranck</a>!)
- Correctly handle empty templates, signatures and OOF messages

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.4.0" target="_blank">v3.4.0</a> - 2022-11-02
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Added
- New parameter '`IncludeMailboxForestDomainLocalGroups`', see '`README`' for details
- LDAP and Global Catalog connectivity is now additionally checked for every child domain of the current user's Active Directory forest and every child domain of cross-forest trusts
- Consider SID history of groups in trusted domains/forests
- New FAQs in '`.\docs\README`': '`What if Outlook is not installed at all?`' and '`What if a user has no Outlook profile or is prohibited from starting Outlook?`'
### Fixed
- Correctly calculate mailbox priority when simulation mode is enabled and/or the e-mail address is a secondary address
- On-prem: Membership in domain local groups is now recognized if the group is in a child domain of a forest connected with a cross-forest trust
- Only consider mailboxes as additional mailboxes when they appear in Outlook's list in the e-mail navigation pane. This avoids falsely adding shared calendars as additional mailboxes.

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.3.0" target="_blank">v3.3.0</a> - 2022-09-05
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Changed
- Use different method to delete files to avoid occassional OneDrive error "access to the cloud file is denied"
- Update logo and icon
### Added
- The script now detects not only primary mailboxes configured in Outlook, but also automapped and additional mailboxes. This behavior can be disabled with the new parameter '`SignaturesForAutomappedAndAdditionalMailboxes`'. See '`README`' for details.

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.2.2" target="_blank">v3.2.2</a> - 2022-08-12
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Fixed
- When the '`EmbedImagesInHtml`' parameter is set to '`false`', correctly handle a certain file system condition instead of stopping processing the current template after the 'Embed local files in HTM format and add marker' step

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.2.1" target="_blank">v3.2.1</a> - 2022-08-04
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Added
- New FAQ: Why does the text color of my signature change sometimes?
### Fixed
- The permission check no longer takes more time than necessary by showing all allow or deny reasons, only the first match. Denies are only evaluated when an allow match has been found before.
- Template file categorization time no longer grows exponentially with each template appearing multiple times in an ini file
- Handle nested attribute names in graph config file correctly ('onPremisesExtensionAttributes.extensionAttribute1' et al.) (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/41" target="_blank">#41</a>) (Thanks <a href="https://github.com/dakolta" target="_blank">@dakolta</a>!)
- Handle ini files with only one section correctly (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/42" target="_blank">#42</a>) (Thanks <a href="https://github.com/dakolta" target="_blank">@dakolta</a>!)
- Include 'state' in list of default replacement variables (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/44" target="_blank">#44</a>) (Thanks <a href="https://github.com/dakolta" target="_blank">@dakolta</a>!)
- The code detecting Outlook and Word registry version, file version and bitness has been corrected

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.2.0" target="_blank">v3.2.0</a> - 2022-07-19
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Added
- Workaround for Word ignoring manual line breaks ('`` `n ``') and paragraph marks ('`` `r`n ``') in replacement variables when converting a template to a signature in RTF format (signatures in HTM and TXT formats are not affected).
- Sample script '`Test-ADTrust.ps1`' to test the connection to all Domain Controllers and Global Catalog server of a trusted domain

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.1.0" target="_blank">v3.1.0</a> - 2022-06-26
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Changed
- Each template reference in an INI file is now considered individually, not just the last entry. See 'How to work with ini files' in README for a usecase example.
- Additional output is now fully available in the verbose stream, and no longer scattered around the debug and the verbose streams
- Rewrite FAQ "Why is dynamic group membership not considered on premises?" to reflect recent substantial changes in Microsoft Graph, which make Set-OutlookSignatures automatically support dynamic groups in the cloud. See the FAQ in README for more details and the reason why dynamic groups are not supported on premises.
- Extend FAQ "How to avoid blank lines when replacement variables return an empty string?" with new examples and sample code that automatically differentiates between DOCX and HTM templates
- Optimized format of 'hashes.txt'
### Fixed
- Convert WebDAV paths to a PowerShell compatible format before accessing them. (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/discussions/38" target="_blank">#38</a>) (Thanks <a href="https://github.com/Johan-Claesson" target="_blank">@Johan-Claesson</a>!)

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.0.0" target="_blank">v3.0.0</a> - 2022-04-20
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
- Update documentation to make clear that 'NetBiosDomain' and 'Example' are just examples which need to be replaced with actual NetBIOS domain names, but 'AzureAD' is not an example
### Removed
- **Breaking:** File name based tags are no longer supported. Use ini files instead.  
This change has been announced with the release of v2.5.0 on 2022-01-14.
- **Breaking:** Parameter AdditionalSignaturePathFolder is no longer supported. Just append the folder to the AdditionalSignaturePath parameter.
- All sample files with tags based on file names have been removed

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.5.2" target="_blank">v2.5.2</a> - 2022-02-09
### Fixed
- Use another Windows API to get the Active Directory object of the logged in user. This API also works when 'CN=Computers,DC=[...]' does not exist or the logged in user does not have read access to it. (Thanks <a href="https://www.linkedin.com/in/mariandanisek/" target="_blank">Marián Daníšek</a>!)
- Correct handle objectSid and SidHistory returned from Graph. The format is no longer a byte array as from on-prem Active Directory, but a list of clear text strings ('S-1-[...]').
- Validate SimulateUser and SimulateMailboxes input

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.5.1" target="_blank">v2.5.1</a> - 2022-01-20
### Fixed
- Fix search for mailbox user object across trusts

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.5.0" target="_blank">v2.5.0</a> - 2022-01-14
### Added
- New parameter DeleteScriptCreatedSignaturesWithoutTemplate, see README for details
- New parameter EmbedImagesInHtml, see README for details
- Tags can now not only be used to allow access to a template, but also to deny access. Denies are available for time, group and e-mail based tags. See README for details.
- Consider distribution group membership in addition to security group membership
- Consider sIDHistory in searches across trusts and when comparing msExchMasterAccountSid, which adds support for scenarios in which a mailbox or a linked account has been migrated between Active Directory domains/forests
- Show matching allow and deny tags for each mailbox-template-combination. This makes it easy to find out why a certain template is applied for a certain mailbox and why not.
- Show which tags lead to a classification as time based, common, group based or e-mail address specific template
- New FAQ: Why is membership in dynamic distribution groups and dynamic security groups not considered?
- New FAQ: Why is no admin or user GUI available?
### Fixed
- Don't throw an error when UseHtmTemplates is set to true and OOFIniFile is used, but there is no \*.htm file in OOFTemplatePath
- Correct mapping of Graph businessPhones attribute, so the replacement variable `$CURRENT[...]TELEPHONE$` is populated (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/26" target="_blank">#26</a>)  (Thanks <a href="https://github.com/vitorpereira" target="_blank">@vitorpereira</a>!)
- Fix Outlook 2013 registry key handling and temporary folder handling in environments without Outlook or Outlook profile (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/27" target="_blank">#27</a>)  (Thanks <a href="https://github.com/Imaginos" target="_blank">@Imaginos</a>!)
### Changed
- Cache group SIDs across all types of templates to reduce network load and increase script speed
- Deprecate file name based tags. They work as-is, no new features will be added and support for file name based tags will be removed completely in the next months. Please switch to ini files, see README for details.
- Update usage examples in script

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.4.0" target="_blank">v2.4.0</a> - 2021-12-10
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
- New parameter GraphCredentialFile
- New FAQ: How to deploy signatures for "Send As", "Send On Behalf" etc.?
- New FAQ: Can I centrally manage and deploy Outook stationery with this script?
- Report templates that are mentioned in the ini file but do not exist in the file system, and vice versa
### Fixed
- Do not ignore remote mailboxes when searching mailboxes in Active Directory (Thanks <a href="https://www.linkedin.com/in/lwhdk/" target="_blank">Lars Würtz Hammer</a>!)
- Correctly handle hybrid scenarios with basic auth disabled in the cloud (Thanks <a href="https://www.linkedin.com/in/lwhdk/" target="_blank">Lars Würtz Hammer</a>!)
- Correctly handle time based tags, so they are not checked twice (the first check is positive, the second one returns 'unknown tag')

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.3.1" target="_blank">v2.3.1</a> - 2021-11-05
### Fixed
- Ignore mail-enabled users an mailbox search to avoid binding to the wrong Exchange object in migration scenarios (which would lead to wrong replacement variable data and group membership)
- When connecting to Exchange Online, check for valid mailbox in addition to valid credentials
- Clarify port requirements and group membership evaluation in documentation

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.3.0" target="_blank">v2.3.0</a> - 2021-10-08
### Changed
- The parameter TrustsToCheckForGroups is also available under the more descriptive name TrustsToCheckForGroups. Both names can be used, functionality is unchanged.
- Contribution opportunities in '.\docs\CONTRIBUTING.html'
### Added
- Support for mailboxes in Microsoft 365, including hybrid and cloud only scenarios (see '.\docs\README.html' and '.\config\default graph config.ps1' for details)
- Possibility to use ini files instead of file name tags, including settings for template sort order, sort culture, and custom Outlook signature names (see parameters 'SignatureIniPath' and 'OOFIniPath' for details)
- New default replacement variables `$CURRENT[...]OFFICE$` and `$CURRENT[...]COMPANY$`, including updated templates
- Enterprise ready workaround for Word security warning when converting documents with linked images
- FAQ: The script hangs at HTM/RTF export, Word shows a security warning!?
- FAQ: Isn't a plural noun in the script name against PowerShell best practices?
- FAQ: How to avoid empty lines when replacement variables return an empty string?
- FAQ: Is there a roadmap for future versions?
- Code of Conduct (see '.\docs\CODE_OF_CONDUCT.html' for details)
### Fixed
- User could connect to hidden Word instance used for conversion of DOCX templates
- Do no classify templates with unknown tags as common templates 
- Word settings temporarily changed by the script are now also restored to their original values when the script ends due to an unexpected error
- Do not try to change read-only Word attributes \<image>.hyperlink.name and \<image>.hyperlink.addressold (regression bug)

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.2.1" target="_blank">v2.2.1</a> - 2021-09-15
### Fixed
- Allow multi-relative paths (Example: 'c:\a\b\x\\..\c\y\z\\..\\..\d' -> 'c:\a\b\c\d')

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.2.0" target="_blank">v2.2.0</a> - 2021-09-15
### Changed
- Make script compatible with PowerShell Core 7.x on Windows (Linux and MacOS are not supported yet)
- Reduce and speed up Active Directory queries by only accepting input in the 'Domain\User' or UPN (User Principal Name) format for the 'SimulateUser' parameter
- Reduce and speed up Active Directory queries by only accepting e-mail addresses as input for the 'SimulateMailboxes' parameter
- Revise repository structure, as well as the process for development, build and release
### Added
- Full support for Exchange mailboxes added in Outlook as POP3 or IMAP4 accounts
- Add FAQs: "Where can I find the changelog?", "How can I contribute, propose a new feature or file a bug?", "What is the recommended approach for custom configuration files?"
- Add file hash of build artifacts to release information and hashes.txt 
- Add dark mode support and badges to documentation files
### Fixed
- Do not show an error message when no default Outlook profile is configured
- Avoid additional blank lines at the end of TXT signature files when DOCX templates are used (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/13" target="_blank">#13</a>)
- Detect user's domain correctly when user and computer belong to different AD forests
- Do not show an error message when only external or interal OOF message is set
- Set current user's OWA signature even when mailbox is in another Outlook profile than the default one

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.1.2" target="_blank">v2.1.2</a> - 2021-09-03
### Fixed
- Correct extension attributes being shown as empty in replacement variables (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/11" target="_blank">#11</a>) (Thanks <a href="https://github.com/goranko73" target="_blank">@goranko73</a>!)

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.1.1" target="_blank">v2.1.1</a> - 2021-08-26
### Changed
- Disable positional binding of passed arguments for easier debugging.
- Rename '\bin\licenses.txt' to '\bin\LICENSE.txt'
### Added
- "implementation approach.html" describes the recommended approach for implementing the software, based on real-life experience implementing the script in a multi-client environment with a five-digit number of mailboxes.
- New FAQ "How to create a shortcut to the script with parameters?"
- New FAQ "What is the recommended approach for implementing the software?"
- Add multi-client capability hint to script description and readme file

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.1.0" target="_blank">v2.1.0</a> - 2021-08-13
### Changed
- Enhance long file path handling
- Enhance FullLanguage mode detection
### Added
- FAQ: How do I start the script from the command line or a scheduled task?
- Added command line and task scheduler example to script
- Logo and icon files are now part of the download package

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.0.2" target="_blank">v2.0.2</a> - 2021-07-23
### Changed
- Inform the user when an Active Directory search returns less or more than one result
- Readme chapter "Simulation mode" updated
### Added
- Readme FAQ "Can multiple script instances run in parallel?"
### Fixed
- Readme link about MS Word ExportPictureWithMetafile registry key to avoid huge RTF files supplemented by alternate link to Internet Archive Wayback Machine (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/9" target="_blank">#9</a>) (Thanks <a href="https://github.com/nitishkanu820" target="_blank">@nitishkanu820</a>!)

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.0.1" target="_blank">v2.0.1</a> - 2021-07-22
_Do not use this release. It was withdrawn due to a severe problem._

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.0.0" target="_blank">v2.0.0</a> - 2021-07-21
### Changed
- **Breaking:** The configuration file is no longer a plain text file, but a full PowerShell script. This allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
- When enumerating and categoring templates, category defining information is logged (assigned mail address, group name including SID)
- Advanced primary mailbox detection and sorting
- Easier readable output (s cript start and end times are shown, logical grouping of tasks (Outlook tasks first, basic AD tasks second, template enumeration third, then signature handling tasks), vertical whitespace between main tasks makes output easier consumable visually)
### Added
- Simulation mode
- Major script steps now show a timestamp
- Readme for "Simulation mode" and FAQ "How can I log the script output?"
### Fixed
- Templates with multiple groups or multiple mail addresses were not applied correctly and led to redundant Active Directory queries

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.6.1" target="_blank">v1.6.1</a> - 2021-06-30
### Fixed
- Empty AdditionalSignaturePath leads to error (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/8" target="_blank">#8</a>)

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.6.0" target="_blank">v1.6.0</a> - 2021-06-26
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

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.4" target="_blank">v1.5.4</a> - 2021-06-24
### Added
- New FAQ: Why DOCX as template format and not HTML? Signatures in Outlook sometimes look different than my DOCX templates.
### Fixed
- Fix: Consider images with different text wrapping setting (Shapes for "in line with text" and InlineShapes for all other text wrapping settings).

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.3" target="_blank">v1.5.3</a> - 2021-06-23
### Fixed
- Fix problem connecting to WebDAV-paths.
- Fix readme file to reflect enhanced WebDAV possibilities in path parameters.

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.2" target="_blank">v1.5.2</a> - 2021-06-21
### Fixed
- Fix handling of Outlook Web connection error.
- Fix readme: OOF templates are not applied when currently active or scheduled.

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.1" target="_blank">v1.5.1</a> - 2021-06-20
## Fixed
- Provide readme.html in releases, not markdown file
- Update link formatting in readme files
- Add attribution for logo source
- Update logo path and dependencies

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.0" target="_blank">v1.5.0</a> - 2021-06-18
### Added
- Add support for Out of Office (OOF) auto reply messages
- New parameter SetCurrentUserOOFMessage
- New parameter OOFTemplatePath
- Add sample files for OOF templates '.\OOF templates'

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.4.0" target="_blank">v1.4.0</a> - 2021-06-17
### Added
- New parameter AdditionalSignaturePath

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.3.0" target="_blank">v1.3.0</a> - 2021-06-16
### Added
- New parameter DeleteUserCreatedSignatures
- New parameter SetCurrentUserOutlookWebSignature

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.2.1" target="_blank">v1.2.1</a>  2021-06-14
### Fixed
- Fix signature group name to SID mapping
- Make logo work with every background (transparency, white glow)

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.2.0" target="_blank">v1.2.0</a> - 2021-06-11
## Changed
- Reduce LDAP queries by getting replacement variable data per mailbox, not per signature file and mailbox
- Speed up variable replacement in image metadata
## Added
- Show replacement variable values in output
- Show variables and script root in output
- Show warning when replacement variable config file can not be accessed
- Update signature template file 'Test all signature replacement variables.docx'
- Include info about case sensitivity in file 'default replacement variables.txt'

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.1.0" target="_blank">v1.1.0</a> - 2021-06-10
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

## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.0.0" target="_blank">v1.0.0</a> - 2021-06-01
_Initial release._

## v0.1.0 - 2021-04-21
_First lines of code were written as proof of concept, but never published._
