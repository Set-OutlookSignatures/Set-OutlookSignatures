<!-- omit in toc -->
# <a href="https://github.com/GruberMarkus/Set-OutlookSignatures"><img src="../src/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-outlookSignatures"></a><br>Centrally&nbsp;manage&nbsp;and&nbsp;deploy Outlook&nbsp;text&nbsp;signatures&nbsp;and Out&nbsp;of&nbsp;Office&nbsp;auto&nbsp;reply&nbsp;messages.<br><a href="https://github.com/GruberMarkus/Set-OutlookSignatures/blob/main/license.txt"><img src="https://img.shields.io/github/license/grubermarkus/Set-OutlookSignatures" alt=""></a> <a href="https://www.paypal.com/donate?business=JBM584K3L5PX4&no_recurring=0&currency_code=EUR"><img src="https://img.shields.io/badge/sponsor-grey?logo=paypal" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases"><img src="https://img.shields.io/github/downloads/GruberMarkus/Set-OutlookSignatures/total" alt="" data-external="1"></a> <a href="https://github.com/grubermarkus/set-outlooksignatures/stargazers"><img src="https://img.shields.io/github/stars/grubermarkus/set-outlooksignatures" alt="" data-external="1"></a> <a href="https://github.com/grubermarkus/set-outlooksignatures/issues"><img src="https://img.shields.io/github/issues/grubermarkus/set-outlooksignatures" alt="" data-external="1"></a>  

# Changelog

## [v2.2.0] - 2021-09-08
### Changed
- Make script compatible with PowerShell versions greater than 5.1 (a.k.a PowerShell Core based on .Net Core)
- Revise repository structure, as well as the process for development, build and release
### Added
- Add FAQs: "Where can I find the changelog?", "How can I contribute, propose a new feature or file a bug?"
- Add file hash of build artifacts to release information and hashes.txt 
- Add dark mode support and badges to documentation files
### Fixed
- Do not show an error message when no default Outlook profile is configured
- Avoid additional blank lines at the end of .txt signature files when .doxc templates are used ([#13](https://github.com/GruberMarkus/Set-OutlookSignatures/issues/13))

## [v2.1.2] - 2021-09-03
### Fixed
- Correct extension attributes being shown as empty in replacement variables ([#11](https://github.com/GruberMarkus/Set-OutlookSignatures/issues/11)) ([@goranko73](https://github.com/goranko73))

## [v2.1.1] - 2021-08-26
### Changed
- Disable positional binding of passed arguments for easier debugging.
- Rename '\bin\licenses.txt' to '\bin\LICENSE.txt'
### Added
- "implementation approach.html" describes the recommended approach for implementing the software, based on real-life experience implementing the script in a multi-client environment with a five-digit number of mailboxes.
- New FAQ "How to create a shortcut to the script with parameters?"
- New FAQ "What is the recommended approach for implementing the software?"
- Add multi-client capability hint to script description and readme file

## [v2.1.0] - 2021-08-13
### Changed
- Enhance long file path handling
- Enhance FullLanguage mode detection
### Added
- FAQ: How do I start the script from the command line or a scheduled task?
- Added command line and task scheduler example to script
- Logo and icon files are now part of the download package

## [v2.0.2] - 2021-07-23
### Changed
- Inform the user when an Active Directory search returns less or more than one result
- Readme chapter "Simulation mode" updated
### Added
- Readme FAQ "Can multiple script instances run in parallel?"
### Fixed
- Readme link https://support.microsoft.com/kb/224663 (info about MS Word ExportPictureWithMetafile registry key to avoid huge RTF files) supplemented by alternate link to Internet Archive Wayback Machine. ([#9](https://github.com/GruberMarkus/Set-OutlookSignatures/issues/9)) ([@nitishkanu820](https://github.com/nitishkanu820))

## [v2.0.1] - 2021-07-22
_Do not use this release. It was withdrawn due to a severe problem._

## [v2.0.0] - 2021-07-21
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

## [v1.6.1] - 2021-06-30
### Fixed
- Empty AdditionalSignaturePath leads to error ([#8](https://github.com/GruberMarkus/Set-OutlookSignatures/issues/8))

## [v1.6.0] - 2021-06-26
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

## [v1.5.4] - 2021-06-24
### Added
- New FAQ: Why DOCX as template format and not HTML? Signatures in Outlook sometimes look different than my DOCX templates.
### Fixed
- Fix: Consider images with different text wrapping setting (Shapes for "in line with text" and InlineShapes for all other text wrapping settings).

## [v1.5.3] - 2021-06-23
### Fixed
- Fix problem connecting to WebDAV-paths.
- Fix readme file to reflect enhanced WebDAV possibilities in path parameters.

## [v1.5.2] - 2021-06-21
### Fixed
- Fix handling of Outlook Web connection error.
- Fix readme: OOF templates are not applied when currently active or scheduled.

## [v1.5.1] - 2021-06-20
## Fixed
- Provide readme.html in releases, not readme.md
- Update link formatting in readme files
- Add attribution for logo source
- Update logo path and dependencies

## [v1.5.0] - 2021-06-18
### Added
- Add support for Out of Office (OOF) auto reply messages
- New parameter SetCurrentUserOOFMessage
- New parameter OOFTemplatePath
- Add sample files for OOF templates '.\OOF templates'

## [v1.4.0] - 2021-06-17
### Added
- New parameter AdditionalSignaturePath

## [v1.3.0] - 2021-06-16
### Added
- New parameter DeleteUserCreatedSignatures
- New parameter SetCurrentUserOutlookWebSignature

## [v1.2.1]  2021-06-14
### Fixed
- Fix signature group name to SID mapping
- Make logo work with every background (transparency, white glow)

## [v1.2.0] - 2021-06-11
## Changed
- Reduce LDAP queries by getting replacement variable data per mailbox, not per signature file and mailbox
- Speed up variable replacement in image metadata
## Added
- Show replacement variable values in output
- Show variables and script root in output
- Show warning when replacement variable config file can not be accessed
- Update signature template file 'Test all signature replacement variables.docx'
- Include info about case sensitivity in file 'default replacement variables.txt'

## [v1.1.0] - 2021-06-10
### Changed
-Move all replacement variable definitions to './config/default replacement variables.txt'
-Update readme.md: Add logo, modify chapter ordering, document parameter ReplacementVariableConfigFile
### Added
- Add Exchange Extension variables 1..15 to './config/default replacement variables.txt' and 'Test all signature replacement variables.docx'
- Create subdirectories for binaries and configurations, adapt script to work with new subdirectories
- Add a logo
- Adapt script to include script information (version and others) in code and output
### Fixed
- Modify license.txt to that GitHub recognizes the license type
- Add '.gitattributes' file to ignore '.git*' folders and readme.md in relase
- Add 'readme.txt', a plain text version of readme.md, which can be read on all systems with on-board tools.

## [v1.0.0] - 2021-06-01
_Initial release._

## v0.1.0 - 2021-04-21
_First lines of code were written as proof of concept, but never published._

[v2.2.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.2.0
[v2.1.2]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.1.2
[v2.1.1]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.1.1
[v2.1.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.1.0
[v2.0.2]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.0.2
[v2.0.1]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.0.1
[v2.0.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v2.0.0
[v1.6.1]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.6.1
[v1.6.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.6.0
[v1.5.4]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.4
[v1.5.3]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.3
[v1.5.2]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.2
[v1.5.1]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.1
[v1.5.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.0
[v1.4.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.5.0
[v1.3.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.3.0
[v1.2.1]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.2.1
[v1.2.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.2.0
[v1.1.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.1.0
[v1.0.0]: https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v1.0.0
