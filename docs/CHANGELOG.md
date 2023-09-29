<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a>**<br>Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages<p><p><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt="this release"></a> XXXRemoveWhenBuildingXXX--> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="latest release" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="open issues" data-external="1"></a> <img src="https://img.shields.io/badge/dynamic/json?color=brightgreen&label=views&query=%24.count&url=https%3A%2F%2Fraw.githubusercontent.com%2FGruberMarkus%2Fmy-traffic2badge%2Ftraffic%2Ftraffic-Set-OutlookSignatures%2Ftraffic_views.json" alt="views" data-external="1"> <img src="https://img.shields.io/badge/dynamic/json?color=brightgreen&label=clones&query=%24.count&url=https%3A%2F%2Fraw.githubusercontent.com%2FGruberMarkus%2Fmy-traffic2badge%2Ftraffic%2Ftraffic-Set-OutlookSignatures%2Ftraffic_clones.json" alt="clones" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures?color=brightgreen" alt="stars" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/donate,%20support,%20sponsor-white?logo=githubsponsors" alt="donate or sponsor"></a> <a href="./Benefactor%20Circle.md" target="_blank"><img src="https://img.shields.io/badge/unlock%20all%20features%20with-Benefactor%20Circle-gold" alt="unlock all features with Benefactor Circle"></a>
**A big "Thank you!" for listing, featuring, supporting or sponsoring Set-OutlookSignatures!**
<pre><a href="https://explicitconsulting.at" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Others/ExplicIT Consulting, color on black.png" height="100" title="ExplicIT Consulting" alt="ExplicIT Consulting"></a>  <a href="https://joinup.ec.europa.eu/collection/free-and-open-source-software/solution/set-outlooksignatures/about" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Others/EC Joinup Interoperable Europe.png" height="100" title="European Commission Joinup/Interoperable Europe programs" alt="European Commission Joinup/Interoperable Europe programs"></a>  <a href="https://startups.microsoft.com" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Others/MS_Startups_Celebration_Badge_Dark.png" height="100" title="Proud to partner with Microsoft for Startups" alt="Proud to partner with Microsoft for Startups"></a>  <a href="https://archiveprogram.github.com/" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Others/GitHub-Archive-Program-logo.png" height="100" title="GitHub Archive Program" alt="GitHub Archive Program"></a></pre>

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


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v4.5.0" target="_blank">v4.5.0</a> - 2023-09-29
_**Some features are exclusive to the commercial Benefactor Circle add-on**_  
_See [`.\docs\Benefactor Circle`](Benefactor%20Circle.md) or [`https://explicitonsulting.at`](https://explicitconsulting.at/open-source/set-outlooksignatures) for details about these features and how you can benefit from them with a Benefactor Circle licence._

_**Attention, cloud mailbox users:**_  
_Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._  
_Set-OutlookSignatures can experimentally handle cloud roaming signatures since v4.0.0. See `MirrorLocalSignaturesToCloud` in `.\docs\README` for details._
### Changed
- Adopt program logic to cloud roaming signatures API changes introduced by Microsoft
- Updated FAQ `How can I log the script output?`. See `README` for details.
### Added
- New parameter `EmbedImagesInHtmlAdditionalSignaturePath`. See `README` for details.
- New FAQ `How can I start the script only when there is a connection to the Active Directory on-prem?`. See `README` for details.
### Fixed
- Variables in HTM templates have not been replaced with actual values because of a wrong RegEx syntax
- Content of path defined in `AdditionalSignaturePath` was not deleted before copy operations.


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v4.4.0" target="_blank">v4.4.0</a> - 2023-09-20
_**Some features are exclusive to the commercial Benefactor Circle add-on**_  
_See [`.\docs\Benefactor Circle`](Benefactor%20Circle.md) or [`https://explicitonsulting.at`](https://explicitconsulting.at/open-source/set-outlooksignatures) for details about these features and how you can benefit from them with a Benefactor Circle licence._

_**Attention, cloud mailbox users:**_  
_Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._  
_Set-OutlookSignatures can experimentally handle cloud roaming signatures since v4.0.0. See `MirrorLocalSignaturesToCloud` in `.\docs\README` for details._
### Added
- Support for New Outlook. Mailboxes are taken from the first matching source:
  1. Simulation mode is enabled: Mailboxes defined in SimulateMailboxes
  2. Outlook is installed and has profiles, and New Outlook is not set as default: Mailboxes from Outlook profiles
  3. New Outlook is installed: Mailboxes from New Outlook (including manually added and automapped mailboxes for the currently logged-in user)
  4. If none of the above matches: Mailboxes from Outlook Web (including manually added mailboxes, automapped mailboxes follow when Microsoft updates Outlook Web to match the New Outlook experience)
### Fixed
- Correctly handly write protected files when copying to AdditionalSignaturePath


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v4.3.0" target="_blank">v4.3.0</a> - 2023-09-08
_**Some features are exclusive to the commercial Benefactor Circle add-on.**_
- _See [`.\docs\Benefactor Circle`](Benefactor%20Circle.md) for details about these features and how you can benefit from them with a Benefactor Circle licence._

_**Attention, cloud mailbox users:**_
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._  
- _Set-OutlookSignatures can experimentally handle roaming signatures since v4.0.0. See `MirrorLocalSignaturesToCloud` in `.\docs\README` for details._
### Added
- When no mailboxes are configured in Outlook, additional mailboxes configured in Outlook Web are used. Thanks to our partner [ExplicIT Consulting](https://explicitconsulting.at) for donating this code, enabling another world-first feature and bringing us even closer to supporting the "new Outlook" client (codename "Monarch") in the future!
- Add hint to TLS 1.2 when Entra ID/Azure AD/Graph authentication is not successful (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/85" target="_blank">#85</a>) (Thanks <a href="https://github.com/halatovic" target="_blank">@halatovic</a>!)
- Update '`Quick Start Guide`' in '`README`' file with clearer instructions on how to register the Entra ID/Azure AD app required for hybrid and cloud-only environments


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v4.2.1" target="_blank">v4.2.1</a> - 2023-08-16
_**Some features are exclusive to the commercial Benefactor Circle add-on.** See [`.\docs\Benefactor Circle`](Benefactor%20Circle.md) for details about these features and how you can benefit from them with a Benefactor Circle licence._

_**Attention, cloud mailbox users:**_
- _**Set-OutlookSignatures can now experimentally handle roaming signatures!** See `MirrorLocalSignaturesToCloud` in `.\docs\README` for details._
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._
### Fixed
- MoveCSSInline may not find a dependent DLL on some systems (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/84" target="_blank">#84</a>) (Thanks <a href="https://github.com/panki27" target="_blank">@panki27</a>!)
- An error occurred when a trust of the forest root domain of an on-prem Active Directory to itself was detected (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/83" target="_blank">#83</a>) (Thanks <a href="https://github.com/panki27" target="_blank">@panki27</a>!)


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v4.2.0" target="_blank">v4.2.0</a> - 2023-08-10
_**Some features are exclusive to the commercial Benefactor Circle add-on.** See [`.\docs\Benefactor Circle`](Benefactor%20Circle.md) for details about these features and how you can benefit from them with a Benefactor Circle licence._

_**Attention, cloud mailbox users:**_
- _**Set-OutlookSignatures can now experimentally handle roaming signatures!** See `MirrorLocalSignaturesToCloud` in `.\docs\README` for details._
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._
### Added
- New parameter `MoveCSSInline` to move CSS to inline style attributes, for maximum e-mail client compatibility. This parameter is enabled per default, as a workaround to Microsoft's problem with formatting in Outlook Web (M365 roaming signatures and font sizes, especially).
### Fixed
- Set Word WebOptions in correct order, so that they do not overwrite each other


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v4.1.0" target="_blank">v4.1.0</a> - 2023-07-28
_**Some features are exclusive to the commercial Benefactor Circle add-on.** See [`.\docs\Benefactor Circle`](Benefactor%20Circle.md) for details about these features and how you can benefit from them with a Benefactor Circle licence._

_**Attention, cloud mailbox users:**_
- _**Set-OutlookSignatures can now experimentally handle roaming signatures!** See `MirrorLocalSignaturesToCloud` in `.\docs\README` for details._
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._
### Added
- Templates can now be **assigned to or excluded for specific replacement variables of the current user or the current mailbox**. Thanks to [ExplicIT Consulting](https://explicitconsulting.at) for donating this code!  
See `Template tags and ini files` in `README` for details and examples.  
Use cases:
  - Assign template to a specific mailbox or user, but only if user or mailbox is member in multiple groups at the same time.
  - Assign template to users or mailboxes which have or have not a value in a replacement variable.
  - Every replacement variable can be used: Current user and current mailbox, their managers, or tailored replacement variables defined in a custom replacement variable config file.
- Templates can now be **assigned to or excluded for specific e-mail addresses or groups SIDs of the _mailbox of the current user_**. Thanks to [ExplicIT Consulting](https://explicitconsulting.at) for donating this code!  
See `Template tags and ini files` in `README` for details and examples.  
Use cases:
  - Assign template to a specific mailbox, but not if the _mailbox of the current user_ has a specific e-mail address or is member of a specific group. It does not matter if this personal mailbox is added in Outlook or not.  
This is useful for delegate and boss-secretary scenarios - secretaries get specific delegate template for boss's mailbox, but the boss not. **Combine this with the feature that one template can be used multiple times in the ini file, and you basically only need one template file for all delegate combinations in the company!**
  - Assign a template to the mailbox of a specific logged-in user or deny a template for the mailbox of a specific user, no matter which mailboxes the user has added in Outlook.
- The attribute 'GroupsSIDs' is now also available in the `$AdPropsCurrentUser` replacement variable. It contains all the SIDs of the groups the mailbox of the current user is a member of, which allows for replacement variable content based on group membership, as well as assigning or denying templates for specific users. See `Delete images when attribute is empty, variable content based on group membership` in `README` for details and examples.
- Replacement variables are no longer case sensitive. This eliminates a common error source and makes replacement variables in template files easier to read.
- New chapter `Proposed template and signature naming convention` in `README` file. Thanks to [ExplicIT Consulting](https://explicitconsulting.at) for donating this piece of documentation!
- Microsoft has renamed Azure AD to Entra ID. Documentation and code have been updated where possible. In configuration files, 'EntraID' and 'AzureAD' are interchangeable.
### Fixed
- The attribute 'GroupsSIDs' is now reliably available in the `$AdPropsCurrentMailbox` replacement variable.
- Correctly log group and e-mail address specific exclusions (only an optical issue, no technical one)


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v4.0.0" target="_blank">v4.0.0</a> - 2023-07-12
_**Some features are exclusive to the commercial Benefactor Circle add-on.** See [`.\docs\Benefactor Circle`](Benefactor%20Circle.md) for details about these features and how you can benefit from them with a Benefactor Circle licence._

_**Attention, cloud mailbox users:**_
- _**Set-OutlookSignatures can now experimentally handle roaming signatures!** See `MirrorLocalSignaturesToCloud` in `.\docs\README` for details._
- _Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._
### Changed
- **Breaking:** Benefactor Circle members have exclusive access to the following features:
  - Prioritized support and feature requests
    - Issues are handled with priority via a Benefactor Circle exclusive email address and a callback option.
    - Requests for new features are checked for feasability with priority.
    - All release upgrades during the licence period are for free, no matter if it is a patch, feature or major release.
  - Script features
    - Time-based campaigns by assigning time range constraints to templates
    - Signatures for automapped and additional mailboxes
    - Set current user Outlook Web signature (classic Outlook Web signature and roaming signature)
    - Download and upload roaming signatures
    - Set current user Out of Office messages
    - Delete signatures created by the script, where the template no longer exists or is no longer assigned
    - Delete user created signatures
    - Additional signature path (when used outside of simulation mode)
    - High resolution images from DOCX templates
  - Additional documentation: Implementation approach
    - The content is based on real-life experiences implementing the script in multi-client environments with a five-digit number of mailboxes.
    - Proven procedures and recommendations for product managers, architects, operations managers, account managers, mail and client administrators. Suited for service providers as well as for clients.
    - It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.
    - Available in English and German.
  - Sample code
    - SimulateAndDeploy.ps1: Deploy signatures without end user interaction, running Set-OutlookSignatures on a server
    - Test-ADTrust.ps1: Detect why a client cannot query Active Directory information
- **Breaking:** The `CreateRTFSignatures` parameter now defaults to `false`, because the RTF format for e-mails is hardly used nowadays.
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
- The attribute 'GroupsSIDs' is now available in the `$CurrentMailbox...]` variable for use with replacement variables. It contains all the SIDs of the groups the current mailbox is a member of, which allows for replacement variable content based on group membership. See `Delete images when attribute is empty, variable content based on group membership` in `README` for details and examples.
- New sample script `SimulationModeHelper.ps1` make simulation mode usage easier. An admin sets the parameters in the script, the content creators execute it and just have to enter the values required for simulation:
  - The user to simulate (mandatory)
  - The mailbox(es) to simulate (optional)
  - The time to simulate (optional)
  - The output path (optional)
- A basic configuration user interface with grouped parameter sets, just run `Show-Command .\Set-OutlookSignatures.ps1` in PowerShell.
- The new template tag `WriteProtect` write protects individual signature files. See `README` for details and restrictions.
- The new script parameter `SimulateTime` allows to use a specific time when running simulation mode, which is handy for testing time-based templates.
- New parameter `DisableRoamingSignatures`. See `README` for details.
- New FAQ `Start Set-OutlookSignatures in hidden/invisible mode`. See `README` for details.
- Show a warning message when setting the Outlook Web signature is not possible because Outlook Web has not been initialized yet, making it impossible to set signature options in Outlook Web without breaking the first log in experience for this mailbox (getting asked for language, timezone, etc.)
- Copy HTM image width and height attributes to style attribute
- Show a warning when a template contains images formatted as non-inline shapes, as these image formatting options may not be supported by Outlook (e.g., behind the text)
- Support for mailboxes in the user's Entra ID/Azure AD tenant with different UPN/user ID and primary SMTP address
- The Word registry key `DontUseScreenDpiOnOpen` is set to `1` automatically, according to Microsoft documentation (see `README` for details). This helps avoid image sizing problems on devices with non-standard DPI settings.
### Fixed
- Simulation mode
  - Simulation mode partly returned data of the currently logged on user instead of data of the simulated user
  - Simulation mode did not prioritize mailbox list correctly
  - When SimulateUser is not defined, but other simulation parameters, the script exits
  - When multiple mailboxes to simulate were passed, only the first one was considered 
- Graph queries did not correctly handle paged results
- Restoring the Word setting 'ShowFieldCodes' now works correctly in more error scenarios
- Benefactor Circle members only: Additional and automapped mailboxes have not been detected reliably
- Categorizing template files is now much faster than before (two seconds instead of two minutes for 250 templates)
- Replacing variables in DOCX templates is now faster than before, as only variables actually being used in the document are replaced
- Realiably remove '$Current[...]Photo$' string from image alt text
- `SimulateAndDeploy.ps1`: Correctly convert HTML image tags with embedded images and additional options
- Display sort order for was not handled correctly when primary smtp address of a mailboxes has been changed after it was already added to Outlook


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.6.1" target="_blank">v3.6.1</a> - 2023-05-22
_Attention cloud mailbox users: Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._
### Fixed
- Signatures created with `DocxHighResImageConversion true` in combination with `EmbedImagesInHtml false` include high resolution image files, but these images could not be displayed because a wrong path was set in the HTM file


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.6.0" target="_blank">v3.6.0</a> - 2023-01-24
_Attention cloud mailbox users: Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `.\docs\README` for details, known problems and workarounds._
### Changed
- Microsoft Information Protection sensitivity labels are now supported when using DOCX templates. See `How to make Set-OutlookSignatures work with Microsoft Information Protection?` in `.\docs\README` for details.
- Shrinking RTF files is now compatible with Microsoft Information Protection sensitivity labels
- Updated chapter in `.\docs\README`: `Photos from Active Directory`
### Added
- New FAQ in `.\docs\README`: `How to make Set-OutlookSignatures work with Microsoft Information Protection?`
- New parameter `DocxHighResImageConversion`. Enabled by default, this parameter creates HTM signatures with high resolution images from DOCX templates. See `.\docs\README` for details.
### Fixed
- User photo placeholders were not replaced when using HTM templates with images stored in connected sub-folders
- Sample template files `Test all default replacement variables.docx` and `Test all default replacement variables.htm` did not contain all default replacement variables
- Correctly handle empty `AdditionalSignaturePath` parameter


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.5.1" target="_blank">v3.5.1</a> - 2022-12-20
_Attention cloud mailbox users: Microsoft actively enables roaming signatures in Exchange Online. See `What about the roaming signatures feature in Exchange Online?` in `README` for details, known problems and workarounds._
### Fixed
- Use different code to determine Outlook and Word executable file bitness, as the .Net APIs used before seem to fail randomly with the latest Windows and .Net updates (especially when using 32-bit PowerShell on 64-bit Windows)
- Do not stop the script when `SignaturesForAutomappedAndAdditionalMailboxes` is enabled and the Outlook file path can not be determined


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.5.0" target="_blank">v3.5.0</a> - 2022-12-19
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


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.4.1" target="_blank">v3.4.1</a> - 2022-11-25
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Fixed
- Correctly handle logged in user with empty mail attribute
- Correctly enumerate SID and SidHistory when connected to a local Active Directory user with a mailbox in Exchange Online (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/59" target="_blank">#59</a>) (Thanks <a href="https://github.com/AnotherFranck" target="_blank">@AnotherFranck</a>!)
- Correctly handle empty templates, signatures and OOF messages


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.4.0" target="_blank">v3.4.0</a> - 2022-11-02
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Added
- New parameter `IncludeMailboxForestDomainLocalGroups`, see `README` for details
- LDAP and Global Catalog connectivity is now additionally checked for every child domain of the current user's Active Directory forest and every child domain of cross-forest trusts
- Consider SID history of groups in trusted domains/forests
- New FAQs in `.\docs\README`: `What if Outlook is not installed at all?` and `What if a user has no Outlook profile or is prohibited from starting Outlook?`
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
- The script now detects not only primary mailboxes configured in Outlook, but also automapped and additional mailboxes. This behavior can be disabled with the new parameter `SignaturesForAutomappedAndAdditionalMailboxes`. See `README` for details.


## <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases/tag/v3.2.2" target="_blank">v3.2.2</a> - 2022-08-12
_Attention cloud mailbox users: Microsoft will make roaming signatures available in late 2022. See 'What about the roaming signatures feature announced by Microsoft?' in README for details and recommended preparation steps._
### Fixed
- When the `EmbedImagesInHtml` parameter is set to `false`, correctly handle a certain file system condition instead of stopping processing the current template after the 'Embed local files in HTM format and add marker' step


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
- Workaround for Word ignoring manual line breaks (`` `n ``) and paragraph marks (`` `r`n ``) in replacement variables when converting a template to a signature in RTF format (signatures in HTM and TXT formats are not affected).
- Sample script `Test-ADTrust.ps1` to test the connection to all Domain Controllers and Global Catalog server of a trusted domain


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
- Update documentation to make clear that 'NetBiosDomain' and 'Example' are just examples which need to be replaced with actual NetBIOS domain names, but 'EntraID' and 'AzureAD' are not examples
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
- Correct mapping of Graph businessPhones attribute, so the replacement variable `$Current[...]Telephone$` is populated (<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues/26" target="_blank">#26</a>)  (Thanks <a href="https://github.com/vitorpereira" target="_blank">@vitorpereira</a>!)
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
- New default replacement variables `$Current[...]Office$` and `$Current[...]Company$`, including updated templates
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
- "implementation approach.html" describes the recommended approach for implementing the software, based on real-life experience implementing the script in multi-client environments with a five-digit number of mailboxes.
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
### Changed
- Reduce LDAP queries by getting replacement variable data per mailbox, not per signature file and mailbox
- Speed up variable replacement in image metadata
### Added
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