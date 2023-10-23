<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Set-OutlookSignatures%20Benefactor%20Circle%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures Benefactor Circle"></a>**<br>Enhance Set-OutlookSignatures with a great set of additional features and commercial support<p><p><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt="this release"></a> XXXRemoveWhenBuildingXXX--> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="latest release" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="open issues" data-external="1"></a> <img src="https://img.shields.io/badge/dynamic/json?color=brightgreen&label=views&query=%24.count&url=https%3A%2F%2Fraw.githubusercontent.com%2FGruberMarkus%2Fmy-traffic2badge%2Ftraffic%2Ftraffic-Set-OutlookSignatures%2Ftraffic_views.json" alt="views" data-external="1"> <img src="https://img.shields.io/badge/dynamic/json?color=brightgreen&label=clones&query=%24.count&url=https%3A%2F%2Fraw.githubusercontent.com%2FGruberMarkus%2Fmy-traffic2badge%2Ftraffic%2Ftraffic-Set-OutlookSignatures%2Ftraffic_clones.json" alt="clones" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures?color=brightgreen" alt="stars" data-external="1"></a>

# What is the Benefactor Circle?
Benefactor Circle is the result of the partnership between Set-OutlookSignatures and <a href="https://explicitconsulting.at" target="_blank">ExplicIT Consulting</a>.  
<pre><a href="https://explicitconsulting.at" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Others/ExplicIT Consulting, color on black.png" height="100" title="ExplicIT Consulting" alt="ExplicIT Consulting"></a></pre>
ExplicIT Consulting's Benefactor Circle enhances Set-OutlookSignatures with new features and extended support in form of a commercial add-on, ensuring that the core of Set-OutlookSignatures can remain Free and Open-Source Software (FOSS) and continues to evolve.

Visit <a href="https://explicitconsulting.at/open-source/set-outlooksignatures" target="_blank">ExplicIT Consulting's Set-OutlookSignatures Benefactor Circle site</a> for details and pricing information.

# Why choose Set-OutlookSignatures?
- Runs only on your clients, no server side installation
- Mails are not routed through a cloud service, no SPF record change
- Software does not call home
- Works with on-prem, hybrid and cloud-only configurations
- Supports Exchange Online roaming signatures, New Outlook and Outlook Web
- Multi-customer capable
- Works with linked mailboxes in resource forest scenarios
- Users see signature when writing emails
- More cost-effective than other cloud based products, more features than other on-prem products

The features reserved for Benefactor Circle members are available at a very, very competitive price compared to other commercial solutions that work with on-prem, hybrid and cloud-only configurations.

There are also topics you need to be aware of:
- As there is no server component, signatures can not be automatically added to mails sent from mobile devices. This will change as soon as Microsoft's roaming signature feature will be accessible by an API, and mobile applications start using this feature.  
Set-OutlookSignatures Benefactor Circle already supports the roaming signature feature.
- There is no graphical user interface. This is on purpose:
  - End users typically never see the tool, only results.
  - Admins typically need around two hours for the basic setup, as the default parameters are very well chosen and documented.
  - Template maintainers need nothing but Word to create, modify and configure templates.
# Benefactor Circle benefits
- Software features
  - Time-based campaigns by assigning time range constraints to templates
  - Signatures for automapped and additional mailboxes
  - Set current user Outlook Web signature (classic Outlook Web signature and roaming signatures)
  - Download and upload roaming signatures
  - Set current user out of office replies
  - Delete signatures created by the script, where the templates no longer exist or are no longer assigned
  - Delete user created signatures
  - Additional signature path (when used outside of simulation mode)
  - High resolution images from DOCX templates
- Prioritized support and feature requests
  - Issues are handled with priority via a Benefactor Circle exclusive email address and a callback option.
  - Protected web storage allowing a secure upload of log files for analysis.
  - Requests for new features are checked for feasability with priority.
  - All release upgrades during the license period are for free, no matter if it is a patch, feature or major release.
# How can I or my company become a member and obtain a license?
Membership is charged annually in advance.

You receive a Benefactor Circle license file and a corresponding Id, which you can use to unlock exclusive features in Set-OutlookSignatures.

The license file contains the following information:
- Invoice address of the Benefactor Circle member
- End date of the membership
- DNS domain name, SID and maximum number of members for one or multiple license groups

30-day trial licenses are available upon request.
## How much is a membership?
The cost of the membership depends on the number of mailbox licenses included.

Each mailbox, for which an exlusive feature shall be used, needs a license. The mailboxes do not need to be named, you just have to define at least one Active Directory or Graph group containing the mailboxes and the maximum number of recursive members of the groups defined.

Licenses are paid in advance and are valid for one year from the day the full payment is received. There is no automatic renewal.

At the time of the release of this version of Set-OutlookSignatures, the following prices were set by [ExplicIT Consulting](https://explicitconsulting.at). Please check their homepage for updates.

```
The net price in EUR currently is 1.50 € per mailbox and year, with a minimum annual total sum of 100 €.

Yes, that's right: Per year, not per month.

All release upgrades during the license period are for free, no matter if it is a patch, feature or major release.

Support may be chargeable. This includes workshops, implementation support, all forms of remote or on site outsourcing, support for topics already well-explained in the documentation and support for problems with the root cause outside of Set-OutlookSignatures or Set-OutlookSignatures Benefactor Circle.
```
# How do license groups work?
Each Benefactor Circle license is bound to one or more Active Directory or Entra ID/Azure AD groups. Each mailbox of your company needs to be a direct or indirect (a.k.a. nested, recursive or transitive) member of a license group, so that it can receive a signature. Primary group membership is not considered due to Active Directory and Entra ID/Azure AD query restrictions.

Each group may only contain as many mailboxes as direct or indirect members as defined in the license. The user running Set-OutlookSignatures must be able to resolve all direct and indirect members of the license group, even across trusts.

License groups are defined by the DNS domain name of the domain (or 'EntraID' or 'AzureAD' for non-synced groups), their SID (security identifier) and the number of members licensed.
- Use 'EntraID' or 'AzureAD' if the group only exists in Azure Active Directory and is not synced to on-prem. Only one pure Entra ID/Azure AD group is supported, it must be the group with the highest priority (first list entry).
- If you have multiple domains in a forest or multiple forests, you can have multiple license groups, each with a separate maximum member count. For each license, there can be one license group per AD domain. There must be a default group, which is used for mailboxes which are not covered by separate license groups.

When the license has a license group for the mailbox's domain, this license group is used. If not, the license group defined as default will be used.

There are three situations where Set-OutlookSignatures uses Entra ID/Azure AD via Graph API insteed of on-prem AD: Parameter GraphOnly is set to true, no connection to the on-prem AD is possible, or the current user has a mailbox in Exchange Online and either OOF messages or Outlook Web signatures should be set.
In these cases, license groups are handled as follows:
- If the current mailbox has the Graph "onPremisesDomainName" attribute set:
  - If there is a license group associated with this DNS domain name, it is queried via Graph
  - if there is no license group associated with this DNS domain name, the license group defined as default is queried via Graph
- If the current mailbox does not have the Graph "onPremisesDomainName" attribute set, the license group defined as default is queried via Graph
# Buying, extending and changing licenses
## Buying a new license
Just place a request for quotation with the following information:
- Your billing address
- The VAT number of your company (if applicable)
- email addresses
  - One for receiving invoices
  - One for receiving the download link for the license file, updates and other non invoice related information
- List of license groups and maximum members in the following format:
  - DNS domain name of the Active Directory Domain the group is in.
    - Use 'EntraID' or 'AzureAD' if the group only exists in Azure Active Directory and is not synced to on-prem. Only one pure Entra ID/Azure AD group is supported, it must be the group with the highest priority (first list entry).
  - SID (security identifier) of the group, as string in the "S-[...]" format
  - Maximum number of recursive members in the group (add a buffer for future growth)
  - If multiple license groups are defined, designate one of these groups as default or fallback group. For details, see 'How do license groups work?' later in this document.

The total number of mailboxes to license is the sum of the maximum members defined for each license group.

You will receive an offer within a few days. As soon as all the details are ironed out, you place the final order, receive an invoice and start the payment process. The license file and corresponding Benefactor Circle member Id is sent via email after receipt of payment.
## Extending an existing license
A license period cannot be extended. Licenses are valid for one year, starting with the date the payment is received, and do not auto-renew.
To continue using Set-OutlookSignatures with Benefactor Circle member benefits, just place a new order to receive a new license file.

You will be informed in advance that your license is about to expire.
## Reducing the number of licensed mailboxes

The total number of licensed mailboxes can not be reduced during a license period (one year starting from the date of payment reception), as the license fees are paid in advance.
## Moving licensed mailboxes between license groups
Moving licenses means that the total number of licensed mailboxes does not change, but their distribution across license groups. This can, for example, be necessary due to Active Directory consolidations.

Shifting licenses between license groups is possible once per license period (one year starting from date of payment reception).

If more license shifts are required, additional licenses have to be acquired temporarily, the total number of licenses can then be reduced when the new license period begins.
## Increasing the number of licensed mailboxes
When adding licenses during a license period, you only pay for the new mailboxes and only for the remaining months in the running license period.
The new payment does not extend the existing license period, but it increases the number of licensed mailboxes in it.
An example:
- After a trial with 20 mailboxes, you start a pilot with 110 mailboxes in mid of April 2023. The license is valid until mid of April 2024, with the following cost:
    max(100; 110 * 1.50) = 165.00 € net
- As the pilot is a success, the number of licensed mailboxes is raised to 7,500 in July 2023.
  - The license period does not change, the license is still valid from mid of April 2023 to mid of April 2024, of course with the higher number of mailboxes.
- The added licenses result in the following costs:
  - Year 1 total cost of 9,402.50 €, consisting of
    - Year 1 cost for 110 mailboxes for 12 months: max(100; 110 * 1.50) = 165.00 € net
    - Year 1 additional maiboxes for 10 months (July 2023 to mid of April 2024): max(100; (7,500 - 110) * 1.50)/12*10 = 9,237.50 € net
    - As long as the price is not changing, the consecutive years will cost: max(100; 7,500 * 1,50) = 11,250.00 € net

# License and script version
License and script versions go hand in hand, so every new release of Set-OutlookSignatures also means a new license release, and vice-versa.

Using different versions of script and license file is not supported, as this may lead to unexpected results. When a version mismatch is detected, a warning message is logged.
# Data protection notice
Set-OutlookSignatures and the Benefactor Circle license add-on do not store any telemetry data, do not "phone home", and do not transfer any data, only necessary data between:
- the end user's Windows client,
- the end user's Active Directory or Azure Active Directory,
- and the end user's Exchange or Exchange Online system,
always in the security context of the user executing the program or explicitely assigned to a dedicated enterprise application in Azure Active Directory.

For license purposes, only the absolutely required information is stored and processed: Invoice address, email addresses for technical and commercial communication, license group information (domain, SID, maximum members) and payment information.
