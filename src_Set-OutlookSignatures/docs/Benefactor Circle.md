<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Set-OutlookSignatures%20Benefactor%20Circle%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures Benefactor Circle"></a>**<br>Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages<p><p><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt="this release"></a> XXXRemoveWhenBuildingXXX--> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="latest release" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="open issues" data-external="1"></a> <img src="https://img.shields.io/badge/dynamic/json?color=brightgreen&label=views&query=%24.count&url=https%3A%2F%2Fraw.githubusercontent.com%2FGruberMarkus%2Fmy-traffic2badge%2Ftraffic%2Ftraffic-Set-OutlookSignatures%2Ftraffic_views.json" alt="views" data-external="1"> <img src="https://img.shields.io/badge/dynamic/json?color=brightgreen&label=clones&query=%24.count&url=https%3A%2F%2Fraw.githubusercontent.com%2FGruberMarkus%2Fmy-traffic2badge%2Ftraffic%2Ftraffic-Set-OutlookSignatures%2Ftraffic_clones.json" alt="clones" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures?color=brightgreen" alt="stars" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/donate,%20support,%20sponsor-white?logo=githubsponsors" alt="donate or sponsor"></a> <a href="./Benefactor%20Circle.md" target="_blank"><img src="https://img.shields.io/badge/unlock%20all%20features%20with-Benefactor%20Circle-gold" alt="unlock all features with Benefactor Circle"></a>

# What is the Benefactor Circle?
Benefactor Circle is the result of the partnership between Set-OutlookSignatures and <a href="https://explicitconsulting.at" target="_blank">ExplicIT Consulting</a>.  
<pre><a href="https://explicitconsulting.at" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Others/ExplicIT Consulting, color on black.png" height="100" title="ExplicIT Consulting" alt="ExplicIT Consulting"></a></pre>
ExplicIT Consulting's Benefactor Circle enhances Set-OutlookSignatures with new features and extended support in form of a commercial add-on, ensuring that the core of Set-OutlookSignatures remains Free and Open-Source Software (FOSS) and continues to evolve.

Visit <a href="https://explicitconsulting.at/open-source/set-outlooksignatures" target="_blank">ExplicIT Consulting's Set-OutlookSignatures Benefactor Circle site</a> for details and pricing information.

# Why choose Set-OutlookSignatures?
- Runs only on your clients, no server side installation
- Mails are not routed through a cloud service, no SPF record change
- Software does not call home
- Works with on-prem, hybrid and cloud-only configurations
- Supports Exchange Online roaming signatures
- Multi-customer capable
- Works with linked mailboxes in resource forest scenarios
- Users see signature when writing e-mails
- More cost-effective than other cloud based products, more features than other on-prem products

The features reserved for Benefactor Circle members are available at a very, very competitive price compared to other commercial solutions that work with on-prem, hybrid and cloud-only configurations.

At the end of its commercial life, the Benefactor Circle source code will be handed over to the team developing the core version of Set-OutlookSignatures. This will allow the Benefactor Circle code to be integrated into the Free and Open-Source (FOSS) version of Set-OutlookSignatures.

There are also topics you need to be aware of:
- As there is no server component, signatures can not be automatically added to mails sent from mobile devices. This will change as soon as Microsoft's roaming signature feature will be accessible by an API, and mobile applications start using this feature.  
Set-OutlookSignatures Benefactor Circle already supports the roaming signature feature.
- There is no graphical user interface. This is on purpose:
  - End users typically never see the tool, only results.
  - Admins typically need around two hours for the basic setup, as the default parameters are very well chosen and documented.
  - Template maintainers need nothing but Word to create, modify and configure templates.
# Benefactor Circle benefits
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
# How can I or my company become a member and obtain a licence?
Membership is charged annually in advance.

You receive a Benefactor Circle licence file and a corresponding Id, which you can use to unlock exclusive features in Set-OutlookSignatures.

The licence file contains the following information:
- Invoice address of the Benefactor Circle member
- End date of the membership
- DNS domain name, SID and maximum number of members for one or multiple licence groups

30-day trial licences are available upon request.
## How much is a membership?
The cost of the membership depends on the number of mailbox licences included.

Each mailbox, for which an exlusive feature shall be used, needs a licence. The mailboxes do not need to be named, you just have to define at least one Active Directory or Graph group containing the mailboxes and the maximum number of recursive members of the groups defined.

Licences are paid in advance and are valid for one year from the day the full payment is received. There is no automatic renewal.

At the time of the release of this version of Set-OutlookSignatures, the following prices were set by [ExplicIT Consulting](https://explicitconsulting.at). Please check their homepage for updates.

```
The net price in EUR currently is 1.50 € per mailbox and year, with a minimum annual total sum of 100 €.

Yes, that's right: Per year, not per month.

All release upgrades during the licence period are for free, no matter if it is a patch, feature or major release.

Support may be chargeable. This includes workshops, implementation support, all forms of remote or on site outsourcing, support for topics already well-explained in the documentation and support for problems with the root cause outside of Set-OutlookSignatures or Set-OutlookSignatures Benefactor Circle.
```
# How do licence groups work?
Each Benefactor Circle licence is bound to one or more Active Directory or Entra ID/Azure AD groups. Each mailbox of your company needs to be a direct or indirect (a.k.a. nested, recursive or transitive) member of a licence group, so that it can receive a signature. Primary group membership is not considered due to Active Directory and Entra ID/Azure AD query restrictions.

Each group may only contain as many mailboxes as direct or indirect members as defined in the licence. The user running Set-OutlookSignatures must be able to resolve all direct and indirect members of the licence group, even across trusts.

Licence groups are defined by the DNS domain name of the domain (or 'EntraID' or 'AzureAD' for non-synced groups), their SID (security identifier) and the number of members licensed.
- Use 'EntraID' or 'AzureAD' if the group only exists in Azure Active Directory and is not synced to on-prem. Only one pure Entra ID/Azure AD group is supported, it must be the group with the highest priority (first list entry).
- If you have multiple domains in a forest or multiple forests, you can have multiple licence groups, each with a separate maximum member count. For each licence, there can be one licence group per AD domain. There must be a default group, which is used for mailboxes which are not covered by separate licence groups.

When the licence has a licence group for the mailbox's domain, this licence group is used. If not, the licence group defined as default will be used.

There are three situations where Set-OutlookSignatures uses Entra ID/Azure AD via Graph API insteed of on-prem AD: Parameter GraphOnly is set to true, no connection to the on-prem AD is possible, or the current user has a mailbox in Exchange Online and either OOF messages or Outlook Web signatures should be set.
In these cases, licence groups are handled as follows:
- If the current mailbox has the Graph "onPremisesDomainName" attribute set:
  - If there is a licence group associated with this DNS domain name, it is queried via Graph
  - if there is no licence group associated with this DNS domain name, the licence group defined as default is queried via Graph
- If the current mailbox does not have the Graph "onPremisesDomainName" attribute set, the licence group defined as default is queried via Graph
# Buying, extending and changing licences
## Buying a new licence
Just place a request for quotation with the following information:
- Your billing address
- The VAT number of your company (if applicable)
- E-mail addresses
  - One for receiving invoices
  - One for receiving the download link for the licence file, updates and other non invoice related information
- List of licence groups and maximum members in the following format:
  - DNS domain name of the Active Directory Domain the group is in.
    - Use 'EntraID' or 'AzureAD' if the group only exists in Azure Active Directory and is not synced to on-prem. Only one pure Entra ID/Azure AD group is supported, it must be the group with the highest priority (first list entry).
  - SID (security identifier) of the group, as string in the "S-[...]" format
  - Maximum number of recursive members in the group (add a buffer for future growth)
  - If multiple licence groups are defined, designate one of these groups as default or fallback group. For details, see 'How do licence groups work?' later in this document.

The total number of mailboxes to licence is the sum of the maximum members defined for each licence group.

You will receive an offer within a few days. As soon as all the details are ironed out, you place the final order, receive an invoice and start the payment process. The licence file and corresponding Benefactor Circle member Id is sent via e-mail after receipt of payment.
## Extending an existing licence
A licence period cannot be extended. Licences are valid for one year, starting with the date the payment is received, and do not auto-renew.
To continue using Set-OutlookSignatures with Benefactor Circle member benefits, just place a new order to receive a new licence file.

You will be informed in advance that your licence is about to expire.
## Reducing the number of licenced mailboxes

The total number of licenced mailboxes can not be reduced during a licence period (one year starting from the date of payment reception), as the licence fees are paid in advance.
## Moving licenced mailboxes between licence groups
Moving licences means that the total number of licenced mailboxes does not change, but their distribution across licence groups. This can, for example, be necessary due to Active Directory consolidations.

Shifting licences between licence groups is possible once per licence period (one year starting from date of payment reception).

If more licence shifts are required, additional licences have to be acquired temporarily, the total number of licences can then be reduced when the new licence period begins.
## Increasing the number of licenced mailboxes
When adding licences during a licence period, you only pay for the new mailboxes and only for the remaining months in the running licence period.
The new payment does not extend the existing licence period, but it increases the number of licenced mailboxes in it.
An example:
- After a trial with 20 mailboxes, you start a pilot with 110 mailboxes in mid of April 2023. The licence is valid until mid of April 2024, with the following cost:
    max(100; 110 * 1.50) = 165.00 € net
- As the pilot is a success, the number of licenced mailboxes is raised to 7,500 in July 2023.
  - The licence period does not change, the licence is still valid from mid of April 2023 to mid of April 2024, of course with the higher number of mailboxes.
- The added licences result in the following costs:
  - Year 1 total cost of 9,402.50 €, consisting of
    - Year 1 cost for 110 mailboxes for 12 months: max(100; 110 * 1.50) = 165.00 € net
    - Year 1 additional maiboxes for 10 months (July 2023 to mid of April 2024): max(100; (7,500 - 110) * 1.50)/12*10 = 9,237.50 € net
    - As long as the price is not changing, the consecutive years will cost: max(100; 7,500 * 1,50) = 11,250.00 € net

# Licence and script version
Licence and script versions go hand in hand, so every new release of Set-OutlookSignatures also means a new licence release, and vice-versa.

Using different versions of script and licence file is not supported, as this may lead to unexpected results. When a version mismatch is detected, a warning message is logged.
# Data protection notice
Set-OutlookSignatures and the Benefactor Circle licence add-on do not store any telemetry data, do not "phone home", and do not transfer any data, only necessary data between:
- the end user's Windows client,
- the end user's Active Directory or Azure Active Directory,
- and the end user's Exchange or Exchange Online system,
always in the security context of the user executing the program or explicitely assigned to a dedicated enterprise application in Azure Active Directory.

For licence purposes, only the absolutely required information is stored and processed: Invoice address, e-mail addresses for technical and commercial communication, licence group information (domain, SID, maximum members) and payment information.
