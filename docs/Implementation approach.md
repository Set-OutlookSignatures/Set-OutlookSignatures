<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="../src/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a>**<br>Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages.<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Set-OutlookSignatures" alt=""></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Set-OutlookSignatures/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a>  

# What is the recommended approach for implementing the software? <!-- omit in toc -->
There is certainly no definitive generic recommendation, but this document should be a good starting point.

The content is based on real-life experience implementing the script in a multi-client environment with a five-digit number of mailboxes.

It contains proven procedures and recommendations for product managers, architects, operations managers, account managers and mail and client administrators. It is suited for service providers as well as for clients.

It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The document is available in English and German language.
<br><br>
**Dear businesses using Set-OutlookSignatures:**
- Being Free and Open-Source Software, Set-OutlookSignatures can save you thousands or even tens of thousand Euros/US-Dollars per year in comparison to commercial software.  
Please consider <a href="https://github.com/sponsors/GruberMarkus" target="_blank">sponsoring this project</a> to ensure continued support, testing and enhancements.
- Invest in the open-source projects you depend on. Contributors are working behind the scenes to make open-source better for everyone - give them the help and recognition they deserve.
- Sponsor the open-source software your team has built its business on. Fund the projects that make up your software supply chain to improve its performance, reliability, and stability.
# Table of Contents  <!-- omit in toc -->
- [1. English](#1-english)
  - [1.1. Overview](#11-overview)
  - [1.2. Manual maintenance of signatures](#12-manual-maintenance-of-signatures)
    - [1.2.1. Signatures in Outlook](#121-signatures-in-outlook)
    - [1.2.2. Signature in Outlook on the Web](#122-signature-in-outlook-on-the-web)
  - [1.3. Automatic maintenance of signatures](#13-automatic-maintenance-of-signatures)
    - [1.3.1. Server-based signatures](#131-server-based-signatures)
    - [1.3.2. Client-based signatures](#132-client-based-signatures)
  - [1.4. Synchronization of signatures between different devices](#14-synchronization-of-signatures-between-different-devices)
  - [1.5. Recommendation: Set-OutlookSignatures](#15-recommendation-set-outlooksignatures)
    - [1.5.1. Scope of functions](#151-scope-of-functions)
      - [1.5.1.1. General description, licence model](#1511-general-description-licence-model)
      - [1.5.1.2. Features](#1512-features)
  - [1.6. Administration](#16-administration)
    - [1.6.1. System requirements](#161-system-requirements)
      - [1.6.1.1. Client](#1611-client)
      - [1.6.1.2. Server](#1612-server)
    - [1.6.2. Filing the script](#162-filing-the-script)
    - [1.6.3. Storage of the templates](#163-storage-of-the-templates)
    - [1.6.4. Template management](#164-template-management)
    - [1.6.5. Execute the script](#165-execute-the-script)
      - [1.6.5.1. Parameters](#1651-parameters)
      - [1.6.5.2. Runtime and visibility of the script](#1652-runtime-and-visibility-of-the-script)
      - [1.6.5.3. Using Outlook and Word while the script is running](#1653-using-outlook-and-word-while-the-script-is-running)
    - [1.6.6. Support from the service provider](#166-support-from-the-service-provider)
      - [1.6.6.1. Counselling and introductory phase](#1661-counselling-and-introductory-phase)
        - [1.6.6.1.1. Initial vote on textual signatures](#16611-initial-vote-on-textual-signatures)
          - [1.6.6.1.1.1. Participants](#166111-participants)
          - [1.6.6.1.1.2. Agenda and goals](#166112-agenda-and-goals)
          - [1.6.6.1.1.3. Duration](#166113-duration)
        - [1.6.6.1.2. Training for template administrators](#16612-training-for-template-administrators)
          - [1.6.6.1.2.1. Participants](#166121-participants)
          - [1.6.6.1.2.2. Agenda and goals](#166122-agenda-and-goals)
          - [1.6.6.1.2.3. Duration](#166123-duration)
          - [1.6.6.1.2.4. Prerequisites](#166124-prerequisites)
        - [1.6.6.1.3. Client management training](#16613-client-management-training)
          - [1.6.6.1.3.1. Participants](#166131-participants)
          - [1.6.6.1.3.2. Agenda and goals](#166132-agenda-and-goals)
          - [1.6.6.1.3.3. Duration](#166133-duration)
          - [1.6.6.1.3.4. Prerequisites:](#166134-prerequisites)
      - [1.6.6.2. Tests, pilot operation, rollout](#1662-tests-pilot-operation-rollout)
    - [1.6.7. Running operation](#167-running-operation)
      - [1.6.7.1. Create and maintain templates](#1671-create-and-maintain-templates)
      - [1.6.7.2. Create and maintain storage shares for templates and script components](#1672-create-and-maintain-storage-shares-for-templates-and-script-components)
      - [1.6.7.3. Setting and maintaining AD attributes](#1673-setting-and-maintaining-ad-attributes)
      - [1.6.7.4. Configuration adjustments](#1674-configuration-adjustments)
      - [1.6.7.5. Problems and questions during operation](#1675-problems-and-questions-during-operation)
      - [1.6.7.6. Supported versions](#1676-supported-versions)
      - [1.6.7.7. New versions](#1677-new-versions)
      - [1.6.7.8. Adjustments to the code of the product](#1678-adjustments-to-the-code-of-the-product)
- [2. German (Deutsch)](#2-german-deutsch)
  - [2.1. Überblick](#21-überblick)
  - [2.2. Manuelle Wartung von Signaturen](#22-manuelle-wartung-von-signaturen)
    - [2.2.1. Signaturen in Outlook](#221-signaturen-in-outlook)
    - [2.2.2. Signatur in Outlook im Web](#222-signatur-in-outlook-im-web)
  - [2.3. Automatische Wartung von Signaturen](#23-automatische-wartung-von-signaturen)
    - [2.3.1. Serverbasierte Signaturen](#231-serverbasierte-signaturen)
    - [2.3.2. Clientbasierte Signaturen](#232-clientbasierte-signaturen)
  - [2.4. Abgleich von Signaturen zwischen verschiedenen Geräten](#24-abgleich-von-signaturen-zwischen-verschiedenen-geräten)
  - [2.5. Empfehlung: Set-OutlookSignatures](#25-empfehlung-set-outlooksignatures)
    - [2.5.1. Funktionsumfang](#251-funktionsumfang)
      - [2.5.1.1. Allgemeine Beschreibung, Lizenzmodell](#2511-allgemeine-beschreibung-lizenzmodell)
      - [2.5.1.2. Funktionen](#2512-funktionen)
  - [2.6. Administration](#26-administration)
    - [2.6.1. Systemanforderungen](#261-systemanforderungen)
      - [2.6.1.1. Client](#2611-client)
      - [2.6.1.2. Server](#2612-server)
    - [2.6.2. Ablage des Scripts](#262-ablage-des-scripts)
    - [2.6.3. Ablage der Vorlagen](#263-ablage-der-vorlagen)
    - [2.6.4. Verwaltung der Vorlagen](#264-verwaltung-der-vorlagen)
    - [2.6.5. Ausführen des Scripts](#265-ausführen-des-scripts)
      - [2.6.5.1. Parameter](#2651-parameter)
      - [2.6.5.2. Laufzeit und Sichtbarkeit des Scripts](#2652-laufzeit-und-sichtbarkeit-des-scripts)
      - [2.6.5.3. Nutzung von Outlook und Word während der Laufzeit](#2653-nutzung-von-outlook-und-word-während-der-laufzeit)
    - [2.6.6. Unterstützung durch den Service-Provider](#266-unterstützung-durch-den-service-provider)
      - [2.6.6.1. Beratungs- und Einführungsphase](#2661-beratungs--und-einführungsphase)
        - [2.6.6.1.1. Erstabstimmung zu textuellen Signaturen](#26611-erstabstimmung-zu-textuellen-signaturen)
          - [2.6.6.1.1.1. Teilnehmer](#266111-teilnehmer)
          - [2.6.6.1.1.2. Inhalt und Ziele](#266112-inhalt-und-ziele)
          - [2.6.6.1.1.3. Dauer](#266113-dauer)
        - [2.6.6.1.2. Schulung der Vorlagen-Verwalter](#26612-schulung-der-vorlagen-verwalter)
          - [2.6.6.1.2.1. Teilnehmer](#266121-teilnehmer)
          - [2.6.6.1.2.2. Inhalt und Ziele](#266122-inhalt-und-ziele)
          - [2.6.6.1.2.3. Dauer](#266123-dauer)
          - [2.6.6.1.2.4. Voraussetzungen](#266124-voraussetzungen)
        - [2.6.6.1.3. Schulung des Clientmanagements](#26613-schulung-des-clientmanagements)
          - [2.6.6.1.3.1. Teilnehmer](#266131-teilnehmer)
          - [2.6.6.1.3.2. Inhalt und Ziele](#266132-inhalt-und-ziele)
          - [2.6.6.1.3.3. Dauer](#266133-dauer)
          - [2.6.6.1.3.4. Voraussetzungen](#266134-voraussetzungen)
      - [2.6.6.2. Tests, Pilotbetrieb, Rollout](#2662-tests-pilotbetrieb-rollout)
    - [2.6.7. Laufender Betrieb](#267-laufender-betrieb)
      - [2.6.7.1. Erstellen und Warten von Vorlagen](#2671-erstellen-und-warten-von-vorlagen)
      - [2.6.7.2. Erstellen und Warten von Ablage-Shares für Vorlagen und Script-Komponenten](#2672-erstellen-und-warten-von-ablage-shares-für-vorlagen-und-script-komponenten)
      - [2.6.7.3. Setzen und Warten von AD-Attributen](#2673-setzen-und-warten-von-ad-attributen)
      - [2.6.7.4. Konfigurationsanpassungen](#2674-konfigurationsanpassungen)
      - [2.6.7.5. Probleme und Fragen im laufenden Betrieb](#2675-probleme-und-fragen-im-laufenden-betrieb)
      - [2.6.7.6. Unterstützte Versionen](#2676-unterstützte-versionen)
      - [2.6.7.7. Neue Versionen](#2677-neue-versionen)
      - [2.6.7.8. Anpassungen am Code des Produkts](#2678-anpassungen-am-code-des-produkts)
  
  
# 1. English  
## 1.1. Overview
Textual signatures are not only an essential aspect of corporate identity, but together with the disclaimer are usually a legal necessity.

This document provides a general overview of signatures, guidance for end users, and details of the service provider's recommended solution for centralised management and automated distribution of textual signatures.

The word "signature" in this document is always to be understood as a textual signature and is not to be confused with a digital signature, which serves to encrypt e-mails and/or legitimise the sender.  
## 1.2. Manual maintenance of signatures  
In the case of manual maintenance, the user is provided with a template for the textual signature via the intranet, for example.

Each user sets up the signature himself/herself. Depending on the technical configuration of the client, signatures move with or have to be set up again when the computer used is changed.  
There is no central maintenance.

At the time of writing, most of the service provider's customers are using this variant, while at the same time virtually all existing and also new customers have confirmed their desire for a centrally controlled solution.  
### 1.2.1. Signatures in Outlook
In Outlook, practically any number of signatures can be created per mailbox. This is practical, for example, to distinguish between internal and external e-mails or e-mails in different languages.

In addition, a standard signature for new e-mails and one for replies can be set per mailbox.   
### 1.2.2. Signature in Outlook on the Web
If you also work with Outlook on the Web, you must set up your signature in Outlook on the Web independently of your signature on the client:  
1. Log on to <a href="https://mail.example.com" target="_blank">https<area>://mail.example.com</a> in a web browser. Enter your user name and password, then click Log In.  
2. From the navigation bar, select Settings > Options.  
3. Under Options, select Settings > Email.  
4. Enter the signature you want to use in the text field under Email Signature. Use the Format mini toolbar to change the appearance of the signature.  
5. If you want your signature to appear automatically at the end of all outgoing messages, including replies and forwarded messages, tick Automatically include signature in my sent messages. If you do not activate this option, you can add your signature to each message manually.  
6. Click on Save.  
  
In Outlook on the Web, only one signature is possible.  
## 1.3. Automatic maintenance of signatures  
For some service provider customers, signatures for personal or group mailboxes are automatically created and associated settings configured. For details, please contact your local IT.


The service provider recommends a free script-based solution with central administration and extended client-side functionality, which can be operated and maintained by the customers themselves with the support of the service provider. For details see "Recommendation: Set-OutlookSignatures".  
### 1.3.1. Server-based signatures  
The biggest advantage of a server-based solution is that a defined set of rules is used to capture every email, regardless of which application or device it was sent from.

Since the signature is only attached at the server, the user does not see which signature is used during the creation of an e-mail.

After the signature has been attached to the server, the now modified e-mail must be downloaded again by the client so that it is displayed correctly in the "Sent items" folder. This generates additional network traffic.

If a message is already digitally signed or encrypted when it is created, the textual signature cannot be added on the server side without breaking the digital signature and the encryption. Alternatively, the message is adapted so that the content consists only of the textual signature and the unchanged original message is sent as an attachment.

When evaluating server-based products, the following aspects, among others, should be examined:  
- Can the product handle the number of AD and mail objects in the environment without reproducible crashes or incomplete search results?  
- Does the product have to be installed directly on the mail servers? This means additional dependencies and sources of errors, and can have a negative impact on the availability and reliability of the AD and mail system.  
- Can the administration of the products be delegated directly on the mail servers without granting significant rights? Can customers be authorised separately from each other?  
- Can variables in the signatures only be filled with values from the Active Directory in which Exchange is also located, or also with values from the Active Directory in which the actually authorised users are located when using Linked Mailboxes?  
- How high are the acquisition and maintenance costs? Are these above the tender limit?  
### 1.3.2. Client-based signatures  
In client-based solutions, templates and application rules for textual signatures are defined on a central repository. A component on the client checks the central configuration when it is called up automatically or manually and applies it locally.  
Client-based solutions, unlike server-based solutions, are tied to specific email clients and specific operating systems.


The user already sees the signature during the creation of the e-mail and can adjust it if necessary.


Encryption and digital signing of messages are not a problem on either the client or server side.

When evaluating server-based products, the following aspects, among others, should be examined:  
- Can the product handle the number of AD and mail objects in the environment without reproducible crashes or incomplete search results?  
- Can variables in the signatures only be filled with values from the Active Directory in which Exchange is also located, or also with values from the Active Directory in which the actually authorised users are located when using Linked Mailboxes?  
- Can the product handle group mailboxes and additional connected mailboxes?  
- Can the administration of the products be delegated? Can customers be authorised separately from each other?  
- How high are the acquisition and maintenance costs? Are these above the tender limit?
  
Due to the costs and the fulfilled requirements for functionality and maintainability, the service provider recommends the use of the open-source product Set-OutlookSignatures and offers its customers support for implementation and operation.  
## 1.4. Synchronization of signatures between different devices
The signatures in Outlook, Outlook on the Web and other clients (e.g. in smartphone apps) are not synchronised and must therefore be set up separately.

Depending on the client configuration, Outlook signatures may or may not roam with the user between different Windows devices, please contact your local IT for details.

The client-based tool recommended by the service provider can set signatures in Outlook as well as in Outlook on the Web and also offers the user a simple way to transfer existing signatures to other e-mail clients.   
## 1.5. Recommendation: Set-OutlookSignatures  
The service provider recommends the free open-source software Set-OutlookSignatures after surveying customer requirements and testing several server- and client-based products.

This document provides an overview of the functional scope and administration of the recommended solution, support of the service provider during implementation and operation, as well as associated expenses.  
### 1.5.1. Scope of functions  
#### 1.5.1.1. General description, licence model  
<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank">Set-OutlookSignatures</a> is a free open-source product.  
The product is used for the central administration and local distribution of textual signatures and out-of-office messages to clients. Outlook on Windows is supported as the target platform.

By using the MIT licence, there are virtually no restrictions on the use or modification of the existing code, even for commercial use.

The script is written in PowerShell. PowerShell has been a fixed component of Windows since Windows 7, is actively developed further by Microsoft and is now open-source itself. PowerShell is designed to make it as easy as possible to read and create code.

Since virtually every Microsoft product is administered via PowerShell (even many graphical user interfaces issue PowerShell commands in the background), there is extensive skill and knowledge of this language at the service provider, including in the mail and client product teams.

The integration of PowerShell scripts into the client secured with the help of AppLocker and other mechanisms is also technically and organisationally possible through established measures (such as the digital signing of PowerShell scripts).  
#### 1.5.1.2. Features
Signatures and absence messages can be  
- be generated from templates in DOCX or HTML file format  
- be customised with a wide range of variables, including photos, from Active Directory and other sources  
- be applied to all primary mailboxes (incl. group mailboxes), specific mailbox groups or specific email addresses in all Outlook profiles  
- be assigned time ranges within which they are valid  
- be set as the default signature for new mails or for replies and forwards (signatures only)  
- be defined as an absence message for internal or external recipients (only OOF messages)  
- be set in Outlook Web for the currently logged in user  
- are only centrally managed or exist in parallel with user-created signatures (signatures only)  
- be copied to an alternative path to facilitate access on mobile devices that are not directly supported by this script (signatures only)

Sample templates for signatures and out-of-office messages demonstrate all available functions and are provided as . docx and . htm files.

The product is designed for large and complex environments (Exchange Resource Forest scenarios, across AD trusts, linked mailboxes, multi-level AD subdomains, many objects). The product is multi-client capable.  
## 1.6. Administration
### 1.6.1. System requirements  
#### 1.6.1.1. Client
- Outlook and Word, each from version 2010  
- The script must run in the security context of the user currently logged in.  
- The PowerShell script must be executed in "Full Language Mode". The "Constrained Language Mode" is not supported because certain functions such as Base64 conversions are not available in this mode or require very slow alternatives.  
- If AppLocker or comparable solutions are used, the script may need to be digitally signed.  
- Network releases:  
	- Ports 389 (LDAP) and 3268 (Global Catalog), both TCP and UDP, must be enabled between the client and all domain controllers. If this is not the case, signature-relevant information and variables cannot be retrieved. The script checks with each run whether access is possible.  
	- To access the SMB file share with the script components, the following ports are required: 137 UDP, 138 UDP, 139 TCP, 445 TCP (for details see <a href="https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc731402(v=ws.11)" target="_blank">this Microsoft article</a>).  
	- Port 443 TCP is required to access WebDAV shares (e.g. SharePoint document libraries).  
#### 1.6.1.2. Server
- An SMB file share in which the script and its components are stored. All users must have read access to this file share and its contents.  
- One or more SMB file shares or WEBDAV shares (e.g. SharePoint document libraries) in which the templates for signatures and out-of-office messages are stored and managed.

If variables (e.g. first name, last name, telephone number) are used in the templates, the corresponding values must be available in the Active Directory. In the case of Linked Mailboxes, a distinction can be made between the attributes of the current user and the attributes of the mailbox, which are located in different AD forests.  
### 1.6.2. Filing the script  
As described in the system requirements, the script and its components must be stored on an SMB file share. Alternatively, it can be distributed to the clients by any mechanism and executed from there.

All users need read access to the script and all its components.

As long as these requirements are met, any SMB file share can be used, for example  
- the NETLOGON share of an Active Directory  
- a share on a Windows server in any architecture (single server or cluster, classic share or DFS in all variations)  
- a share on a Windows client  
- a share on any non-Windows system, e.g. via SAMBA

As long as all customers use the same version of the script and only configure it via parameters, a central repository for the script components is sufficient.

For maximum performance and flexibility, it is recommended that each client places the script in its own SMB file share and replicates this across sites to different servers if necessary.  
### 1.6.3. Storage of the templates  
As described in the system requirements, templates for signatures and out-of-office messages can be stored on SMB file shares or WebDAV shares (e.g. SharePoint document libraries) analogous to the script itself.

SharePoint document libraries have the advantage of optional versioning of files, so that in the event of an error, template administrators can quickly restore a previous version of a template.

At least one share with separate subdirectories for signature and absence templates is recommended per client.  
Users need read access to all templates.

By simply granting write access to the entire template folder or to individual files within it, the creation and management of signature and absence templates is delegated to a defined group of people. Typically, the templates are defined, created and maintained by the Corporate Communications and Marketing departments.

For maximum performance and flexibility, it is recommended that each client places the script in its own SMB file share and replicates this across sites to different servers if necessary.  
### 1.6.4. Template management  
By simply granting write access to the template folder or to individual files in it, the creation and management of signature and absence templates is delegated to a defined group of people. Typically, the templates are defined, created and maintained by the Corporate Communications and Marketing departments.

The script can process templates in DOCX or HTML format. For a start, the use of the DOCX format is recommended; the reasons for this recommendation and the advantages and disadvantages of the respective format are described in the script's "readme.html" file.

The file "readme.html" supplied with the script provides an overview of how templates are to be named so that they are  
- Only apply to certain groups or mailboxes  
- be set as the default signature for new mails or replies and forwards  
- be set as an internal or external absence message  
In "readme.html" and the sample templates, the replaceable variables, the extension with user-defined variables and the handling of photos from the Active Directory are also described.

The example file "Test all signature replacement variables.docx" provided contains all variables.  
### 1.6.5. Execute the script  
The script can be executed via any mechanism, for example  
- When the user logs in as part of the logon script or as a script of its own  
- via task scheduling at fixed times or for specific events  
- by the user himself, e.g. via a shortcut on the desktop  
- through a client management tool

Since Set-OutlookSignatures is a pure PowerShell script, it is called like any other script of this file type:
```  
powershell.exe <PowerShell parameter> -file <path to Set-OutlookSignatures.ps1> <Script parameter>  
```  
#### 1.6.5.1. Parameters
The behaviour of the script can be controlled via parameters. Particularly relevant are SignatureTemplatePath and OOFTemplatePath, which are used to specify the path to the signature and absence templates.

The following is an example where the signature templates are on an SMB file share and the out-of-office service provider templates are on a WebDAV share:
```  
powershell.exe -file '\netlogon\set-outlooksignatures\set-outlooksignatures.ps1' -SignatureTemplatePath '\netlogon\set-outlooksignatures\set-outlooksignatures Outlook' -OOFTemplatePath 'https://webdav.example.com/CorporateCommunications/Templates/Out of Office templates'.  
```
At the time of writing, other parameters were available. The following is a brief overview of the possibilities, for details please refer to the documentation of the script in the file "readme.html":  
- SignatureTemplatePath: Path to the signature templates. Can be an SMB or WebDAV share.  
- ReplacementVariableConfigFile: Path to the file in which variables deviating from the standard are defined. Can be an SMB or WebDAV share.  
- TrustsToCheckForGroups: By default, all trusts are queried for mailbox information. This parameter can be used to remove certain domains and add non-trusted domains.  
- DeleteUserCreatedSignatures: Should signatures created by the user be deleted? This is not done by default.  
- SetCurrentUserOutlookWebSignature: By default, a signature is set in Outlook on the web for the logged in user. This can be prevented via this parameter.  
- SetCurrentUserOOFMessage: By default, the text of the out-of-office messages is set. This parameter can be used to change this behaviour.  
- OOFTemplatePath: Path to the absence templates. Can be an SMB or WebDAV share.  
- AdditionalSignaturePath: Path to an additional share to which all signatures are to be copied, e.g. for access from a mobile device and for simplified configuration of clients not supported by the script. Can be an SMB or WebDAV share.  
- UseHtmTemplates: By default, templates are processed in DOCX format. This button can be used to switch to HTML (. htm).  
See '.\docs\README.htm' for more parameters.
#### 1.6.5.2. Runtime and visibility of the script  
The script is designed for fast turnaround time and minimal network load, but the runtime of the script still depends on many parameters:  
- General speed of the client (CPU, RAM, HDD)  
- Number of mailboxes configured in Outlook  
- Number of Trusted Domains  
- Response time of the domain controller and file server  
- Response time of the Exchange servers (setting signatures in Outlook Web, out-of-office notifications)  
- Number of templates and complexity of variables in them (e.g. photos)

A reproducible running time of approx. 30 seconds was measured under the following conditions:  
- Standard client  
- Connected to the company network via VPN  
- 4 mailboxes  
- Query all domains connected by trust  
- 9 signature templates to be processed, all with variables and graphics (but without user photos), partly restricted to groups and mail addresses  
- 8 absence templates to be processed, all with variables and graphics (but without user photos), partly restricted to groups and mail addresses  
- Setting the signature in Outlook on the Web  
- No copying of signatures to an additional network path

Since the script does not require any user interaction, it can be minimised or hidden using the usual mechanisms.  
#### 1.6.5.3. Using Outlook and Word while the script is running
The script does not start Outlook, all queries and configurations are done via the file system and the registry.  
Outlook can be started, used or closed at will while the script is running.

All changes to signatures and out-of-office notifications are immediately visible and usable for the user, with one exception: If the name of the default signature to be used for new e-mails or for replies and forwardings changes, this change only takes effect the next time Outlook is started. If only the content changes, but not the name of one of the standard signatures, this change is immediately available.

Word can be started, used or closed at will while the script is running.

The script uses Word to replace variables in DOCX templates and to convert DOCX and HTML to RTF and TXT. Word is started as a separate invisible process. This process can practically not be influenced by the user and does not influence Word processes started by the user.  
### 1.6.6. Support from the service provider  
The service provider not only recommends the Set-OutlookSignature.ps1 software, but also offers its customers defined free support.

Additional support can be obtained by prior agreement at a separate charge.

The central point of contact for all kinds of questions is Mail Product Management.  
#### 1.6.6.1. Counselling and introductory phase  
The following services are covered by the product price:  
##### 1.6.6.1.1. Initial vote on textual signatures  
###### 1.6.6.1.1.1. Participants  
- Client: Corporate communications, marketing, client management, project coordinator  
- Service provider: mail product management, mail operations management or mail architecture  
###### 1.6.6.1.1.2. Agenda and goals  
- Client: Presentation of own wishes for textual signatures  
- Service provider: Brief description of the basic options for textual signatures, advantages and disadvantages of the different approaches, reasons for deciding on the recommended product.  
- Comparison of customer wishes with the technical and organisational possibilities  
- Live demonstration of the product taking into account the customer's wishes  
- Determining the next steps  
###### 1.6.6.1.1.3. Duration  
4 hours  
##### 1.6.6.1.2. Training for template administrators  
###### 1.6.6.1.2.1. Participants  
- Client: Template administrator (corporate communications, marketing, analyst), optional client management, coordinator of the project.  
- Service provider: mail product management, mail operations management or mail architecture  
###### 1.6.6.1.2.2. Agenda and goals  
- Summary of the previous meeting "Initial agreement on textual signatures", with a focus on desired and realisable functions  
- Presentation of the structure of the template directories, with a focus on  
- Naming conventions  
- Application order (general, group-specific, postbox-specific, alphabetical in each group)  
- Setting default signatures for new emails and for replies and forwards  
- Definition of absence texts for internal and external recipients.  
- Determination of the temporal validity of templates  
- Variables and user photos in templates  
- Differences DOCX and HTML format  
- Options for integrating a disclaimer  
- Joint development of initial templates based on existing templates and customer requirements  
- Live demonstration on a standard client with a test user and test mailboxes of the customer (see requirements)  
###### 1.6.6.1.2.3. Duration  
4 hours  
###### 1.6.6.1.2.4. Prerequisites  
- The client provides a standard client with Outlook and Word.  
- The screen content of the client must be able to be projected by a beamer or displayed on an appropriately large monitor for joint work.  
- The client provides a test user. This test user must be registered on the standard client.  
	- be allowed to download script files from the Internet (github.com) once (alternatively, the client can provide a BitLocker-encrypted USB stick for data transfer).  
	- be allowed to run unsigned PowerShell scripts in Full Language Mode  
	- have a mailbox  
	- Have full access to various test mailboxes (personal mailboxes or group mailboxes) which, if possible, are directly or indirectly members of various groups or distribution lists. For full access, the user can be appropriately authorised to the other mailboxes, or the user name and password of the additional mailboxes are known.  
##### 1.6.6.1.3. Client management training  
###### 1.6.6.1.3.1. Participants  
- Client: Client management, optionally an administrator of the Active Directory, optionally an administrator of the file server and/or SharePoint server, optionally corporate communications and marketing, project coordinator  
- Service provider: mail product management, mail operations management or mail architecture, a representative of the client team at corresponding clients  
###### 1.6.6.1.3.2. Agenda and goals  
- Summary of the previous meeting "Initial agreement on textual signatures", with a focus on desired and realisable functions  
- Presentation of the possibilities with a focus on  
- Basic procedure of the script  
- Client system requirements (Office, PowerShell, AppLocker, digital signature of the script, network ports)  
- System requirements server (storage of templates)  
- Possibilities for integrating the product (logon script, scheduled task, desktop shortcut)  
- Parameterisation of the script, among other things:  
- Announcement of the template folders  
- Consider Outlook on the Web?  
- Take absence messages into account?  
- Which trusts to consider?  
- How to define additional variables?  
- Allow user-created signatures?  
- Store signatures on an additional path?  
- Joint tests based on templates and customer requirements previously developed by the customer  
- Determination of next steps  
###### 1.6.6.1.3.3. Duration  
4 hours  
###### 1.6.6.1.3.4. Prerequisites:  
- The client provides a standard client with Outlook and Word.  
- The screen content of the client must be able to be projected via a beamer or displayed on an appropriately large monitor for joint work.  
- The client provides a test user. This test user must be registered on the standard client.  
	- be allowed to download script files from the Internet (github.com) once (alternatively, the client can provide a BitLocker-encrypted USB stick for data transfer).  
	- be allowed to run unsigned PowerShell scripts in Full Language Mode
	- have a mailbox  
	- Have full access to various test mailboxes (personal mailboxes or group mailboxes) which, if possible, are directly or indirectly members of various groups or distribution lists. For full access, the user can be appropriately authorised to the other mailboxes, or the user name and password of the additional mailboxes are known.  
- The Client shall provide at least one central SMB file or WebDAV share for the storage of the templates.  
- The client provides a central SMB file share for the storage of the script and its components.  
#### 1.6.6.2. Tests, pilot operation, rollout  
The planning and coordination of tests, pilot operation and rollout is carried out by the client's project manager.

The concrete technical implementation is carried out by the client. If, in addition to mail, the client is also serviced by service providers, the client team will assist with the integration of the script (logon script, scheduled task, desktop shortcut).

In the event of fundamental technical problems, Mail Product Management provides support in researching the causes, works out proposals for solutions and, if necessary, establishes contact with the manufacturer of the product.

The creation and maintenance of templates is the responsibility of the client.

For the procedure for adjustments to the code or the release of new functions, see chapter "Ongoing operation".  
### 1.6.7. Running operation
#### 1.6.7.1. Create and maintain templates  
The creation and maintenance of templates is the responsibility of the client.

Mail product management is available to advise on feasibility and impact issues.
#### 1.6.7.2. Create and maintain storage shares for templates and script components
The creation and maintenance of storage shares for templates and script components is the responsibility of the client.  
Mail product management is available to advise on feasibility and impact issues.  
#### 1.6.7.3. Setting and maintaining AD attributes  
Setting and maintaining AD attributes related to textual signatures (e.g. attributes for variables, user photos, group memberships) is the responsibility of the client.

Mail product management is available to advise on feasibility and impact issues.  
#### 1.6.7.4. Configuration adjustments  
Configuration adjustments explicitly provided for by the developers of the script are supported at all times.

Mail product management is available to advise on questions regarding the feasibility and impact of desired adjustments.

The planning and coordination of tests, pilot operation and rollout in connection with configuration adjustments is carried out by the customer, as is the concrete technical implementation.

If, in addition to mail, the client is also managed by the service provider, the client team will assist with the integration of the script (logon script, scheduled task, desktop shortcut).  
#### 1.6.7.5. Problems and questions during operation
In the event of fundamental technical problems, Mail Product Management provides support in researching the causes, works out proposals for solutions and, if necessary, establishes contact with the manufacturer of the product.

For general questions about the product and its application possibilities, the Mail product management is also available.  
#### 1.6.7.6. Supported versions
The version numbers of the product follow the specifications of Semantic Versioning and are therefore structured according to the format "Major.Minor.Patch".  
- "Major" is raised when compatibility with previous versions is no longer given.  
- "Minor" is increased when new functions compatible with previous versions are introduced.  
- "Patch" is raised if the changes include only bug fixes compatible with previous versions.  
- In addition, pre-release identifiers and build metadata are available as attachments to the "Major.Minor.Patch" format, e.g. "-Beta1".

Versions supported by the service provider:  
- The highest version of the product released by the service provider, regardless of its release date.  
- All versions of the product in the highest major branch released by the service provider, provided they are not older than three months.  
- All versions of the product in the second most recent major branch released by the service provider, provided they are not older than three months.
  
Customers therefore have three months after the release of a new version to switch to this version before the service provider support for previously released versions expires.

The release of major branches by the service provider ensures that no more than one changeover in the major area has to take place in the 3-month period. This protects both customers and service providers from gross errors in product development that force incompatibilities and thus new major versions in quick succession.  
#### 1.6.7.7. New versions
When new versions of the product are released, Mail Product Management informs contacts defined by the customer about the changes associated with this version, possible effects on the existing configuration and indicates upgrade options.

The planning and coordination of the introduction of the new version is carried out by the contact person at the customer.

The concrete technical implementation is also carried out by the client. If, in addition to mail, the client is also serviced by service providers, the client team provides support with the integration of the script (logon script, scheduled task, desktop shortcut).

In the event of fundamental technical problems, Mail Product Management provides support in researching the causes, works out proposals for solutions and, if necessary, establishes contact with the manufacturer of the product.  
#### 1.6.7.8. Adjustments to the code of the product
If adaptations to the code of the product are desired, the associated expenses will be estimated and charged separately after commissioning.

In line with the open-source spirit of the product, code adaptations are submitted to the product's developers as suggestions for improvement.

To ensure the maintainability of the product, the service provider can only support code that is officially included in the product. Each customer is free to adapt the code of the product itself, but in this case the service provider can no longer offer support. For details, see "Supported versions".  
# 2. German (Deutsch)  
## 2.1. Überblick  
Textuelle Signaturen sind nicht nur ein wesentlicher Aspekt der Corporate Identity, sondern gemeinsam mit dem Disclaimer im Regelfall eine rechtliche Notwendigkeit.

Dieses Dokument bietet einen generellen Überblick über Signaturen, Anleitungen für Endbenutzer, sowie Details zur vom Service-Provider empfohlenen Lösung zur zentralen Verwaltung und automatisierten Verteilung von textuellen Signaturen.

Das Wort "Signatur" ist in diesem Dokument immer als textuelle Signatur zu verstehen und nicht mit einer digitalen Signatur, die der Verschlüsselung von E-Mails und/oder der Legitimierung des Absenders dient, zu verwechseln.  
## 2.2. Manuelle Wartung von Signaturen  
Bei der manuellen Wartung wird dem Benutzer z. B. über das Intranet eine Vorlage für die textuelle Signatur zur Verfügung gestellt.

Jeder Benutzer richtet sich die Signatur selbst ein. Je nach technischer Konfiguration des Clients wandern Signaturen bei einem Wechsel des verwendeten Computers mit oder sind neu einzurichten.

Eine zentrale Wartung gibt es nicht.

Zum Zeitpunkt der Erstellung dieses Dokuments nutzen die meisten Kunden des Service-Providers diese Variante, während gleichzeitig praktisch alle Bestands- und auch Neukunden ihren Wunsch nach einer zentral gesteuerten Lösung bekräftigt haben.  
### 2.2.1. Signaturen in Outlook  
In Outlook können pro Postfach praktisch beliebig viele Signaturen erstellt werden. Dies ist beispielsweise praktisch, um zwischen internen und externen E-Mails, oder E-Mails in verschiedenen Sprachen zu unterscheiden.

Pro Postfach kann darüber hinaus eine Standard-Signatur für neue E-Mails und eine für Antworten festgelegt werden.   
### 2.2.2. Signatur in Outlook im Web  
Falls Sie auch mit Outlook im Web arbeiten, müssen Sie sich unabhängig von Ihrer Signatur am Client Ihre Signatur in Outlook im Web einrichten:  
1. Melden Sie sich in einem Webbrowser auf <a href="https://mail.example.com" target="_blank">https<area>://mail.example.com</a> an. Geben Sie Ihren Benutzernamen und Ihr Kennwort ein, und klicken Sie dann auf Anmelden.  
2. Wählen Sie auf der Navigationsleiste Einstellungen > Optionen aus.  
3. Wählen Sie unter Optionen den Befehl Einstellungen > E-Mail aus.  
4. Geben Sie im Textfeld unter E-Mail-Signatur die Signatur ein, die Sie verwenden möchten. Verwenden Sie die Minisymbolleiste "Formatieren", um das Aussehen der Signatur zu ändern.  
5. Wenn Ihre Signatur automatisch am Ende aller ausgehenden Nachrichten angezeigt werden soll, und zwar auch in Antworten und weitergeleiteten Nachrichten, aktivieren Sie Signatur automatisch in meine gesendeten Nachrichten einschließen. Wenn Sie diese Option nicht aktivieren, können Sie Ihre Signatur jeder Nachricht manuell hinzufügen.  
6. Klicken Sie auf Speichern.

In Outlook im Web ist nur eine einzige Signatur möglich.  
## 2.3. Automatische Wartung von Signaturen  
Bei einigen Kunden des Service-Providers werden Signaturen für persönliche Postfächer oder Gruppenpostfächer automatisch erstellt und zugehörige Einstellungen konfiguriert. Für Details wenden Sie sich bitte an Ihre lokale IT.

Der Service-Provider empfiehlt eine kostenlose scriptbasierte Lösung mit zentraler Verwaltung und erweitertem clientseitigen Funktionsumfang, die mit Unterstützung des Service-Providers von den Kunden selbst betrieben und gewartet werden kann. Details siehe "Empfehlung: Set-OutlookSignatures".  
### 2.3.1. Serverbasierte Signaturen  
Der größte Vorteil einer serverbasierten Lösung ist, dass an Hand eines definierten Regelsets jedes E-Mail erfasst wird, ganz gleich, von welcher Applikation oder welchem Gerät es verschickt wurde.

Da die Signatur erst am Server angehängt wird, sieht der Benutzer während der Erstellung eines E-Mails nicht, welche Signatur verwendet wird.

Nachdem die Signatur am Server angehängt wurde, muss das nun veränderte E-Mail vom Client neu heruntergeladen werden, damit es im Ordner „Gesendete Elemente“ korrekt angezeigt wird. Das erzeugt zusätzlichen Netzwerkverkehr.

Wird eine Nachricht schon bei Erstellung digital signiert oder verschlüsselt, kann die textuelle Signatur serverseitig nicht hinzugefügt werden, ohne die digitale Signatur und die Verschlüsselung zu brechen. Alternativ wird die Nachricht so angepasst, dass der Inhalt nur aus der textuellen Signatur besteht und unveränderte ursprüngliche Nachricht als Anhang mitgeschickt wird.

Bei der Evaluierung von serverbasierten Produkten sollten unter anderem folgende Aspekte geprüft werden:  
- Kann das Produkt mit der Anzahl der AD- und Mail-Objekte in der Umgebung ohne reproduzierbare Abstürze oder unvollständige Suchergebnissen umgehen?  
- Muss das Produkt direkt auf den Mail-Servern installiert werden? Das bedeutet zusätzliche Abhängigkeiten und Fehlerquellen, und kann sich negativ auf Verfügbarkeit und Zuverlässigkeit des AD- und Mail-Systems auswirken.  
- Kann die Administration der Produkte ohne Vergabe erheblicher Rechte direkt auf den Mail-Servern delegiert werden? Können Kunden separat voneinander berechtigt werden?  
- Können Variablen in den Signaturen nur mit Werten aus dem Active Directory befüllt werden, in dem sich auch Exchange befindet, oder auch mit Werten aus dem Active Directory, in dem sich bei der Verwendung von Linked Mailboxen die tatsächlich berechtigten Benutzer befinden?  
- Wie hoch sind die Anschaffungs- und Wartungskosten? Liegen diese über der Ausschreibungsgrenze?  
### 2.3.2. Clientbasierte Signaturen  
Bei clientbasierten Lösungen werden auf einer zentralen Ablage Vorlagen und Anwendungsregeln für textuelle Signaturen definiert. Eine Komponente am Client prüft bei automatisiertem oder manuellen Aufruf die zentrale Konfiguration und wendet sie lokal an.

Clientbasierte Lösungen sind im Gegensatz zu serverbasierten Lösungen an bestimmte E-Mail-Clients und bestimmte Betriebssysteme gebunden.

Der Benutzer sieht die Signatur bereits während der Erstellung des E-Mails und kann diese gegebenenfalls anpassen.

Die Verschlüsselung und das digitale Signieren von Nachrichten stellen weder client- noch serverseitig ein Problem dar.

Bei der Evaluierung von serverbasierten Produkten sollten unter anderem folgende Aspekte geprüft werden:  
- Kann das Produkt mit der Anzahl der AD- und Mail-Objekte in der Umgebung ohne reproduzierbare Abstürze oder unvollständige Suchergebnissen umgehen?  
- Können Variablen in den Signaturen nur mit Werten aus dem Active Directory befüllt werden, in dem sich auch Exchange befindet, oder auch mit Werten aus dem Active Directory, in dem sich bei der Verwendung von Linked Mailboxen die tatsächlich berechtigten Benutzer befinden?  
- Kann das Produkt mit Gruppenpostfächer und zusätzlich verbundenen Postfächern umgehen?  
- Kann die Administration der Produkte delegiert werden? Können Kunden separat voneinander berechtigt werden?  
- Wie hoch sind die Anschaffungs- und Wartungskosten? Liegen diese über der Ausschreibungsgrenze?
  
Auf Grund der Kosten und der erfüllten Anforderungen an Funktionalität und Wartbarkeit empfiehlt der Service-Provider die Verwendung des Open-Source-Produkts Set-OutlookSignatures und bietet seinen Kunden Unterstützung bei Einführung und Betrieb an.  
## 2.4. Abgleich von Signaturen zwischen verschiedenen Geräten  
Die Signaturen in Outlook, Outlook im Web und anderen Clients (z. B. in Smartphone-Apps) sind nicht synchronisiert und müssen daher separat eingerichtet werden.

Je nach Client-Konfiguration wandern Outlook-Signaturen mit dem Benutzer zwischen verschiedenen Windows-Geräten mit oder nicht, für Details wenden Sie sich bitte an Ihre lokale IT.

Das vom Service-Provider empfohlene clientbasierte Werkzeug kann Signaturen sowohl in Outlook als auch in Outlook im Web setzen und bietet dem Benutzer darüber hinaus eine einfache Möglichkeit zur Übernahme bestehender Signaturen in weitere E-Mail-Clients an.   
## 2.5. Empfehlung: Set-OutlookSignatures  
Der Service-Provider empfiehlt nach einer Erhebung der Kundenanforderungen und Tests mehrerer server- und clientbasierten Produkte die kostenlose Open-Source-Software Set-OutlookSignatures.

Dieses Dokument bietet einen Überblick über Funktionsumfang und Administration der empfohlenen Lösung, Unterstützung des Service-Providers bei Einführung und Betrieb, sowie damit verbundene Aufwände.  
### 2.5.1. Funktionsumfang  
#### 2.5.1.1. Allgemeine Beschreibung, Lizenzmodell  
<a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank">Set-OutlookSignatures</a> ist ein kostenloses Open-Source-Produkt.

Das Produkt dient der zentralen Verwaltung und lokalen Verteilung textueller Signaturen und Abwesenheits-Nachrichten auf Clients. Als Zielplattform wird dabei Outlook auf Windows unterstützt.

Durch die Verwendung der MIT-Lizenz gibt es auch bei kommerzieller Nutzung praktisch keine Einschränkungen in Bezug auf Nutzung oder Veränderung des bestehenden Codes.

Das Script ist in PowerShell geschrieben. PowerShell ist seit Windows 7 fixer Bestandteil von Windows, wird von Microsoft aktiv weiterentwickelt und ist mittlerweile selbst Open-Source-Software. PowerShell ist darauf ausgelegt, den Code möglichst einfach lesen und erstellen zu können.

Da praktisch jedes Microsoft-Produkt über PowerShell administriert wird (selbst viele grafische Oberflächen setzen im Hintergrund PowerShell-Befehle ab), gibt es beim Service-Provider umfangreiches Können und Wissen zu dieser Sprache, auch in den Mail- und Client-Produktteams.

Auch die Einbindung von PowerShell-Skripten in den mit Hilfe von AppLocker und anderen Mechanismen abgesicherten Client ist durch etablierte Maßnahmen (wie z. B. dem digitalen Signieren von PowerShell-Skripten) technisch und organisatorisch möglich.  
#### 2.5.1.2. Funktionen  
Signaturen und Abwesenheits-Nachrichten können  
- aus Vorlagen im DOCX- oder HTML-Dateiformat erzeugt werden  
- mit einer breiten Palette von Variablen, einschließlich Fotos, aus Active Directory und anderen Quellen angepasst werden  
- auf alle primären Postfächer (inkl. Gruppenpostfächer), bestimmte Postfach-Gruppen oder bestimmte E-Mail-Adressen in allen Outlook-Profilen angewandt werden  
- Zeitbereichen zugewiesen werden, innerhalb derer sie gültig sind  
- als Standardsignatur für neue Mails oder für Antworten und Weiterleitungen (nur Signaturen) festgelegt werden  
- als Abwesenheits-Nachricht für interne oder externe Empfänger definiert werden (nur OOF-Nachrichten)  
- in Outlook Web für den aktuell angemeldeten Benutzer gesetzt werden  
- nur zentral verwaltet werden oder parallel mit vom Benutzer erstellten Signaturen existieren (nur Signaturen)  
- in einen alternativen Pfad kopiert werden, um den Zugriff auf mobilen Geräten zu erleichtern, die nicht direkt von diesem Skript unterstützt werden (nur Signaturen)
  
Beispielvorlagen für Signaturen und Abwesenheits-Nachrichten demonstrieren alle verfügbaren Funktionen und werden als .docx- und .htm-Dateien bereitgestellt.

Das Produkt ist auf große und komplexe Umgebungen ausgelegt (Exchange Resource Forest-Szenarien, über AD-Trusts hinweg, Linked Mailboxen, mehrstufige AD-Subdomänen, viele Objekte). Das Produkt ist mandantenfähig.  
## 2.6. Administration  
### 2.6.1. Systemanforderungen  
#### 2.6.1.1. Client  
- Outlook und Word, jeweils ab Version 2010  
- Das Script muss im Sicherheitskontext des aktuell angemeldeten Benutzers laufen.  
- Das PowerShell-Script muss im „Full Language Mode” ausgeführt werden. Der „Constrained Language Mode“ wird nicht unterstützt, da gewisse Funktionen wie z. B. Base64-Konvertierungen in diesem Modus nicht verfügbar sind oder sehr langsame Alternativen benötigen.  
- Falls AppLocker oder vergleichbare Lösungen zum Einsatz kommen, muss das Script möglicherweise digital signiert werden.  
- Netzwerkfreischaltungen:  
	- Die Ports 389 (LDAP) and 3268 (Global Catalog), jeweils TCP and UDP, müssen zwischen Client und allen Domain Controllern freigeschaltet sein. Falls dies nicht der Fall ist, können signaturrelevante Informationen und Variablen nicht abgerufen werden. Das Script prüft bei jedem Lauf, ob der Zugriff möglich ist.  
	- Für den Zugriff auf den SMB-File-Share mit den Script-Komponenten werden folgende Ports benötigt: 137 UDP, 138 UDP, 139 TCP, 445 TCP (Details <a href="https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc731402(v=ws.11)" target="_blank">in diesem Microsoft-Artikel</a>).  
	- Für den Zugriff auf WebDAV-Shares (z. B. SharePoint Dokumentbibliotheken) wird Port 443 TCP benötigt.  
#### 2.6.1.2. Server  
- Ein SMB-File-Share, in den das Script und seine Komponenten abgelegt werden. Auf diesen File-Share und seine Inhalte müssen alle Benutzer lesend zugreifen können.  
- Ein oder mehrere SMB-File-Shares oder WEBDAV-Shares (z. B. SharePoint Dokumentbibliotheken), in den die Vorlagen für Signaturen und Abwesenheitsnachrichten gespeichert und verwaltet werden.

Falls in den Vorlagen Variablen (z. B. Vorname, Nachname, Telefonnummer) genutzt werden, müssen die entsprechenden Werte im Active Directory vorhanden sein. Im Fall von Linked Mailboxes kann dabei zwischen den Attributen des aktuellen Benutzers und den Attributen des Postfachs, die sich in unterschiedlichen AD-Forests befinden, unterschieden werden.  
### 2.6.2. Ablage des Scripts  
Wie in den Systemanforderungen beschrieben, ist das Script samt seinen Komponenten auf einem SMB-File-Share abzulegen. Alternativ kann es durch einen beliebigen Mechanismus auf die Clients verteilt und von dort ausgeführt werden.

Alle Benutzer benötigen Lesezugriff auf das Script und alle seine Komponenten.

Solange diese Anforderungen erfüllt sind, kann jeder beliebige SMB-File-Share genutzt werden, beispielsweise  
- der NETLOGON-Share eines Active Directory  
- ein Share auf einem Windows-Server in beliebiger Architektur (einzelner Server oder Cluster, klassischer Share oder DFS in allen Variationen)  
- ein Share auf einem Windows-Client  
- ein Share auf einem beliebigen Nicht-Windows-System, z. B. über SAMBA

Solange alle Kunden die gleiche Version des Scripts einsetzen und dieses nur über Parameter konfigurieren, genügt eine zentrale Ablage für die Script-Komponenten.

Für maximale Leistung und Flexibilität wird empfohlen, dass jeder Kunde das Script in einem eigenen SMB-File-Share ablegt und diesen gegebenenfalls über Standorte hinweg auf verschiedene Server repliziert.  
### 2.6.3. Ablage der Vorlagen  
Wie in den Systemanforderungen beschrieben, können Vorlagen für Signaturen und Abwesenheitsnachrichten analog zum Script selbst auf SMB-File-Shares oder WebDAV-Shares (z. B. SharePoint Dokumentbibliotheken) abgelegt werden.

SharePoint-Dokumentbibliotheken haben den Vorteil der optionalen Versionierung von Dateien, so dass im Fehlerfall durch die Vorlagen-Verwalter rasch eine frühere Version einer Vorlage wiederhergestellt werden kann.

Es wird pro Kunde zumindest ein Share mit separaten Unterverzeichnissen für Signatur- und Abwesenheits-Vorlagen empfohlen.

Benutzer benötigen lesenden Zugriff auf alle Vorlagen.

Durch simple Vergabe von Schreibrechten auf den gesamten Vorlagen-Ordner oder auf einzelne Dateien darin wird die Erstellung und Verwaltung von Signatur- und Abwesenheits-Vorlagen an eine definierte Gruppe von Personen delegiert. Üblicherweise werden die Vorlagen von den Abteilungen Unternehmenskommunikation und Marketing definiert, erstellt und gewartet.

Für maximale Leistung und Flexibilität wird empfohlen, dass jeder Kunde das Script in einem eigenen SMB-File-Share ablegt und diesen gegebenenfalls über Standorte hinweg auf verschiedene Server repliziert.  
### 2.6.4. Verwaltung der Vorlagen  
Durch simple Vergabe von Schreibrechten auf den Vorlagen-Ordner oder auf einzelne Dateien darin wird die Erstellung und Verwaltung von Signatur- und Abwesenheits-Vorlagen an eine definierte Gruppe von Personen delegiert. Üblicherweise werden die Vorlagen von den Abteilungen Unternehmenskommunikation und Marketing definiert, erstellt und gewartet.

Das Script kann Vorlagen im DOCX- oder im HTML-Format verarbeiten. Für den Anfang wird die Verwendung des DOCX-Formats empfohlen; die Gründe für diese Empfehlung und die Vor- und Nachteile des jeweiligen Formats werden in der Datei „readme.html“ des Scripts beschrieben.

Die mit dem Script mitgelieferte Datei „readme.html“ bietet eine Übersicht, wie Vorlagen zu benennen sind, damit sie  
- nur für bestimmte Gruppen oder Postfächer gelten  
- als Standard-Signatur für neue Mails oder Antworten und Weiterleitungen gesetzt werden  
- als interne oder externe Abwesenheits-Nachricht gesetzt werden

In „readme.html“ und den Beispiel-Vorlagen werden zudem die ersetzbaren Variablen, die Erweiterung um benutzerdefinierte Variablen und der Umgang mit Fotos aus dem Active Directory beschrieben.

In der mitgelieferten Beispiel-Datei „Test all signature replacement variables.docx“ sind alle Variablen enthalten.  
### 2.6.5. Ausführen des Scripts  
Das Script kann über einen beliebigen Mechanismus ausgeführt werden, beispielsweise  
- bei Anmeldung des Benutzers als Teil des Logon-Scripts oder als eigenes Script  
- über die Aufgabenplanung zu fixen Zeiten oder bei bestimmten Ereignissen  
- durch den Benutzer selbst, z. B. über eine Verknüpfung auf dem Desktop  
- durch ein Werkzeug zur Client-Verwaltung

Da es sich bei Set-OutlookSignatures um ein reines PowerShell-Script handelt, erfolgt der Aufruf wie bei jedem anderen Script dieses Dateityps:  
```  
powershell.exe <PowerShell-Parameter> -file <Pfad zu Set-OutlookSignatures.ps1> <Script-Parameter>  
```  
#### 2.6.5.1. Parameter  
Das Verhalten des Scripts kann über Parameter gesteuert werden. Besonders relevant sind dabei SignatureTemplatePath und OOFTemplatePath, über die der Pfad zu den Signatur- und Abwesenheits-Vorlagen angegeben wird.

Folgend ein Beispiel, bei dem die Signatur-Vorlagen auf einem SMB-File-Share und die AbwesenheService-Providerorlagen auf einem WebDAV-Share liegen:  
```  
powershell.exe -file '\\example.com\netlogon\set-outlooksignatures\set-outlooksignatures.ps1' –SignatureTemplatePath '\\example.com\DFS-Share\Common\Templates\Signatures Outlook' –OOFTemplatePath 'https://webdav.example.com/CorporateCommunications/Templates/Out of Office templates'  
```

Zum Zeitpunkt der Erstellung dieses Dokuments waren noch weitere Parameter verfügbar. Folgend eine kurze Übersicht der Möglichkeit, für Details sei auf die Dokumentation des Scripts in der Datei „readme.html“ verwiesen:  
- SignatureTemplatePath: Pfad zu den Signatur-Vorlagen. Kann ein SMB- oder WebDAV-Share sein.  
- ReplacementVariableConfigFile: Pfad zur Datei, in der vom Standard abweichende Variablen definiert werden. Kann ein SMB- oder WebDAV-Share sein.  
- TrustsToCheckForGroups: Standardmäßig werden alle Trusts nach Postfachinformationen abgefragt. Über diesen Parameter können bestimmte Domains entfernt und nicht-getrustete Domains hinzugefügt werden.  
- DeleteUserCreatedSignatures: Sollen vom Benutzer selbst erstelle Signaturen gelöscht werden? Standardmäßig erfolgt dies nicht.  
- SetCurrentUserOutlookWebSignature: Standardmäßig wird für den angemeldeten Benutzer eine Signatur in Outlook im Web gesetzt. Über diesen Parameter kann das verhindert werden.  
- SetCurrentUserOOFMessage: Standardmäßig wird der Text der Abwesenheits-Nachrichten gesetzt. Über diesen Parameter kann dieses Verhalten geändert werden.  
- OOFTemplatePath: Pfad zu den Abwesenheits-Vorlagen. Kann ein SMB- oder WebDAV-Share sein.  
- AdditionalSignaturePath: Pfad zu einem zusätzlichen Share, in den alle Signaturen kopiert werden sollen, z. B. für den Zugriff von einem mobilen Gerät aus und zur vereinfachten Konfiguration nicht vom Script unterstützter Clients. Kann ein SMB- oder WebDAV-Share sein.  
- UseHtmTemplates: Standardmäßig werden Vorlagen im DOCX-Format verarbeitet. Über diesen Schalter kann auf HTML (.htm) umgeschaltet werden.  
Die Datei '.\docs\README.htm' enthält weitere Parameter.
#### 2.6.5.2. Laufzeit und Sichtbarkeit des Scripts  
Das Script ist auf schnelle Durchlaufzeit und minimale Netzwerkbelastung ausgelegt, die Laufzeit des Scripts hängt dennoch von vielen Parametern ab:  
- allgemeine Geschwindigkeit des Clients (CPU, RAM, HDD)  
- Anzahl der in Outlook konfigurierten Postfächer  
- Anzahl der Trusted Domains  
- Reaktionszeit der Domain Controller und File Server  
- Reaktionszeit der Exchange-Server (Setzen von Signaturen in Outlook Web, Abwesenheits-Benachrichtigungen)  
- Anzahl der Vorlagen und Komplexität der Variablen darin (z. B. Fotos)

Unter folgenden Rahmenbedingungen wurde eine reproduzierbare Laufzeit von ca. 30 Sekunden gemessen:  
- Standard-Client  
- Über VPN mit dem Firmennetzwerk verbunden  
- 4 Postfächer  
- Abfrage aller per Trust verbundenen Domains  
- 9 zu verarbeitende Signatur-Vorlagen, alle mit Variablen und Grafiken (aber ohne Benutzerfotos), teilweise auf Gruppen und Mail-Adressen eingeschränkt  
- 8 zu verarbeitende Abwesenheits-Vorlagen, alle mit Variablen und Grafiken (aber ohne Benutzerfotos), teilweise auf Gruppen und Mail-Adressen eingeschränkt  
- Setzen der Signatur in Outlook im Web  
- Kein Kopieren der Signaturen auf einen zusätzlichen Netzwerkpfad
  
Da das Script keine Benutzerinteraktion erfordert, kann es über die üblichen Mechanismen minimiert oder versteckt ausgeführt werden.  
#### 2.6.5.3. Nutzung von Outlook und Word während der Laufzeit  
Das Script startet Outlook nicht, alle Abfragen und Konfigurationen erfolgen über das Dateisystem und die Registry.

Outlook kann während der Ausführung des Scripts nach Belieben gestartet, verwendet oder geschlossen werden.

Sämtliche Änderungen an Signaturen und Abwesenheits-Benachrichtigungen sind für den Benutzer sofort sichtbar und verwendbar, mit einer Ausnahme: Falls sich der Name der zu verwendenden Standard-Signatur für neue E-Mails oder für Antworten und Weiterleitungen ändert, so greift diese Änderung erst beim nächsten Start von Outlook. Ändert sich nur der Inhalt, aber nicht der Name einer der Standard-Signaturen, so ist diese Änderung sofort verfügbar.

Word kann während der Ausführung des Scripts nach Belieben gestartet, verwendet oder geschlossen werden.

Das Script nutzt Word zum Ersatz von Variablen in DOCX-Vorlagen und zum Konvertieren von DOCX und HTML nach RTF und TXT. Word wird dabei als eigener unsichtbarer Prozess gestartet. Dieser Prozess kann vom Benutzer praktisch nicht beeinflusst werden und beeinflusst vom Benutzer gestartete Word-Prozesse nicht.  
### 2.6.6. Unterstützung durch den Service-Provider  
Der Service-Provider empfiehlt die Software Set-OutlookSignature.ps1 nicht nur, sondern bietet seinen Kunden auch definierte kostenlose Unterstützung an.

Darüberhinausgehende Unterstützung kann nach vorheriger Abstimmung gegen separate Verrechnung bezogen werden.

Zentrale Anlaufstelle für Fragen aller Art ist das Mail-Produktmanagement.  
#### 2.6.6.1. Beratungs- und Einführungsphase  
Folgende Leistungen sind mit dem Produktpreis abgedeckt:  
##### 2.6.6.1.1. Erstabstimmung zu textuellen Signaturen  
###### 2.6.6.1.1.1. Teilnehmer  
- Kunde: Unternehmenskommunikation, Marketing, Clientmanagement, Koordinator des Vorhabens  
- Service-Provider: Mail-Produktmanagement, Mail-Betriebsführung oder Mail-Architektur  
###### 2.6.6.1.1.2. Inhalt und Ziele  
- Kunde: Vorstellung der eigenen Wünsche zu textuellen Signaturen  
- Service-Provider: Kurze Beschreibung zu prinzipiellen Möglichkeiten rund um textuelle Signaturen, Vor- und Nachteile der unterschiedlichen Ansätze, Gründe für die Entscheidung zum empfohlenen Produkt  
- Abgleich der Kundenwünsche mit den technisch-organisatorischen Möglichkeiten  
- Live-Demonstration des Produkts unter Berücksichtigung der Kundenwünsche  
- Festlegung der nächsten Schritte  
###### 2.6.6.1.1.3. Dauer  
4 Stunden  
##### 2.6.6.1.2. Schulung der Vorlagen-Verwalter  
###### 2.6.6.1.2.1. Teilnehmer  
- Kunde: Vorlagen-Verwalter (Unternehmenskommunikation, Marketing, Analytiker), optional Clientmanagement, Koordinator des Vorhabens  
- Service-Provider: Mail-Produktmanagement, Mail-Betriebsführung oder Mail-Architektur  
###### 2.6.6.1.2.2. Inhalt und Ziele  
- Zusammenfassung des vorangegangenen Termins „Erstabstimmung zu textuellen Signaturen“, mit Fokus auf gewünschte und realisierbare Funktionen  
- Vorstellung des Aufbaus der Vorlagen-Verzeichnisse, mit Fokus auf  
- Namenskonventionen  
- Anwendungsreihenfolge (allgemein, gruppenspezifisch, postfachspezifisch, in jeder Gruppe alphabetisch)  
- Festlegung von Standard-Signaturen für neue E-Mails und für Antworten und Weiterleitungen  
- Festlegung von Abwesenheits-Texten für interne und externe Empfänger.  
- Festlegung der zeitlichen Gültigkeit von Vorlagen  
- Variablen und Benutzerfotos in Vorlagen  
- Unterschiede DOCX- und HTML-Format  
- Möglichkeiten zur Einbindung eines Disclaimers  
- Gemeinsame Erarbeitung erster Vorlagen auf Basis bestehender Vorlagen und Kundenanforderungen  
- Live-Demonstration auf einem Standard-Client mit einem Testbenutzer und Testpostfächern des Kunden (siehe Voraussetzungen)  
###### 2.6.6.1.2.3. Dauer  
4 Stunden  
###### 2.6.6.1.2.4. Voraussetzungen  
- Der Kunde stellt einen Standard-Client mit Outlook und Word zu fVerfügung.  
- Der Bildschirminhalt des Clients muss zur gemeinsamen Arbeit per Beamer projiziert oder auf einem entsprechend großen Monitor dargestellt werden können.  
- Der Kunde stellt einen Testbenutzer zur Verfügung. Dieser Testbenutzer muss auf dem Standard-Client  
	- einmalig Script-Dateien aus dem Internet (github.com) herunterladen dürfen (alternativ kann der Kunde einen BitLocker-verschlüsselten USB-Stick für die Datenübertragung stellen).  
	- unsignierte PowerShell-Scripte im Full Language Mode ausführen dürfen  
	- über ein Mail-Postfach verfügen  
	- Vollzugriff auf diverse Testpostfächer (persönliche Postfächer oder Gruppenpostfächer) haben, die nach Möglichkeit direkt oder indirekt Mitglied in diversen Gruppen oder Verteilerlisten sind. Für den Vollzugriff kann der Benutzer auf die anderen Postfächer entsprechend berechtigt sein, oder Benutzername und Passwort der zusätzlichen Postfächer sind bekannt.  
##### 2.6.6.1.3. Schulung des Clientmanagements  
###### 2.6.6.1.3.1. Teilnehmer  
- Kunde: Clientmanagement, optional ein Administrator des Active Directory, optional ein Administrator des File-Servers und/oder SharePoint-Server, optional Unternehmenskommunikation und Marketing, Koordinator des Vorhabens  
- Service-Provider: Mail-Produktmanagement, Mail-Betriebsführung oder Mail-Architektur, ein Vertreter des Client-Teams bei entsprechenden Kunden  
###### 2.6.6.1.3.2. Inhalt und Ziele  
- Zusammenfassung des vorangegangenen Termins „Erstabstimmung zu textuellen Signaturen“, mit Fokus auf gewünschte und realisierbare Funktionen  
- Vorstellung der Möglichkeiten mit Fokus auf  
- Prinzipieller Ablauf des Scripts  
- Systemanforderungen Client (Office, PowerShell, AppLocker, digitale Signatur des Scripts, Netzwerk-Ports)  
- Systemanforderungen Server (Ablage der Vorlagen)  
- Möglichkeiten der Einbindung des Produkts (Logon-Script, geplante Aufgabe, Desktop-Verknüpfung)  
- Parametrisierung des Scripts, unter anderem:  
- Bekanntgabe der Vorlagen-Ordner  
- Outlook im Web berücksichtigen?  
- Abwesenheitsnachrichten berücksichtigen?  
- Welche Trusts berücksichtigen?  
- Wie zusätzliche Variablen definieren?  
- Vom Benutzer erstellte Signaturen erlauben?  
- Signaturen auf einem zusätzlichen Pfad ablegen?  
- Gemeinsame Tests auf Basis zuvor vom Kunden erarbeiteter Vorlagen und Kundenanforderungen  
- Festlegung nächster Schritte  
###### 2.6.6.1.3.3. Dauer  
4 Stunden  
###### 2.6.6.1.3.4. Voraussetzungen  
- Der Kunde stellt einen Standard-Client mit Outlook und Word zu Verfügung.  
- Der Bildschirminhalt des Clients muss zur gemeinsamen Arbeit per Beamer projiziert oder auf einem entsprechend großen Monitor dargestellt werden können.  
- Der Kunde stellt einen Testbenutzer zur Verfügung. Dieser Testbenutzer muss auf dem Standard-Client  
	- einmalig Script-Dateien aus dem Internet (github.com) herunterladen dürfen (alternativ kann der Kunde einen BitLocker-verschlüsselten USB-Stick für die Datenübertragung stellen).  
	- unsignierte PowerShell-Scripte im Full Language Mode ausführen dürfen
	- über ein Mail-Postfach verfügen  
	- Vollzugriff auf diverse Testpostfächer (persönliche Postfächer oder Gruppenpostfächer) haben, die nach Möglichkeit direkt oder indirekt Mitglied in diversen Gruppen oder Verteilerlisten sind. Für den Vollzugriff kann der Benutzer auf die anderen Postfächer entsprechend berechtigt sein, oder Benutzername und Passwort der zusätzlichen Postfächer sind bekannt.  
- Der Kunde stellt mindestens einen zentralen SMB-File- oder WebDAV-Share für die Ablage der Vorlagen zur Verfügung.  
- Der Kunde stellt einen zentralen SMB-File-Share für die Ablage des Scripts und seiner Komponenten zur Verfügung.  
#### 2.6.6.2. Tests, Pilotbetrieb, Rollout  
Die Planung und Koordination von Tests, Pilotbetrieb und Rollout erfolgt durch den Vorhabens-Verantwortlichen des Kunden.

Die konkrete technische Umsetzung erfolgt durch den Kunden. Falls zusätzlich zu Mail auch der Client durch Service-Provider betreut wird, unterstützt das Client-Team bei der Einbindung des Scripts (Logon-Script, geplante Aufgabe, Desktop-Verknüpfung).

Bei prinzipiellen technischen Problemen unterstützt das Mail-Produktmanagement bei der Ursachenforschung, arbeitet Lösungsvorschläge aus und stellt gegebenenfalls den Kontakt zum Hersteller des Produkts her.

Die Erstellung und Wartung von Vorlagen ist Aufgabe des Kunden

Zur Vorgehensweise bei Anpassungen am Code oder der Veröffentlichung neuer Funktionen siehe Kapitel „Laufender Betrieb“.  
### 2.6.7. Laufender Betrieb  
#### 2.6.7.1. Erstellen und Warten von Vorlagen  
Das Erstellen und Warten von Vorlagen ist Aufgabe des Kunden.  
Das Mail-Produktmanagement steht für Fragen zu Realisierbarkeit und Auswirkungen beratend zur Verfügung.

#### 2.6.7.2. Erstellen und Warten von Ablage-Shares für Vorlagen und Script-Komponenten  
Das Erstellen und Warten von Ablage-Shares für Vorlagen und Script-Komponenten ist Aufgabe des Kunden.

Das Mail-Produktmanagement steht für Fragen zu Realisierbarkeit und Auswirkungen beratend zur Verfügung.  
#### 2.6.7.3. Setzen und Warten von AD-Attributen  
Das Setzen und Warten von AD-Attributen, die im Zusammenhang mit textuellen Signaturen stehen (z. B. Attribute für Variablen, Benutzerfotos, Gruppenmitgliedschaften), ist Aufgabe des Kunden.

Das Mail-Produktmanagement steht für Fragen zu Realisierbarkeit und Auswirkungen beratend zur Verfügung.  
#### 2.6.7.4. Konfigurationsanpassungen  
Konfigurationsanpassungen, die von den Entwicklern des Scripts explizit vorgesehen sind, werden jederzeit unterstützt.

Das Mail-Produktmanagement steht für Fragen zur Realisierbarkeit und den Auswirkungen gewünschter Anpassungen beratend zur Verfügung.

Die Planung und Koordination von Tests, Pilotbetrieb und Rollout im Zusammenhang mit Konfigurationsanpassungen erfolgt ebenso durch den Kunden wie die konkrete technische Umsetzung.

Falls zusätzlich zu Mail auch der Client durch den Service-Provider betreut wird, unterstützt das Client-Team bei der Einbindung des Scripts (Logon-Script, geplante Aufgabe, Desktop-Verknüpfung).  
#### 2.6.7.5. Probleme und Fragen im laufenden Betrieb  
Bei prinzipiellen technischen Problemen unterstützt das Mail-Produktmanagement bei der Ursachenforschung, arbeitet Lösungsvorschläge aus und stellt gegebenenfalls den Kontakt zum Hersteller des Produkts her.

Für allgemeine Fragen zum Produkt und dessen Anwendungsmöglichkeiten steht ebenfalls das Mail-Produktmanagement zur Verfügung.  
#### 2.6.7.6. Unterstützte Versionen  
Die Versionsnummern des Produkts folgen den Vorgaben des Semantic Versioning und sind daher nach dem Format „Major.Minor.Patch“ aufgebaut.  
- „Major“ wird erhöht, wenn die Kompatibilität zu bisherigen Versionen nicht mehr gegeben ist.  
- „Minor“ wird erhöht, wenn neue Funktionen, die zu bisherigen Versionen kompatibel sind, eingeführt werden.  
- „Patch“ wird erhöht, wenn die Änderungen ausschließlich zu bisherigen Versionen kompatible Fehlerbehebungen umfassen.  
- Zusätzlich sind Bezeichner für Vorveröffentlichungen und Build-Metadaten als Anhänge zum „Major.Minor.Patch“-Format verfügbar, z. B. „-Beta1“.

Vom Service-Provider unterstützte Versionen:  
- Die höchste vom Service-Provider freigegebene Version des Produkts, unabhängig von deren Veröffentlichungsdatum.  
- Alle Versionen des Produkts im höchsten von der Service-Provider freigegebenen Major-Zweig, sofern diese nicht älter als drei Monate sind.  
- Alle Versionen des Produkts im zweitaktuellsten von der Service-Provider freigegebenen Major-Zweig, sofern diese nicht älter als drei Monate sind.

Kunden haben nach Freigabe einer neuen Version also drei Monate Zeit, auf diese Version umzusteigen, bevor der Service-Provider-Support für davor freigegebene Versionen erlischt.

Die Freigabe von Major-Zweigen durch den Service-Provider stellt sicher, dass im 3-Monats-Zeitraum nicht mehr als eine Umstellung im Major-Bereich erfolgen muss. Dies schützt sowohl Kunden als auch Service-Provider vor groben Fehlern in der Produktentwicklung, die in rascher Folge Inkompatibilitäten und damit neue Major-Versionen erzwingen.  
#### 2.6.7.7. Neue Versionen  
Wenn neue Versionen des Produkts veröffentlicht werden, informiert das Mail-Produktmanagement vom Kunden definierte Ansprechpartner über die mit dieser Version verbundenen Änderungen, mögliche Auswirkungen auf die bestehende Konfiguration und zeigt Aktualisierungsmöglichkeiten auf.

Die Planung und Koordination der Einführung der neuen Version erfolgt durch den Ansprechpartner beim Kunden.

Die konkrete technische Umsetzung erfolgt ebenfalls durch den Kunden. Falls zusätzlich zu Mail auch der Client durch Service-Provider betreut wird, unterstützt das Client-Team bei der Einbindung des Scripts (Logon-Script, geplante Aufgabe, Desktop-Verknüpfung).

Bei prinzipiellen technischen Problemen unterstützt das Mail-Produktmanagement bei der Ursachenforschung, arbeitet Lösungsvorschläge aus und stellt gegebenenfalls den Kontakt zum Hersteller des Produkts her.  
#### 2.6.7.8. Anpassungen am Code des Produkts  
Falls Anpassungen am Code des Produkts gewünscht werden, werden die damit verbundenen Aufwände geschätzt und nach Beauftragung separat verrechnet.

Entsprechend dem Open-Source-Gedanken des Produkts werden die Code-Anpassungen als Verbesserungsvorschlag an die Entwickler des Produkts übermittelt.

Um die Wartbarkeit des Produkts sicherzustellen, kann der Service-Provider nur Code unterstützen, der auch offiziell in das Produkt übernommen wird. Jedem Kunden steht es frei, den Code des Produkts selbst anzupassen, in diesem Fall kann der Service-Provider allerdings keine Unterstützung mehr anbieten. Für Details, siehe „Unterstützte Versionen“.
