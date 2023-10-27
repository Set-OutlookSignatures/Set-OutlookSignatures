<!-- omit in toc -->
## **<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures" target="_blank"><img src="/src_Set-OutlookSignatures/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a>**<br>The open source gold standard to centrally manage and deploy email signatures and out of office replies for Outlook and Exchange<br><br><a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures" target="_blank"><img src="https://img.shields.io/github/license/Set-OutlookSignatures/Set-OutlookSignatures" alt="MIT license"></a> <!--XXXRemoveWhenBuildingXXX<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt="this release"></a> XXXRemoveWhenBuildingXXX--> <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/v/release/Set-OutlookSignatures/Set-OutlookSignatures?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="latest release" data-external="1"></a> <a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/Set-OutlookSignatures/Set-OutlookSignatures" alt="open issues" data-external="1"></a> <a href="./Benefactor%20Circle.md" target="_blank"><img src="https://img.shields.io/badge/add%20additional%20features%20and%20support%20with-Benefactor%20Circle-gold" alt="add additional features and support with Benefactor Circle"></a>

# What is the recommended approach for implementing the software? <!-- omit in toc -->
There is certainly no definitive generic recommendation, but this document should be a good starting point.

The content is based on real-life experience implementing the software in multi-client environments with a five-digit number of mailboxes.

It contains proven procedures and recommendations for product managers, architects, operations managers, account managers and mail and client administrators. It is suited for service providers as well as for clients.

It covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The document is available in English and German language.
<br><br>
**Dear businesses using Set-OutlookSignatures:**
- Being Free and Open-Source Software, Set-OutlookSignatures can save you thousands or even tens of thousand Euros/US-Dollars per year in comparison to commercial software.
- Invest in the open-source projects you depend on. Contributors are working behind the scenes to make open-source better for everyone - give them the help and recognition they deserve.
- Sponsor the open-source software your team has built its business on. Fund the projects that make up your software supply chain to improve its performance, reliability, and stability.
- You may consider to become a Benefactor Circle member to unlock additional features: See [`.\docs\Benefactor Circle`](Benefactor%20Circle.md) or [`https://explicitonsulting.at`](https://explicitconsulting.at/open-source/set-outlooksignatures) for details about these features and how you can benefit from them with a Benefactor Circle license.
# Table of Contents  <!-- omit in toc -->
- [1. English](#1-english)
  - [1.1. Overview](#11-overview)
  - [1.2. Manual maintenance of signatures](#12-manual-maintenance-of-signatures)
    - [1.2.1. Signatures in Outlook](#121-signatures-in-outlook)
    - [1.2.2. Signatur in Outlook im Web](#122-signatur-in-outlook-im-web)
  - [1.3. Automatic maintenance of signatures](#13-automatic-maintenance-of-signatures)
    - [1.3.1. Server-based signatures](#131-server-based-signatures)
    - [1.3.2. Client-based signatures](#132-client-based-signatures)
  - [1.4. Criteria](#14-criteria)
  - [1.5. Synchronizing signatures between different devices](#15-synchronizing-signatures-between-different-devices)
  - [1.6. Recommendation: Set-OutlookSignatures](#16-recommendation-set-outlooksignatures)
    - [1.6.1. Allgemeine Beschreibung, Lizenzmodell](#161-allgemeine-beschreibung-lizenzmodell)
    - [1.6.2. Features](#162-features)
  - [1.7. Administration](#17-administration)
    - [1.7.1. Client](#171-client)
    - [1.7.2. Server](#172-server)
    - [1.7.3. Storage of templates](#173-storage-of-templates)
    - [1.7.4. Template management](#174-template-management)
    - [1.7.5. Running the software](#175-running-the-software)
      - [1.7.5.1. Parameters](#1751-parameters)
      - [1.7.5.2. Runtime and Visibility of the software](#1752-runtime-and-visibility-of-the-software)
      - [1.7.5.3. Use of Outlook and Word during runtime.](#1753-use-of-outlook-and-word-during-runtime)
  - [1.8. Support from the service provider.](#18-support-from-the-service-provider)
    - [1.8.1. Consulting and implementation phase](#181-consulting-and-implementation-phase)
      - [1.8.1.1. Initial consultation on textual signatures.](#1811-initial-consultation-on-textual-signatures)
        - [1.8.1.1.1. Participants](#18111-participants)
        - [1.8.1.1.2. content and objectives](#18112-content-and-objectives)
        - [1.8.1.1.3. duration](#18113-duration)
      - [1.8.1.2. training of template administrators](#1812-training-of-template-administrators)
        - [1.8.1.2.1. participants](#18121-participants)
        - [1.8.1.2.2. content and objectives](#18122-content-and-objectives)
        - [1.8.1.2.3. duration](#18123-duration)
        - [1.8.1.2.4. prerequisites](#18124-prerequisites)
      - [1.8.1.3. Client management training](#1813-client-management-training)
        - [1.8.1.3.1. participants](#18131-participants)
        - [1.8.1.3.2. content and objectives](#18132-content-and-objectives)
        - [1.8.1.3.3. duration](#18133-duration)
        - [1.8.1.3.4. prerequisites](#18134-prerequisites)
    - [1.8.2. 1.8.2 Tests, pilot operation, rollout](#182-182-tests-pilot-operation-rollout)
  - [1.9. running operation](#19-running-operation)
    - [1.9.1. 1.9.1 Creating and maintaining templates](#191-191-creating-and-maintaining-templates)
    - [1.9.2. creating and maintaining storage shares for templates and script components](#192-creating-and-maintaining-storage-shares-for-templates-and-script-components)
    - [1.9.3. 1.9.3 Setting and maintaining AD attributes](#193-193-setting-and-maintaining-ad-attributes)
    - [1.9.4. configuration adjustments](#194-configuration-adjustments)
    - [1.9.5. 1.9.5 Problems and questions during operation](#195-195-problems-and-questions-during-operation)
    - [1.9.6. supported versions](#196-supported-versions)
    - [1.9.7. new versions](#197-new-versions)
    - [1.9.8. Adaptations to the code of the product](#198-adaptations-to-the-code-of-the-product)
- [2. Deutsch (German)](#2-deutsch-german)
  - [2.1. Überblick](#21-überblick)
  - [2.2. Manuelle Wartung von Signaturen](#22-manuelle-wartung-von-signaturen)
    - [2.2.1. Signaturen in Outlook](#221-signaturen-in-outlook)
    - [2.2.2. Signatur in Outlook im Web](#222-signatur-in-outlook-im-web)
  - [2.3. Automatische Wartung von Signaturen](#23-automatische-wartung-von-signaturen)
    - [2.3.1. Serverbasierte Signaturen](#231-serverbasierte-signaturen)
    - [2.3.2. Clientbasierte Signaturen](#232-clientbasierte-signaturen)
  - [2.4. Kritierien](#24-kritierien)
  - [2.5. Abgleich von Signaturen zwischen verschiedenen Geräten](#25-abgleich-von-signaturen-zwischen-verschiedenen-geräten)
  - [2.6. Empfehlung: Set-OutlookSignatures](#26-empfehlung-set-outlooksignatures)
    - [2.6.1. Allgemeine Beschreibung, Lizenzmodell](#261-allgemeine-beschreibung-lizenzmodell)
    - [2.6.2. Funktionen](#262-funktionen)
  - [2.7. Administration](#27-administration)
    - [2.7.1. Client](#271-client)
    - [2.7.2. Server](#272-server)
    - [2.7.3. Ablage der Vorlagen](#273-ablage-der-vorlagen)
    - [2.7.4. Verwaltung der Vorlagen](#274-verwaltung-der-vorlagen)
    - [2.7.5. Ausführen des Scripts](#275-ausführen-des-scripts)
      - [2.7.5.1. Parameter](#2751-parameter)
      - [2.7.5.2. Laufzeit und Sichtbarkeit des Scripts](#2752-laufzeit-und-sichtbarkeit-des-scripts)
      - [2.7.5.3. Nutzung von Outlook und Word während der Laufzeit](#2753-nutzung-von-outlook-und-word-während-der-laufzeit)
  - [2.8. Unterstützung durch den Service-Provider](#28-unterstützung-durch-den-service-provider)
    - [2.8.1. Beratungs- und Einführungsphase](#281-beratungs--und-einführungsphase)
      - [2.8.1.1. Erstabstimmung zu textuellen Signaturen](#2811-erstabstimmung-zu-textuellen-signaturen)
        - [2.8.1.1.1. Teilnehmer](#28111-teilnehmer)
        - [2.8.1.1.2. Inhalt und Ziele](#28112-inhalt-und-ziele)
        - [2.8.1.1.3. Dauer](#28113-dauer)
      - [2.8.1.2. Schulung der Vorlagen-Verwalter](#2812-schulung-der-vorlagen-verwalter)
        - [2.8.1.2.1. Teilnehmer](#28121-teilnehmer)
        - [2.8.1.2.2. Inhalt und Ziele](#28122-inhalt-und-ziele)
        - [2.8.1.2.3. Dauer](#28123-dauer)
        - [2.8.1.2.4. Voraussetzungen](#28124-voraussetzungen)
      - [2.8.1.3. Schulung des Clientmanagements](#2813-schulung-des-clientmanagements)
        - [2.8.1.3.1. Teilnehmer](#28131-teilnehmer)
        - [2.8.1.3.2. Inhalt und Ziele](#28132-inhalt-und-ziele)
        - [2.8.1.3.3. Dauer](#28133-dauer)
        - [2.8.1.3.4. Voraussetzungen](#28134-voraussetzungen)
    - [2.8.2. Tests, Pilotbetrieb, Rollout](#282-tests-pilotbetrieb-rollout)
  - [2.9. Laufender Betrieb](#29-laufender-betrieb)
    - [2.9.1. Erstellen und Warten von Vorlagen](#291-erstellen-und-warten-von-vorlagen)
    - [2.9.2. Erstellen und Warten von Ablage-Shares für Vorlagen und Script-Komponenten](#292-erstellen-und-warten-von-ablage-shares-für-vorlagen-und-script-komponenten)
    - [2.9.3. Setzen und Warten von AD-Attributen](#293-setzen-und-warten-von-ad-attributen)
    - [2.9.4. Konfigurationsanpassungen](#294-konfigurationsanpassungen)
    - [2.9.5. Probleme und Fragen im laufenden Betrieb](#295-probleme-und-fragen-im-laufenden-betrieb)
    - [2.9.6. Unterstützte Versionen](#296-unterstützte-versionen)
    - [2.9.7. Neue Versionen](#297-neue-versionen)
    - [2.9.8. Anpassungen am Code des Produkts](#298-anpassungen-am-code-des-produkts)


# 1. English
## 1.1. Overview  
Textual signatures are not only an essential aspect of corporate identity, but together with the disclaimer usually a legal necessity.

This document provides a general overview of signatures, instructions for end users, and details of the service provider's recommended solution for centralised management and automated distribution of textual signatures.

In this document, the word "signature" should always be understood as a textual signature and should not be confused with a digital signature, which serves to encrypt emails and/or legitimise the sender.  
## 1.2. Manual maintenance of signatures  
In manual maintenance, a template for the textual signature is made available to the user, e.g. via the intranet.

Each user sets up the signature himself. Depending on the technical configuration of the client, signatures move with it when the computer used is changed or have to be set up again.

There is no central maintenance.
### 1.2.1. Signatures in Outlook  
In Outlook, practically any number of signatures can be created per mailbox. This is practical, for example, to distinguish between internal and external emails or emails in different languages.

Pro Postfach kann darüber hinaus eine Standard-Signatur für neue emails und eine für Antworten festgelegt werden.   
### 1.2.2. Signatur in Outlook im Web
If you also work with Outlook on the Web, you must set up your signature in Outlook on the Web independently of your signature on the client.

In Outlook on the Web, only one signature is possible unless the mailbox is in Exchange Online and the Roaming Signatures feature has been enabled.
## 1.3. Automatic maintenance of signatures  
The service provider recommends a free script-based solution with central administration and extended client-side functionality, which can be operated and maintained by the customers themselves with the support of the service provider. For details see "Recommendation: Set-OutlookSignatures Benefact Circle".  
### 1.3.1. Server-based signatures  
The biggest advantage of a server-based solution is that every email is captured using a defined rule set, regardless of the application or device from which it was sent.

Since the signature is only attached at the server, the user does not see which signature is used during the creation of an email.

After the signature has been appended at the server, the now modified email must be re-downloaded by the client so that it appears correctly in the Sent Items folder. This generates additional network traffic.

If a message is already digitally signed or encrypted when it is created, the textual signature cannot be added on the server side without breaking the digital signature and encryption. Alternatively, the message is adapted so that the content consists only of the textual signature and the unchanged original message is sent as an attachment.
### 1.3.2. Client-based signatures  
In client-based solutions, templates and application rules for textual signatures are defined in a central repository. A component on the client checks the central configuration during automated or manual invocation and applies it locally.

Client-based solutions, in contrast to server-based solutions, are bound to specific email clients and specific operating systems.

The user already sees the signature during the creation of the email and can adjust it if necessary.

Encryption and digital signing of messages are not a problem on either the client or server side.
## 1.4. Criteria
When evaluating products, the following aspects, among others, should be checked:  
- Can the product handle the number of AD and mail objects in the environment without reproducible crashes or incomplete search results?  
- Does the product have to be installed directly on the mail servers? This means additional dependencies and sources of errors, and can have a negative impact on the availability and reliability of the AD and mail system.  
- Can the administration of the products be delegated directly on the mail servers without granting significant rights?
- Can customers be authorised separately from each other?  
- Can variables in the signatures only be replaced with values of the current user, or also with values of the current mailbox and the respective manager?
- Can a template file be used under different signature names?
- Can templates be distributed in a targeted manner? Generally, by group membership, by email address? Can only be assigned or also forbidden?
- Can the solution handle shared mailboxes?
- Can the solution handle additional mailboxes distributed e.g. by automapping?
- Can images in signatures be shown and hidden under attribute control?
- Can the solution handle roaming signatures in Exchange Online?
- How high are the acquisition and maintenance costs? Are these above the tender limit?
- Do emails have to be redirected to a cloud of the manufacturer?
- Does the SPF record in the DNS need to be adjusted?
  
## 1.5. Synchronizing signatures between different devices  
The signatures in Outlook, Outlook on the Web and other clients (e.g. in smartphone apps) are not synchronised and must therefore be set up separately.

Depending on the client configuration, Outlook signatures may or may not travel with the user between different Windows devices, please contact your local IT for details.

The client-based tool recommended by the service provider can set signatures in Outlook as well as in Outlook on the Web and also offers the user an easy way to transfer existing signatures to other email clients.

The recommended product already supports the roaming signatures of Exchange Online. It can be assumed that mail clients (e.g. smartphone apps) will follow suit in the foreseeable future.
## 1.6. Recommendation: Set-OutlookSignatures  
The service provider recommends the free open source software Set-OutlookSignatures with the chargeable "Benefactor Circle" extension after a survey of the customer requirements and tests of several server- and client-based products and offers its customers support during introduction and operation.  

This document provides an overview of the functional scope and administration of the recommended solution, support of the service provider during introduction and operation, as well as associated expenses.
### 1.6.1. Allgemeine Beschreibung, Lizenzmodell  
<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures" target="_blank">Set-OutlookSignatures</a> is a free open-source product with a chargeable extension for company-relevant functions.

The product is used for the central administration and local distribution of textual signatures and out of office replies to clients. Outlook on Windows is supported as the target platform.

Integration into the client, which is secured with the help of AppLocker and other mechanisms such as Microsoft Purview Informatoin Protection, is technically and organisationally simple thanks to established measures (such as the digital signing of PowerShell scripts).
### 1.6.2. Features
**Signatures and OOF messages can be:**
- Generated from **templates in DOCX or HTML** file format  
- Customized with a **broad range of variables**, including **photos**, from Active Directory and other sources
  - Variables are available for the **currently logged-on user, this user's manager, each mailbox and each mailbox's manager**
  - Images in signatures can be **bound to the existence of certain variables** (useful for optional social network icons, for example)
- Applied to all **mailboxes (including shared mailboxes)**, specific **mailbox groups** or specific **email addresses**, for **every mailbox across all Outlook profiles** (**automapped and additional mailboxes** are optional)  
- Created with different names from the same template (e.g., **one template can be used for multiple shared mailboxes**)
- Assigned **time ranges** within which they are valid  
- Set as **default signature** for new emails, or for replies and forwards (signatures only)  
- Set as **default OOF message** for internal or external recipients (OOF messages only)  
- Set in **Outlook Web** for the currently logged-in user  
- Centrally managed only or **exist along user created signatures** (signatures only)  
- Copied to an **alternate path** for easy access on mobile devices not directly supported by this script (signatures only)
- **Write protected** (Outlook signatures only)
- Mirrored to the cloud as **roaming signatures**

Set-Outlooksignatures can be **executed by users on clients, or on a server without end user interaction**.  
On clients, it can run as part of the logon script, as scheduled task, or on user demand via a desktop icon, start menu entry, link or any other way of starting a program.  
Signatures and OOF messages can also be created and deployed centrally, without end user or client involvement.

**Sample templates** for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

**Simulation mode** allows content creators and admins to simulate the behavior of the software and to inspect the resulting signature files before going live.
  
the software is **designed to work in big and complex environments** (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). It works **on premises, in hybrid and cloud-only environments**.

It is **multi-client capable** by using different template paths, configuration files and script parameters.

Set-OutlookSignatures requires **no installation on servers or clients**. You only need a standard file share on a server, and PowerShell and Office. 

A **documented implementation approach**, based on real life experiences implementing the software in multi-client environments with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and email and client administrators.  
The implementatin approach is **suited for service providers as well as for clients**, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

the software core is **Free and Open-Source Software (FOSS)**. It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see `.\LICENSE.txt` for copyright and MIT license details.

**Some features are exclusive to Benefactor Circle members.** Benefactor Circle members have access to an extension file enabling the exclusive features. This extension file is chargeable, and it is distributed under a proprietary, non-free and non-open-source license.  Please see `.\docs\Benefactor Circle` for details.  
## 1.7. Administration  
### 1.7.1. Client  
- Outlook and Word, each from version 2010  
- the software must run in the security context of the user currently logged in.  
- The PowerShell script must be executed in "Full Language Mode". The "Constrained Language Mode" is not supported, as certain functions such as Base64 conversions are not available in this mode or require very slow alternatives.  
- If AppLocker or comparable solutions are used, the software is already digitally signed.  
- Network unlocks:  
	- Ports 389 (LDAP) and 3268 (Global Catalog), TCP and UDP respectively, must be enabled between the client and all domain controllers. If this is not the case, signature-relevant information and variables cannot be retrieved. the software checks with each run whether access is possible.
- To access the SMB share with the software components, the following ports are needed: 137 UDP, 138 UDP, 139 TCP, 445 TCP (details <a href="https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc731402(v=ws.11)" target="_blank">in this Microsoft article</a>).  
	- Für access to WebDAV shares (e.g. SharePoint document libraries), port 443 TCP is needed.  
### 1.7.2. Server  
Required are:
- An SMB file share in which the software and its components are stored. All users must have read access to this file share and its contents.  
- One or more SMB file shares or WEBDAV shares (e.g. SharePoint document libraries) in which the templates for signatures and out of office replies are stored and managed.

If variables (e.g. first name, last name, phone number) are used in the templates, the corresponding values must be available in the Active Directory. In the case of Linked Mailboxes, a distinction can be made between the attributes of the current user and the attributes of the mailbox located in different AD forests.  

As described in the system requirements, the software and its components must be stored on an SMB file share. Alternatively, it can be distributed to the clients by any mechanism and executed from there.

All users need read access to the software and all its components.

As long as these requirements are met, any SMB file share can be used, for example  
- the NETLOGON share of an Active Directory  
- a share on a Windows server in any architecture (single server or cluster, classic share or DFS in all variations)  
- a share on a Windows client  
- a share on any non-Windows system, e.g. via SAMBA.

As long as all clients use the same version of the software and only configure it via parameters, a central repository for the software components is sufficient.

For maximum performance and flexibility, it is recommended that each client stores the software in its own SMB file share and, if necessary, replicates this across locations on different servers.
### 1.7.3. Storage of templates  
As described in the system requirements, templates for signatures and out of office replies can be stored on SMB file shares or WebDAV shares (e.g. SharePoint document libraries) analogous to the software itself.

SharePoint document libraries have the advantage of optional versioning of files, so that in the event of an error, template administrators can quickly restore an earlier version of a template.

At least one share per client with separate subdirectories for signature and absence templates is recommended.

Users need read access to all templates.

By simply granting write access to the entire template folder or to individual files within it, the creation and management of signature and absence templates is delegated to a defined group of people. Typically, templates are defined, created and maintained by the Corporate Communications and Marketing departments.

For maximum performance and flexibility, it is recommended that each client places the software in its own SMB file share and replicates this across sites to different servers if necessary.  
### 1.7.4. Template management  
By simply assigning write permissions to the template folder or to individual files within it, the creation and management of signature and absence templates is delegated to a defined group of people. Typically, the templates are defined, created and maintained by the Corporate Communications and Marketing departments.

the software can process templates in DOCX or HTML format. For a start, the use of the DOCX format is recommended; the reasons for this recommendation and the advantages and disadvantages of each format are described in the software's `README' file.

The `README` file supplied with the software provides an overview of how to administer templates so that they are  
- apply only to certain groups or mailboxes  
- be set as the default signature for new mails or replies and forwards  
- be set as an internal or external out of office message
- and much more

In `README` and the sample templates, the replaceable variables, the extension with user-defined variables and the handling of photos from the Active Directory are also described.

The sample file "Test all signature replacement variables.docx" provided contains all variables available by default; in addition, custom variables can be defined.
### 1.7.5. Running the software  
the software can be executed via any mechanism, for example  
- when the user logs in as part of the logon script or as a separate script  
- via the task scheduling at fixed times or at certain events  
- by the user himself, e.g. via a shortcut on the desktop  
- by a tool for client administration

Since Set-OutlookSignatures is a pure PowerShell script, it is called like any other script of this file type:  
```
powershell.exe <PowerShell parameter> -file <path to Set-OutlookSignatures.ps1> <Script parameter>  
```
#### 1.7.5.1. Parameters  
The behaviour of the software can be controlled via parameters. Particularly relevant are SignatureTemplatePath and OOFTemplatePath, which are used to specify the path to the signature and absence templates.

The following is an example where the signature templates are on an SMB file share and the out of office provider templates are on a WebDAV share:  
```
powershell.exe -file '\netlogon\set-outlooksignatures\set-outlooksignatures.ps1' -SignatureTemplatePath '\DFS-Share\Common\Templates\Signatures Outlook' -OOFTemplatePath 'https://webdav.example.com/CorporateCommunications/Templates/Out of Office templates'  
```

At the time of writing, other parameters were available. The following is a brief overview of the possibilities, for details please refer to the documentation of the software in the `README` file:  
- SignatureTemplatePath: path to the signature templates. Can be an SMB or WebDAV share.  
- ReplacementVariableConfigFile: Path to the file in which variables deviating from the standard are defined. Can be an SMB or WebDAV share.  
- TrustsToCheckForGroups: By default, all trusts are queried for mailbox information. This parameter can be used to remove specific domains and add non-trusted domains.  
- DeleteUserCreatedSignatures: Should signatures created by the user be deleted? This is not done by default.  
- SetCurrentUserOutlookWebSignature: By default, a signature is set in Outlook on the web for the logged-in user. This parameter can be used to prevent this.  
- SetCurrentUserOOFMessage: By default, the text of the out of office replies is set. This parameter can be used to change this behaviour.  
- OOFTemplatePath: Path to the absence templates. Can be an SMB or WebDAV share.  
- AdditionalSignaturePath: Path to an additional share to which all signatures should be copied, e.g. for access from a mobile device and for simplified configuration of clients not supported by the software. Can be an SMB or WebDAV share.  
- UseHtmTemplates: By default, templates are processed in DOCX format. This switch can be used to switch to HTML (.htm).  
The 'README' file contains further parameters.
#### 1.7.5.2. Runtime and Visibility of the software  
the software is designed for fast runtime and minimal network load. Nevertheless, the runtime of the software depends on many parameters:  
- general speed of the client (CPU, RAM, HDD)  
- Number of mailboxes configured in Outlook  
- Number of trusted domains  
- Response time of the domain controllers and file servers  
- Response time of Exchange servers (setting signatures in Outlook Web, out of office notifications)  
- Number of templates and complexity of variables in them (e.g. photos)

Under the following general conditions, a reproducible runtime of approx. 30 seconds was measured:  
- Standard client  
- Connected to the company network via VPN  
- 4 mailboxes  
- Query of all domains connected via trust  
- 9 signature templates to be processed, all with variables and graphics (but without user photos), partly restricted to groups and mail addresses  
- 8 absence templates to be processed, all with variables and graphics (but without user photos), partly restricted to groups and mail addresses  
- Setting the signature in Outlook on the web  
- No copying of signatures to an additional network path
  
Since the software does not require any user interaction, it can be minimised or hidden using the usual mechanisms. This makes the runtime of the software almost irrelevant.
#### 1.7.5.3. Use of Outlook and Word during runtime.  
the software does not start Outlook, all queries and configurations are done via the file system and the registry.

Outlook can be started, used or closed at will while the software is running.

All changes to signatures and out of office notifications are immediately visible and usable for the user, with one exception: If the name of the default signature to be used for new emails or for replies and forwardings changes, this change will only take effect the next time Outlook is started. If only the content changes, but not the name of one of the default signatures, this change is available immediately.

Word can be started, used or closed at will while the software is running.

the software uses Word to replace variables in DOCX templates and to convert DOCX and HTML to RTF and TXT. Word is started as a separate invisible process. This process can practically not be influenced by the user and does not affect Word processes started by the user.
## 1.8. Support from the service provider.  
The service provider not only recommends the Set-OutlookSignatures software, but also offers its customers defined support free of charge.

Additional support can be obtained after prior agreement for a separate charge.

The central point of contact for all kinds of questions is Mail Product Management.  
### 1.8.1. Consulting and implementation phase  
The following services are covered by the product price:  
#### 1.8.1.1. Initial consultation on textual signatures.  
##### 1.8.1.1.1. Participants  
- Customer: corporate communications, marketing, client management, project coordinator  
- Service provider: mail product management, mail operations management or mail architecture  
##### 1.8.1.1.2. content and objectives  
- Customer: Presentation of own wishes regarding textual signatures  
- Service provider: Brief description of the basic options for textual signatures, advantages and disadvantages of the different approaches, reasons for deciding on the recommended product.  
- Comparison of customer requirements with technical and organizational possibilities  
- Live demonstration of the product, taking customer requirements into account  
- Determination of the next steps  
##### 1.8.1.1.3. duration  
4 hours
#### 1.8.1.2. training of template administrators  
##### 1.8.1.2.1. participants  
- Customer: template administrators (corporate communications, marketing, analysts), optional client management, project coordinator.  
- Service provider: mail product management, mail operations management, or mail architecture.  
##### 1.8.1.2.2. content and objectives  
- Summary of the previous meeting "Initial coordination on textual signatures", with focus on desired and feasible functions  
- Presentation of the structure of the template directories, with focus on  
- naming conventions  
- Application order (general, group-specific, mailbox-specific, alphabetical in each group)  
- Definition of default signatures for new emails and for replies and forwards  
- Definition of out of office texts for internal and external recipients.  
- Determination of the temporal validity of templates  
- Variables and user photos in templates  
- Differences between DOCX and HTML formats  
- Possibilities for the integration of a disclaimer  
- Joint development of initial templates based on existing templates and customer requirements  
- Live demonstration on a standard client with a test user and test mailboxes of the customer (see requirements)  
##### 1.8.1.2.3. duration  
4 hours
##### 1.8.1.2.4. prerequisites  
- The customer provides a standard client with Outlook and Word.  
- The screen content of the client must be able to be projected by a beamer or displayed on an appropriately large monitor for collaborative work.  
- The customer provides a test user. This test user must be able to run script files on the standard client  
	- be allowed to download script files from the Internet (github.com) once (alternatively, the customer can provide a BitLocker-encrypted USB stick for data transfer).  
	- be allowed to run unsigned PowerShell scripts in full language mode  
	- have a mailbox  
	- have full access to various test mailboxes (personal mailboxes or group mailboxes) that are, if possible, direct or indirect members of various groups or distribution lists. For full access, the user may be authorized to the other mailboxes accordingly, or username and password of the additional mailboxes are known.  
#### 1.8.1.3. Client management training  
##### 1.8.1.3.1. participants  
- Customer: client management, optionally an administrator of the Active Directory, optionally an administrator of the file server and/or SharePoint server, optionally corporate communications and marketing, coordinator of the project  
- Service Provider: mail product management, mail operations management, or mail architecture, a representative of the client team at appropriate clients
##### 1.8.1.3.2. content and objectives  
- Summary of the previous meeting "Initial agreement on textual signatures", with focus on desired and feasible functions  
- Presentation of the possibilities with focus on  
- Basic flow of the software  
- System requirements client (Office, PowerShell, AppLocker, digital signature of the software, network ports)  
- System requirements server (storage of the templates)  
- Possibilities of product integration (logon script, scheduled task, desktop shortcut)  
- Parameterization of the software, among others:  
- Disclosure of template folders  
- Consider Outlook on the web?  
- Consider out of office replies?  
- Which trusts to take into account?  
- How to define additional variables?  
- Allow user created signatures?  
- Place signatures on an additional path?  
- Joint testing based on templates previously developed by the customer and customer requirements.  
- Definition of next steps  
##### 1.8.1.3.3. duration  
4 hours 
##### 1.8.1.3.4. prerequisites  
- The customer provides a standard client with Outlook and Word.  
- The screen content of the client must be able to be projected via beamer or displayed on an appropriately large monitor for collaborative work.  
- The customer provides a test user. This test user must be able to run on the standard client  
	- be allowed to download script files from the Internet (github.com) once (alternatively, the customer can provide a BitLocker-encrypted USB stick for data transfer).  
	- be allowed to run unsigned PowerShell scripts in full language mode
	- have a mailbox  
	- have full access to various test mailboxes (personal mailboxes or group mailboxes) that are, if possible, direct or indirect members of various groups or distribution lists. For full access, the user may be authorized to the other mailboxes accordingly, or the user name and password of the additional mailboxes are known.  
- Customer shall provide at least one central SMB file or WebDAV share for template storage.  
- Customer shall provide a central SMB file share for the storage of the software and its components.
### 1.8.2. 1.8.2 Tests, pilot operation, rollout  
The customer's project manager is responsible for planning and coordinating tests, pilot operation and rollout.

The concrete technical implementation is carried out by the customer. If, in addition to mail, the client is also supported by service providers, the client team will assist with the integration of the software (logon script, scheduled task, desktop shortcut).

In the event of fundamental technical problems, the Mail product management team provides support in researching the causes, prepares proposals for solutions and, if necessary, establishes contact with the manufacturer of the product.

The creation and maintenance of templates is the responsibility of the customer.

For the procedure for adjustments to the code or the release of new functions, see the "Ongoing Operations" chapter.
## 1.9. running operation  
### 1.9.1. 1.9.1 Creating and maintaining templates  
Creating and maintaining templates is the responsibility of the customer.  
Mail Product Management is available to advise on feasibility and impact issues.
### 1.9.2. creating and maintaining storage shares for templates and script components  
The creation and maintenance of storage shares for templates and script components is the responsibility of the customer.

Mail Product Management is available to advise on feasibility and implications.  
### 1.9.3. 1.9.3 Setting and maintaining AD attributes  
Setting and maintaining AD attributes related to textual signatures (e.g., attributes for variables, user photos, group memberships) is the customer's responsibility.

Mail Product Management is available to advise on feasibility and impact issues.
### 1.9.4. configuration adjustments  
Configuration adjustments explicitly provided for by the developers of the software are supported at any time.

Mail product management is available to advise on the feasibility and impact of desired customizations.

The planning and coordination of tests, pilot operation and rollout in connection with configuration adjustments is carried out by the customer, as is the concrete technical implementation.

If, in addition to mail, the client is also supported by the service provider, the client team provides support with the integration of the software (logon script, scheduled task, desktop shortcut).  
### 1.9.5. 1.9.5 Problems and questions during operation  
In the event of fundamental technical problems, Mail Product Management provides support in researching the causes, works out proposed solutions and, if necessary, establishes contact with the manufacturer of the product.

Mail product management is also available to answer general questions about the product and its possible applications.
### 1.9.6. supported versions  
The version numbers of the product follow the specifications of Semantic Versioning and are therefore structured according to the "Major.Minor.Patch" format.  
- "Major" is incremented when there is no compatibility with previous versions.  
- "Minor" is incremented when new features compatible with previous versions are introduced.  
- "Patch" is incremented when changes include only bug fixes compatible with previous versions.  
- Additionally, pre-release and build metadata identifiers are available as attachments to the "Major.Minor.Patch" format, e.g. "-Beta1".

Service Provider Supported Versions:  
- The highest version of the product released by the service provider, regardless of its release date.  
- Support for a released version automatically ends three months after a higher version is released.

This means that customers have three months after a new version is released to upgrade to that version before service provider support for previously released versions expires.

This means that no more than one update is ever required in a 3-month period. This protects both customers and service providers from gross errors in product development.
### 1.9.7. new versions  
When new versions of the product are released, Mail Product Management informs customer-defined contacts of the changes associated with that version, potential impacts on the existing configuration, and identifies upgrade options.

Planning and coordination of the rollout of the new version is done by the customer contact.

The concrete technical implementation is also carried out by the customer. If, in addition to mail, the client is also supported by service providers, the client team provides support in integrating the software (logon script, scheduled task, desktop shortcut).

In the event of fundamental technical problems, Mail product management provides support in researching the causes, works out proposals for solutions and, if necessary, establishes contact with the manufacturer of the product.
### 1.9.8. Adaptations to the code of the product  
If adjustments to the product's code are desired, the associated effort will be estimated and charged separately after commissioning.

In accordance with the open source nature of the Product, the code adjustments will be submitted to the developers of the Product as a suggestion for improvement.

To ensure the maintainability of the product, the service provider can only support code that is also officially adopted into the product. Each customer is free to customize the product's code themselves, but in this case the service provider can no longer provide support. For details, see "Supported versions".
  
# 2. Deutsch (German)  
## 2.1. Überblick  
Textuelle Signaturen sind nicht nur ein wesentlicher Aspekt der Corporate Identity, sondern gemeinsam mit dem Disclaimer im Regelfall eine rechtliche Notwendigkeit.

Dieses Dokument bietet einen generellen Überblick über Signaturen, Anleitungen für Endbenutzer, sowie Details zur vom Service-Provider empfohlenen Lösung zur zentralen Verwaltung und automatisierten Verteilung von textuellen Signaturen.

Das Wort "Signatur" ist in diesem Dokument immer als textuelle Signatur zu verstehen und nicht mit einer digitalen Signatur, die der Verschlüsselung von emails und/oder der Legitimierung des Absenders dient, zu verwechseln.  
## 2.2. Manuelle Wartung von Signaturen  
Bei der manuellen Wartung wird dem Benutzer z. B. über das Intranet eine Vorlage für die textuelle Signatur zur Verfügung gestellt.

Jeder Benutzer richtet sich die Signatur selbst ein. Je nach technischer Konfiguration des Clients wandern Signaturen bei einem Wechsel des verwendeten Computers mit oder sind neu einzurichten.

Eine zentrale Wartung gibt es nicht.
### 2.2.1. Signaturen in Outlook  
In Outlook können pro Postfach praktisch beliebig viele Signaturen erstellt werden. Dies ist beispielsweise praktisch, um zwischen internen und externen emails, oder emails in verschiedenen Sprachen zu unterscheiden.

Pro Postfach kann darüber hinaus eine Standard-Signatur für neue emails und eine für Antworten festgelegt werden.   
### 2.2.2. Signatur in Outlook im Web  
Falls Sie auch mit Outlook im Web arbeiten, müssen Sie sich unabhängig von Ihrer Signatur am Client Ihre Signatur in Outlook im Web einrichten:  
1. Melden Sie sich in einem Webbrowser auf <a href="https://mail.example.com" target="_blank">https<area>://mail.example.com</a> an. Geben Sie Ihren Benutzernamen und Ihr Kennwort ein, und klicken Sie dann auf Anmelden.  
2. Wählen Sie auf der Navigationsleiste Einstellungen > Optionen aus.  
3. Wählen Sie unter Optionen den Befehl Einstellungen > email aus.  
4. Geben Sie im Textfeld unter email-Signatur die Signatur ein, die Sie verwenden möchten. Verwenden Sie die Minisymbolleiste "Formatieren", um das Aussehen der Signatur zu ändern.  
5. Wenn Ihre Signatur automatisch am Ende aller ausgehenden Nachrichten angezeigt werden soll, und zwar auch in Antworten und weitergeleiteten Nachrichten, aktivieren Sie Signatur automatisch in meine gesendeten Nachrichten einschließen. Wenn Sie diese Option nicht aktivieren, können Sie Ihre Signatur jeder Nachricht manuell hinzufügen.  
6. Klicken Sie auf Speichern.

In Outlook im Web ist nur eine einzige Signatur möglich, außer das Postfach befindet sich in Exchange Online und die Funktion Roaming Signatures wurde aktiviert.
## 2.3. Automatische Wartung von Signaturen  
Der Service-Provider empfiehlt eine kostenlose scriptbasierte Lösung mit zentraler Verwaltung und erweitertem clientseitigen Funktionsumfang, die mit Unterstützung des Service-Providers von den Kunden selbst betrieben und gewartet werden kann. Details siehe "Empfehlung: Set-OutlookSignatures Benefact Circle".  
### 2.3.1. Serverbasierte Signaturen  
Der größte Vorteil einer serverbasierten Lösung ist, dass an Hand eines definierten Regelsets jedes email erfasst wird, ganz gleich, von welcher Applikation oder welchem Gerät es verschickt wurde.

Da die Signatur erst am Server angehängt wird, sieht der Benutzer während der Erstellung eines emails nicht, welche Signatur verwendet wird.

Nachdem die Signatur am Server angehängt wurde, muss das nun veränderte email vom Client neu heruntergeladen werden, damit es im Ordner „Gesendete Elemente“ korrekt angezeigt wird. Das erzeugt zusätzlichen Netzwerkverkehr.

Wird eine Nachricht schon bei Erstellung digital signiert oder verschlüsselt, kann die textuelle Signatur serverseitig nicht hinzugefügt werden, ohne die digitale Signatur und die Verschlüsselung zu brechen. Alternativ wird die Nachricht so angepasst, dass der Inhalt nur aus der textuellen Signatur besteht und unveränderte ursprüngliche Nachricht als Anhang mitgeschickt wird.
### 2.3.2. Clientbasierte Signaturen  
Bei clientbasierten Lösungen werden in einer zentralen Ablage Vorlagen und Anwendungsregeln für textuelle Signaturen definiert. Eine Komponente am Client prüft bei automatisiertem oder manuellen Aufruf die zentrale Konfiguration und wendet sie lokal an.

Clientbasierte Lösungen sind im Gegensatz zu serverbasierten Lösungen an bestimmte email-Clients und bestimmte Betriebssysteme gebunden.

Der Benutzer sieht die Signatur bereits während der Erstellung des emails und kann diese gegebenenfalls anpassen.

Die Verschlüsselung und das digitale Signieren von Nachrichten stellen weder client- noch serverseitig ein Problem dar.
## 2.4. Kritierien
Bei der Evaluierung von Produkten sollten unter anderem folgende Aspekte geprüft werden:  
- Kann das Produkt mit der Anzahl der AD- und Mail-Objekte in der Umgebung ohne reproduzierbare Abstürze oder unvollständige Suchergebnissen umgehen?  
- Muss das Produkt direkt auf den Mail-Servern installiert werden? Das bedeutet zusätzliche Abhängigkeiten und Fehlerquellen, und kann sich negativ auf Verfügbarkeit und Zuverlässigkeit des AD- und Mail-Systems auswirken.  
- Kann die Administration der Produkte ohne Vergabe erheblicher Rechte direkt auf den Mail-Servern delegiert werden?
- Können Kunden separat voneinander berechtigt werden?  
- Können Variablen in den Signaturen nur mit Werten des aktuellen Benutzers ersetzt werden, oder auch mit Werten des aktuellen Postfachs und des jeweiligen Managers?
- Kann eine Vorlagen-Datei unter verschiedenen Signatur-Namen verwendet werden?
- Können Vorlagen zielgerichtet verteilt werden? Allgemein, nach Gruppenzugehörigkeit, nach email-Adresse? Kann nur zugewiesen oder auch verboten werden?
- Kann die Lösung mit gemeinsam verwendeten Postfächern umgehen?
- Kann die Lösung mit zusätzlichen Postfächern umgehen, die z. B. per Automapping verteilt wurden?
- Können Bilder in Signaturen attributgesteuert ein- und ausgeblendet werden?
- Kann die Lösung mit Roaming Signatures in Exchange Online umgehen?
- Wie hoch sind die Anschaffungs- und Wartungskosten? Liegen diese über der Ausschreibungsgrenze?
- Müssen emails in eine Cloud des Herstellers umgeleitet werden?
- Muss der SPF-Eintrag im DNS angepasst werden?
  
## 2.5. Abgleich von Signaturen zwischen verschiedenen Geräten  
Die Signaturen in Outlook, Outlook im Web und anderen Clients (z. B. in Smartphone-Apps) sind nicht synchronisiert und müssen daher separat eingerichtet werden.

Je nach Client-Konfiguration wandern Outlook-Signaturen mit dem Benutzer zwischen verschiedenen Windows-Geräten mit oder nicht, für Details wenden Sie sich bitte an Ihre lokale IT.

Das vom Service-Provider empfohlene clientbasierte Werkzeug kann Signaturen sowohl in Outlook als auch in Outlook im Web setzen und bietet dem Benutzer darüber hinaus eine einfache Möglichkeit zur Übernahme bestehender Signaturen in weitere email-Clients an.

Das empfohlene Produkt unterstützt bereits die Roaming Signatures von Exchange Online. Es ist davon auszugehen, dass Mail-Clients (z. B. Smartphone-Apps) in absehbarer Zeit nachziehen.
## 2.6. Empfehlung: Set-OutlookSignatures  
Der Service-Provider empfiehlt nach einer Erhebung der Kundenanforderungen und Tests mehrerer server- und clientbasierten Produkte die kostenlose Open-Source-Software Set-OutlookSignatures mit der kostenpflichtigen "Benefactor Circle"-Erweiterung und bietet seinen Kunden Unterstützung bei Einführung und Betrieb an.  

Dieses Dokument bietet einen Überblick über Funktionsumfang und Administration der empfohlenen Lösung, Unterstützung des Service-Providers bei Einführung und Betrieb, sowie damit verbundene Aufwände.  
### 2.6.1. Allgemeine Beschreibung, Lizenzmodell  
<a href="https://github.com/Set-OutlookSignatures/Set-OutlookSignatures" target="_blank">Set-OutlookSignatures</a> ist ein kostenloses Open-Source-Produkt mit einer kostenpflichtigen Erweiterung für unternehmensrelevante Funktionen.

Das Produkt dient der zentralen Verwaltung und lokalen Verteilung textueller Signaturen und Abwesenheits-Nachrichten auf Clients. Als Zielplattform wird dabei Outlook auf Windows unterstützt.

Die Einbindung in den mit Hilfe von AppLocker und anderen Mechanismen wie z. B. Microsoft Purview Informatoin Protection abgesicherten Client ist durch etablierte Maßnahmen (wie z. B. dem digitalen Signieren von PowerShell-Skripten) technisch und organisatorisch einfach möglich.  
### 2.6.2. Funktionen  
**Signaturen und OOF-Nachrichten können:**
- Aus **Vorlagen im DOCX- oder HTML**-Dateiformat generiert werden  
- Mit einer **großen Auswahl an Variablen**, einschließlich **Fotos**, aus Active Directory und anderen Quellen angepasst werden
  - Variablen sind für den **aktuell angemeldeten Benutzer, den Manager dieses Benutzers, jedes Postfach und den Manager jedes Postfachs** verfügbar
  - Bilder in Signaturen können **an das Vorhandensein bestimmter Variablen** gebunden werden (nützlich z. B. für optionale Icons sozialer Netzwerke)
- Angewandt werden auf alle **Postfächer (einschließlich gemeinsam genutzter Postfächer)**, bestimmte **Postfachgruppen** oder bestimmte **email-Adressen**, für **jedes Postfach in allen Outlook-Profilen** (**automatisierte und zusätzliche Postfächer** sind optional)  
- Mit unterschiedlichen Namen aus derselben Vorlage erstellt (z.B. **eine Vorlage kann für mehrere gemeinsame Postfächer verwendet werden**)
- Innerhalb zugewiesener **Zeitbereiche**, innerhalb derer sie gültig sind, verwendet werden  
- Als **Standardsignatur** für neue emails oder für Antworten und Weiterleitungen festgelegt werden (nur Signaturen)  
- Als **Standard-OOF-Nachricht** für interne oder externe Empfänger festgelegt werden (nur OOF-Nachrichten)  
- Nach **Outlook Web** für den aktuell angemeldeten Benutzer synchronisiert werden  
- Nur zentral verwaltet sein oder **mit vom Benutzer erstellten Signaturen** koexistieren (nur Signaturen)  
- In einen **alternativen Pfad** kopiert werden für einfachen Zugriff auf mobilen Geräten, die nicht direkt von diesem Skript unterstützt werden (nur Signaturen)
- **Schreibgeschützt** (nur Outlook-Signaturen)
- Gespiegelt in die Cloud als **Roaming-Signaturen**

Set-Outlooksignatures kann **von Benutzern auf Clients oder auf einem Server ohne Interaktion des Endbenutzers** ausgeführt werden.  
Auf den Clients kann es als Teil des Anmeldeskripts, als geplante Aufgabe oder auf Wunsch des Benutzers über ein Desktop-Symbol, einen Startmenüeintrag, eine Verknüpfung oder eine andere Art des Programmstarts ausgeführt werden.  
Signaturen und OOF-Nachrichten können auch zentral erstellt und bereitgestellt werden, ohne dass der Endbenutzer oder der Client beteiligt sind.

**Beispielvorlagen** für Signaturen und OOF-Nachrichten demonstrieren alle verfügbaren Funktionen und werden als .docx- und .htm-Dateien bereitgestellt.

Der **Simulationsmodus** ermöglicht es Inhaltserstellern und Administratoren, das Verhalten des Skripts zu simulieren und die resultierenden Signaturdateien zu überprüfen, bevor sie in Betrieb gehen.
  
Das Skript ist **für den Einsatz in großen und komplexen Umgebungen** (Exchange Resource Forest-Szenarien, AD-übergreifende Trusts, mehrstufige AD-Subdomänen, viele Objekte) konzipiert. Es funktioniert **vor Ort, in hybriden und reinen Cloud-Umgebungen**.

Es ist **multimandantenfähig** durch die Verwendung verschiedener Vorlagenpfade, Konfigurationsdateien und Skriptparameter.

Set-OutlookSignatures erfordert **keine Installation auf Servern oder Clients**. Sie benötigen lediglich eine Standard-Dateifreigabe auf einem Server sowie PowerShell und Office.
## 2.7. Administration  
### 2.7.1. Client  
- Outlook und Word, jeweils ab Version 2010  
- Das Script muss im Sicherheitskontext des aktuell angemeldeten Benutzers laufen.  
- Das PowerShell-Script muss im „Full Language Mode” ausgeführt werden. Der „Constrained Language Mode“ wird nicht unterstützt, da gewisse Funktionen wie z. B. Base64-Konvertierungen in diesem Modus nicht verfügbar sind oder sehr langsame Alternativen benötigen.  
- Falls AppLocker oder vergleichbare Lösungen zum Einsatz kommen, ist das Script bereits digital signiert.  
- Netzwerkfreischaltungen:  
	- Die Ports 389 (LDAP) and 3268 (Global Catalog), jeweils TCP and UDP, müssen zwischen Client und allen Domain Controllern freigeschaltet sein. Falls dies nicht der Fall ist, können signaturrelevante Informationen und Variablen nicht abgerufen werden. Das Script prüft bei jedem Lauf, ob der Zugriff möglich ist.  
	- Für den Zugriff auf den SMB-File-Share mit den Script-Komponenten werden folgende Ports benötigt: 137 UDP, 138 UDP, 139 TCP, 445 TCP (Details <a href="https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc731402(v=ws.11)" target="_blank">in diesem Microsoft-Artikel</a>).  
	- Für den Zugriff auf WebDAV-Shares (z. B. SharePoint Dokumentbibliotheken) wird Port 443 TCP benötigt.  
### 2.7.2. Server  
Benötigt werden:
- Ein SMB-File-Share, in den das Script und seine Komponenten abgelegt werden. Auf diesen File-Share und seine Inhalte müssen alle Benutzer lesend zugreifen können.  
- Ein oder mehrere SMB-File-Shares oder WEBDAV-Shares (z. B. SharePoint Dokumentbibliotheken), in den die Vorlagen für Signaturen und Abwesenheitsnachrichten gespeichert und verwaltet werden.

Falls in den Vorlagen Variablen (z. B. Vorname, Nachname, Telefonnummer) genutzt werden, müssen die entsprechenden Werte im Active Directory vorhanden sein. Im Fall von Linked Mailboxes kann dabei zwischen den Attributen des aktuellen Benutzers und den Attributen des Postfachs, die sich in unterschiedlichen AD-Forests befinden, unterschieden werden.  

Wie in den Systemanforderungen beschrieben, ist das Script samt seinen Komponenten auf einem SMB-File-Share abzulegen. Alternativ kann es durch einen beliebigen Mechanismus auf die Clients verteilt und von dort ausgeführt werden.

Alle Benutzer benötigen Lesezugriff auf das Script und alle seine Komponenten.

Solange diese Anforderungen erfüllt sind, kann jeder beliebige SMB-File-Share genutzt werden, beispielsweise  
- der NETLOGON-Share eines Active Directory  
- ein Share auf einem Windows-Server in beliebiger Architektur (einzelner Server oder Cluster, klassischer Share oder DFS in allen Variationen)  
- ein Share auf einem Windows-Client  
- ein Share auf einem beliebigen Nicht-Windows-System, z. B. über SAMBA

Solange alle Kunden die gleiche Version des Scripts einsetzen und dieses nur über Parameter konfigurieren, genügt eine zentrale Ablage für die Script-Komponenten.

Für maximale Leistung und Flexibilität wird empfohlen, dass jeder Kunde das Script in einem eigenen SMB-File-Share ablegt und diesen gegebenenfalls über Standorte hinweg auf verschiedene Server repliziert.  
### 2.7.3. Ablage der Vorlagen  
Wie in den Systemanforderungen beschrieben, können Vorlagen für Signaturen und Abwesenheitsnachrichten analog zum Script selbst auf SMB-File-Shares oder WebDAV-Shares (z. B. SharePoint Dokumentbibliotheken) abgelegt werden.

SharePoint-Dokumentbibliotheken haben den Vorteil der optionalen Versionierung von Dateien, so dass im Fehlerfall durch die Vorlagen-Verwalter rasch eine frühere Version einer Vorlage wiederhergestellt werden kann.

Es wird pro Kunde zumindest ein Share mit separaten Unterverzeichnissen für Signatur- und Abwesenheits-Vorlagen empfohlen.

Benutzer benötigen lesenden Zugriff auf alle Vorlagen.

Durch simple Vergabe von Schreibrechten auf den gesamten Vorlagen-Ordner oder auf einzelne Dateien darin wird die Erstellung und Verwaltung von Signatur- und Abwesenheits-Vorlagen an eine definierte Gruppe von Personen delegiert. Üblicherweise werden die Vorlagen von den Abteilungen Unternehmenskommunikation und Marketing definiert, erstellt und gewartet.

Für maximale Leistung und Flexibilität wird empfohlen, dass jeder Kunde das Script in einem eigenen SMB-File-Share ablegt und diesen gegebenenfalls über Standorte hinweg auf verschiedene Server repliziert.  
### 2.7.4. Verwaltung der Vorlagen  
Durch simple Vergabe von Schreibrechten auf den Vorlagen-Ordner oder auf einzelne Dateien darin wird die Erstellung und Verwaltung von Signatur- und Abwesenheits-Vorlagen an eine definierte Gruppe von Personen delegiert. Üblicherweise werden die Vorlagen von den Abteilungen Unternehmenskommunikation und Marketing definiert, erstellt und gewartet.

Das Script kann Vorlagen im DOCX- oder im HTML-Format verarbeiten. Für den Anfang wird die Verwendung des DOCX-Formats empfohlen; die Gründe für diese Empfehlung und die Vor- und Nachteile des jeweiligen Formats werden in der `README`-Datei des Scripts beschrieben.

Die mit dem Script mitgelieferte `README`-Datei bietet eine Übersicht, wie Vorlagen zu administrieren sind, damit sie  
- nur für bestimmte Gruppen oder Postfächer gelten  
- als Standard-Signatur für neue Mails oder Antworten und Weiterleitungen gesetzt werden  
- als interne oder externe Abwesenheits-Nachricht gesetzt werden
- und vieles mehr

In `README` und den Beispiel-Vorlagen werden zudem die ersetzbaren Variablen, die Erweiterung um benutzerdefinierte Variablen und der Umgang mit Fotos aus dem Active Directory beschrieben.

In der mitgelieferten Beispiel-Datei „Test all signature replacement variables.docx“ sind alle standardmäßig verfügbaren Variablen enthalten; zusätzlich können eigene Variablen definiert werden.
### 2.7.5. Ausführen des Scripts  
Das Script kann über einen beliebigen Mechanismus ausgeführt werden, beispielsweise  
- bei Anmeldung des Benutzers als Teil des Logon-Scripts oder als eigenes Script  
- über die Aufgabenplanung zu fixen Zeiten oder bei bestimmten Ereignissen  
- durch den Benutzer selbst, z. B. über eine Verknüpfung auf dem Desktop  
- durch ein Werkzeug zur Client-Verwaltung

Da es sich bei Set-OutlookSignatures um ein reines PowerShell-Script handelt, erfolgt der Aufruf wie bei jedem anderen Script dieses Dateityps:  
```
powershell.exe <PowerShell-Parameter> -file <Pfad zu Set-OutlookSignatures.ps1> <Script-Parameter>  
```
#### 2.7.5.1. Parameter  
Das Verhalten des Scripts kann über Parameter gesteuert werden. Besonders relevant sind dabei SignatureTemplatePath und OOFTemplatePath, über die der Pfad zu den Signatur- und Abwesenheits-Vorlagen angegeben wird.

Folgend ein Beispiel, bei dem die Signatur-Vorlagen auf einem SMB-File-Share und die AbwesenheService-Providerorlagen auf einem WebDAV-Share liegen:  
```
powershell.exe -file '\\example.com\netlogon\set-outlooksignatures\set-outlooksignatures.ps1' –SignatureTemplatePath '\\example.com\DFS-Share\Common\Templates\Signatures Outlook' –OOFTemplatePath 'https://webdav.example.com/CorporateCommunications/Templates/Out of Office templates'  
```

Zum Zeitpunkt der Erstellung dieses Dokuments waren noch weitere Parameter verfügbar. Folgend eine kurze Übersicht der Möglichkeit, für Details sei auf die Dokumentation des Scripts in der `README`-Datei verwiesen:  
- SignatureTemplatePath: Pfad zu den Signatur-Vorlagen. Kann ein SMB- oder WebDAV-Share sein.  
- ReplacementVariableConfigFile: Pfad zur Datei, in der vom Standard abweichende Variablen definiert werden. Kann ein SMB- oder WebDAV-Share sein.  
- TrustsToCheckForGroups: Standardmäßig werden alle Trusts nach Postfachinformationen abgefragt. Über diesen Parameter können bestimmte Domains entfernt und nicht-getrustete Domains hinzugefügt werden.  
- DeleteUserCreatedSignatures: Sollen vom Benutzer selbst erstelle Signaturen gelöscht werden? Standardmäßig erfolgt dies nicht.  
- SetCurrentUserOutlookWebSignature: Standardmäßig wird für den angemeldeten Benutzer eine Signatur in Outlook im Web gesetzt. Über diesen Parameter kann das verhindert werden.  
- SetCurrentUserOOFMessage: Standardmäßig wird der Text der Abwesenheits-Nachrichten gesetzt. Über diesen Parameter kann dieses Verhalten geändert werden.  
- OOFTemplatePath: Pfad zu den Abwesenheits-Vorlagen. Kann ein SMB- oder WebDAV-Share sein.  
- AdditionalSignaturePath: Pfad zu einem zusätzlichen Share, in den alle Signaturen kopiert werden sollen, z. B. für den Zugriff von einem mobilen Gerät aus und zur vereinfachten Konfiguration nicht vom Script unterstützter Clients. Kann ein SMB- oder WebDAV-Share sein.  
- UseHtmTemplates: Standardmäßig werden Vorlagen im DOCX-Format verarbeitet. Über diesen Schalter kann auf HTML (.htm) umgeschaltet werden.  
Die `README`-Datei enthält weitere Parameter.
#### 2.7.5.2. Laufzeit und Sichtbarkeit des Scripts  
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
  
Da das Script keine Benutzerinteraktion erfordert, kann es über die üblichen Mechanismen minimiert oder versteckt ausgeführt werden. Die Laufzeit des Script wird dadurch nahezu irrelevant.  
#### 2.7.5.3. Nutzung von Outlook und Word während der Laufzeit  
Das Script startet Outlook nicht, alle Abfragen und Konfigurationen erfolgen über das Dateisystem und die Registry.

Outlook kann während der Ausführung des Scripts nach Belieben gestartet, verwendet oder geschlossen werden.

Sämtliche Änderungen an Signaturen und Abwesenheits-Benachrichtigungen sind für den Benutzer sofort sichtbar und verwendbar, mit einer Ausnahme: Falls sich der Name der zu verwendenden Standard-Signatur für neue emails oder für Antworten und Weiterleitungen ändert, so greift diese Änderung erst beim nächsten Start von Outlook. Ändert sich nur der Inhalt, aber nicht der Name einer der Standard-Signaturen, so ist diese Änderung sofort verfügbar.

Word kann während der Ausführung des Scripts nach Belieben gestartet, verwendet oder geschlossen werden.

Das Script nutzt Word zum Ersatz von Variablen in DOCX-Vorlagen und zum Konvertieren von DOCX und HTML nach RTF und TXT. Word wird dabei als eigener unsichtbarer Prozess gestartet. Dieser Prozess kann vom Benutzer praktisch nicht beeinflusst werden und beeinflusst vom Benutzer gestartete Word-Prozesse nicht.  
## 2.8. Unterstützung durch den Service-Provider  
Der Service-Provider empfiehlt die Software Set-OutlookSignatures nicht nur, sondern bietet seinen Kunden auch definierte kostenlose Unterstützung an.

Darüberhinausgehende Unterstützung kann nach vorheriger Abstimmung gegen separate Verrechnung bezogen werden.

Zentrale Anlaufstelle für Fragen aller Art ist das Mail-Produktmanagement.  
### 2.8.1. Beratungs- und Einführungsphase  
Folgende Leistungen sind mit dem Produktpreis abgedeckt:  
#### 2.8.1.1. Erstabstimmung zu textuellen Signaturen  
##### 2.8.1.1.1. Teilnehmer  
- Kunde: Unternehmenskommunikation, Marketing, Clientmanagement, Koordinator des Vorhabens  
- Service-Provider: Mail-Produktmanagement, Mail-Betriebsführung oder Mail-Architektur  
##### 2.8.1.1.2. Inhalt und Ziele  
- Kunde: Vorstellung der eigenen Wünsche zu textuellen Signaturen  
- Service-Provider: Kurze Beschreibung zu prinzipiellen Möglichkeiten rund um textuelle Signaturen, Vor- und Nachteile der unterschiedlichen Ansätze, Gründe für die Entscheidung zum empfohlenen Produkt  
- Abgleich der Kundenwünsche mit den technisch-organisatorischen Möglichkeiten  
- Live-Demonstration des Produkts unter Berücksichtigung der Kundenwünsche  
- Festlegung der nächsten Schritte  
##### 2.8.1.1.3. Dauer  
4 Stunden  
#### 2.8.1.2. Schulung der Vorlagen-Verwalter  
##### 2.8.1.2.1. Teilnehmer  
- Kunde: Vorlagen-Verwalter (Unternehmenskommunikation, Marketing, Analytiker), optional Clientmanagement, Koordinator des Vorhabens  
- Service-Provider: Mail-Produktmanagement, Mail-Betriebsführung oder Mail-Architektur  
##### 2.8.1.2.2. Inhalt und Ziele  
- Zusammenfassung des vorangegangenen Termins „Erstabstimmung zu textuellen Signaturen“, mit Fokus auf gewünschte und realisierbare Funktionen  
- Vorstellung des Aufbaus der Vorlagen-Verzeichnisse, mit Fokus auf  
- Namenskonventionen  
- Anwendungsreihenfolge (allgemein, gruppenspezifisch, postfachspezifisch, in jeder Gruppe alphabetisch)  
- Festlegung von Standard-Signaturen für neue emails und für Antworten und Weiterleitungen  
- Festlegung von Abwesenheits-Texten für interne und externe Empfänger.  
- Festlegung der zeitlichen Gültigkeit von Vorlagen  
- Variablen und Benutzerfotos in Vorlagen  
- Unterschiede DOCX- und HTML-Format  
- Möglichkeiten zur Einbindung eines Disclaimers  
- Gemeinsame Erarbeitung erster Vorlagen auf Basis bestehender Vorlagen und Kundenanforderungen  
- Live-Demonstration auf einem Standard-Client mit einem Testbenutzer und Testpostfächern des Kunden (siehe Voraussetzungen)  
##### 2.8.1.2.3. Dauer  
4 Stunden  
##### 2.8.1.2.4. Voraussetzungen  
- Der Kunde stellt einen Standard-Client mit Outlook und Word zu fVerfügung.  
- Der Bildschirminhalt des Clients muss zur gemeinsamen Arbeit per Beamer projiziert oder auf einem entsprechend großen Monitor dargestellt werden können.  
- Der Kunde stellt einen Testbenutzer zur Verfügung. Dieser Testbenutzer muss auf dem Standard-Client  
	- einmalig Script-Dateien aus dem Internet (github.com) herunterladen dürfen (alternativ kann der Kunde einen BitLocker-verschlüsselten USB-Stick für die Datenübertragung stellen).  
	- unsignierte PowerShell-Scripte im Full Language Mode ausführen dürfen  
	- über ein Mail-Postfach verfügen  
	- Vollzugriff auf diverse Testpostfächer (persönliche Postfächer oder Gruppenpostfächer) haben, die nach Möglichkeit direkt oder indirekt Mitglied in diversen Gruppen oder Verteilerlisten sind. Für den Vollzugriff kann der Benutzer auf die anderen Postfächer entsprechend berechtigt sein, oder Benutzername und Passwort der zusätzlichen Postfächer sind bekannt.  
#### 2.8.1.3. Schulung des Clientmanagements  
##### 2.8.1.3.1. Teilnehmer  
- Kunde: Clientmanagement, optional ein Administrator des Active Directory, optional ein Administrator des File-Servers und/oder SharePoint-Server, optional Unternehmenskommunikation und Marketing, Koordinator des Vorhabens  
- Service-Provider: Mail-Produktmanagement, Mail-Betriebsführung oder Mail-Architektur, ein Vertreter des Client-Teams bei entsprechenden Kunden  
##### 2.8.1.3.2. Inhalt und Ziele  
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
##### 2.8.1.3.3. Dauer  
4 Stunden  
##### 2.8.1.3.4. Voraussetzungen  
- Der Kunde stellt einen Standard-Client mit Outlook und Word zu Verfügung.  
- Der Bildschirminhalt des Clients muss zur gemeinsamen Arbeit per Beamer projiziert oder auf einem entsprechend großen Monitor dargestellt werden können.  
- Der Kunde stellt einen Testbenutzer zur Verfügung. Dieser Testbenutzer muss auf dem Standard-Client  
	- einmalig Script-Dateien aus dem Internet (github.com) herunterladen dürfen (alternativ kann der Kunde einen BitLocker-verschlüsselten USB-Stick für die Datenübertragung stellen).  
	- unsignierte PowerShell-Scripte im Full Language Mode ausführen dürfen
	- über ein Mail-Postfach verfügen  
	- Vollzugriff auf diverse Testpostfächer (persönliche Postfächer oder Gruppenpostfächer) haben, die nach Möglichkeit direkt oder indirekt Mitglied in diversen Gruppen oder Verteilerlisten sind. Für den Vollzugriff kann der Benutzer auf die anderen Postfächer entsprechend berechtigt sein, oder Benutzername und Passwort der zusätzlichen Postfächer sind bekannt.  
- Der Kunde stellt mindestens einen zentralen SMB-File- oder WebDAV-Share für die Ablage der Vorlagen zur Verfügung.  
- Der Kunde stellt einen zentralen SMB-File-Share für die Ablage des Scripts und seiner Komponenten zur Verfügung.  
### 2.8.2. Tests, Pilotbetrieb, Rollout  
Die Planung und Koordination von Tests, Pilotbetrieb und Rollout erfolgt durch den Vorhabens-Verantwortlichen des Kunden.

Die konkrete technische Umsetzung erfolgt durch den Kunden. Falls zusätzlich zu Mail auch der Client durch Service-Provider betreut wird, unterstützt das Client-Team bei der Einbindung des Scripts (Logon-Script, geplante Aufgabe, Desktop-Verknüpfung).

Bei prinzipiellen technischen Problemen unterstützt das Mail-Produktmanagement bei der Ursachenforschung, arbeitet Lösungsvorschläge aus und stellt gegebenenfalls den Kontakt zum Hersteller des Produkts her.

Die Erstellung und Wartung von Vorlagen ist Aufgabe des Kunden

Zur Vorgehensweise bei Anpassungen am Code oder der Veröffentlichung neuer Funktionen siehe Kapitel „Laufender Betrieb“.  
## 2.9. Laufender Betrieb  
### 2.9.1. Erstellen und Warten von Vorlagen  
Das Erstellen und Warten von Vorlagen ist Aufgabe des Kunden.  
Das Mail-Produktmanagement steht für Fragen zu Realisierbarkeit und Auswirkungen beratend zur Verfügung.

### 2.9.2. Erstellen und Warten von Ablage-Shares für Vorlagen und Script-Komponenten  
Das Erstellen und Warten von Ablage-Shares für Vorlagen und Script-Komponenten ist Aufgabe des Kunden.

Das Mail-Produktmanagement steht für Fragen zu Realisierbarkeit und Auswirkungen beratend zur Verfügung.  
### 2.9.3. Setzen und Warten von AD-Attributen  
Das Setzen und Warten von AD-Attributen, die im Zusammenhang mit textuellen Signaturen stehen (z. B. Attribute für Variablen, Benutzerfotos, Gruppenmitgliedschaften), ist Aufgabe des Kunden.

Das Mail-Produktmanagement steht für Fragen zu Realisierbarkeit und Auswirkungen beratend zur Verfügung.  
### 2.9.4. Konfigurationsanpassungen  
Konfigurationsanpassungen, die von den Entwicklern des Scripts explizit vorgesehen sind, werden jederzeit unterstützt.

Das Mail-Produktmanagement steht für Fragen zur Realisierbarkeit und den Auswirkungen gewünschter Anpassungen beratend zur Verfügung.

Die Planung und Koordination von Tests, Pilotbetrieb und Rollout im Zusammenhang mit Konfigurationsanpassungen erfolgt ebenso durch den Kunden wie die konkrete technische Umsetzung.

Falls zusätzlich zu Mail auch der Client durch den Service-Provider betreut wird, unterstützt das Client-Team bei der Einbindung des Scripts (Logon-Script, geplante Aufgabe, Desktop-Verknüpfung).  
### 2.9.5. Probleme und Fragen im laufenden Betrieb  
Bei prinzipiellen technischen Problemen unterstützt das Mail-Produktmanagement bei der Ursachenforschung, arbeitet Lösungsvorschläge aus und stellt gegebenenfalls den Kontakt zum Hersteller des Produkts her.

Für allgemeine Fragen zum Produkt und dessen Anwendungsmöglichkeiten steht ebenfalls das Mail-Produktmanagement zur Verfügung.  
### 2.9.6. Unterstützte Versionen  
Die Versionsnummern des Produkts folgen den Vorgaben des Semantic Versioning und sind daher nach dem Format „Major.Minor.Patch“ aufgebaut.  
- „Major“ wird erhöht, wenn die Kompatibilität zu bisherigen Versionen nicht mehr gegeben ist.  
- „Minor“ wird erhöht, wenn neue Funktionen, die zu bisherigen Versionen kompatibel sind, eingeführt werden.  
- „Patch“ wird erhöht, wenn die Änderungen ausschließlich zu bisherigen Versionen kompatible Fehlerbehebungen umfassen.  
- Zusätzlich sind Bezeichner für Vorveröffentlichungen und Build-Metadaten als Anhänge zum „Major.Minor.Patch“-Format verfügbar, z. B. „-Beta1“.

Vom Service-Provider unterstützte Versionen:  
- Die höchste vom Service-Provider freigegebene Version des Produkts, unabhängig von deren Veröffentlichungsdatum.  
- Die Unterstützung einer freigegeben Version endet automatisch drei monate nach Freigabe einer höheren Version.

Kunden haben nach Freigabe einer neuen Version also drei Monate Zeit, auf diese Version umzusteigen, bevor der Service-Provider-Support für davor freigegebene Versionen erlischt.

Somit ist in einem 3-Monats-Zeitraum nie mehr als eine Aktualisierung notwendig. Dies schützt sowohl Kunden als auch Service-Provider vor groben Fehlern in der Produktentwicklung.  
### 2.9.7. Neue Versionen  
Wenn neue Versionen des Produkts veröffentlicht werden, informiert das Mail-Produktmanagement vom Kunden definierte Ansprechpartner über die mit dieser Version verbundenen Änderungen, mögliche Auswirkungen auf die bestehende Konfiguration und zeigt Aktualisierungsmöglichkeiten auf.

Die Planung und Koordination der Einführung der neuen Version erfolgt durch den Ansprechpartner beim Kunden.

Die konkrete technische Umsetzung erfolgt ebenfalls durch den Kunden. Falls zusätzlich zu Mail auch der Client durch Service-Provider betreut wird, unterstützt das Client-Team bei der Einbindung des Scripts (Logon-Script, geplante Aufgabe, Desktop-Verknüpfung).

Bei prinzipiellen technischen Problemen unterstützt das Mail-Produktmanagement bei der Ursachenforschung, arbeitet Lösungsvorschläge aus und stellt gegebenenfalls den Kontakt zum Hersteller des Produkts her.  
### 2.9.8. Anpassungen am Code des Produkts  
Falls Anpassungen am Code des Produkts gewünscht werden, werden die damit verbundenen Aufwände geschätzt und nach Beauftragung separat verrechnet.

Entsprechend dem Open-Source-Gedanken des Produkts werden die Code-Anpassungen als Verbesserungsvorschlag an die Entwickler des Produkts übermittelt.

Um die Wartbarkeit des Produkts sicherzustellen, kann der Service-Provider nur Code unterstützen, der auch offiziell in das Produkt übernommen wird. Jedem Kunden steht es frei, den Code des Produkts selbst anzupassen, in diesem Fall kann der Service-Provider allerdings keine Unterstützung mehr anbieten. Für Details, siehe „Unterstützte Versionen“.
