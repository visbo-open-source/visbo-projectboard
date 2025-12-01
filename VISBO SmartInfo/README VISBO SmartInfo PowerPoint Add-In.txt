# VISBO SmartInfo PowerPoint Add-In (Native Client)

The VISBO SmartInfo Add-In is a native Microsoft PowerPoint extension that generates project reports from predefined PowerPoint templates containig VISBO reporting components from the VISBO Report cpmonents.pptx. It accesses VISBO Projects in the OnPRemise / PRivate Cloud or VISBO SaaS enables project managers to create and update powerpoint Reports by th press of  button.

This file provides the full technical README for the VISBO PowerPoint Add-In in TXT format for distribution.

------------------------------------------------------------
OVERVIEW
------------------------------------------------------------

The VISBO SmartInfo Add-In allows VISBO users to  populate PowerPoint report templates with up-to-date project Project information. It retrieves project data from VISBO, inserts the values into designated placeholders of the selected template, and produces a fully formatted project status report.

The Add-In is launched directly from within Microsoft PowerPoint and uses Microsoft ClickOnce for installation and updates.

------------------------------------------------------------
FEATURES
------------------------------------------------------------

The PowerPoint Add-In supports the following functionality:

- Open VISBO PowerPoint report templates
- Select Project from VISBO 
- Create Report 

or 
- Open an already existing VISBO PowerPoint Project report 
- Update Report 


A complete user guide is provided under:
./user-docs/

------------------------------------------------------------
ARCHITECTURE
------------------------------------------------------------

The VISBO SmartInfo Add-In is implemented as a VSTO (Visual Studio Tools for Office) Add-In using VB.NET, based on .NET Framework 4.6.1 or newer.

High-Level Architecture:

VISBO SmartInfo Add-In (VSTO / VB.NET)
     ↓  REST API communication (if configured)
VISBO Server REST API

Solution Structure (from VISBO SmartInfo.sln):

VISBO SmartInfo.sln
└── VISBO SmartInfo\VISBO SmartInfo.vbproj    

The solution builds for Any CPU in Debug and Release configurations.

------------------------------------------------------------
REQUIREMENTS
------------------------------------------------------------

End Users:
- Microsoft PowerPoint 2019 or later
- Alternatively: Microsoft 365 Desktop Edition
- Windows 10 or later
- .NET Framework 4.6.1 or later
- Network access to VISBO Server (if using live project data)

Developers:
- Microsoft Visual Studio 2019 or newer
- Office Developer Tools for Visual Studio
- .NET Framework 4.6.1 Targeting Pack
- GitHub access for VISBO repositories

------------------------------------------------------------
INSTALLATION
------------------------------------------------------------

The Add-In is deployed using **Microsoft ClickOnce**.

Installation characteristics:
- No administrator rights required
- Automatic update support
- Per-user installation on Windows

The installation package, including version 7.2.0.1 (dated 1 December 2025), is located in:

./admin-doc/

This folder contains:
- A ZIP archive with the ClickOnce install package 
- A installation guide
- All required deployment files

------------------------------------------------------------
BUILDING FROM SOURCE
------------------------------------------------------------

To build the VISBO SmartInfo PowerPoint Add-In locally:

1. Clone the repository:

2. Open the Visual Studio solution:

VISBO SmartInfo.sln

3. Select build configuration:
- beware that libraries from visbo-open-surce/visbo-projet-board/projectboard are ecessary 

4. Build the project:

Build → Build Solution 

Notes:
- The Add-In targets Any CPU.
- Ensure that Office Developer Tools for Visual Studio are installed.
- ClickOnce publishing settings are defined in the project configuration.

------------------------------------------------------------
USAGE GUIDE
------------------------------------------------------------

End-user documentation is included in:

./user-docs/

The guide covers:
- Selecting and loading PowerPoint templates
- Generating full project reports
- Refreshing report data from VISBO
- Troubleshooting common issues

------------------------------------------------------------
FOLDER STRUCTURE
------------------------------------------------------------

/VISBO SmartInfo
    VB.NET source code for the PowerPoint Add-In

/admin-doc
    ClickOnce deployment package
    Installation guide
    Version 7.2.0.1 installation ZIP

/user-docs
    End-user documentation for the PowerPoint Add-In

README.txt
LICENSE.md
CLA.md
COMMERCIAL-LICENSE.md

------------------------------------------------------------
CONTRIBUTING
------------------------------------------------------------

The VISBO PowerPoint Add-In is open for community contributions.

Contribution Process:
1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Submit a Pull Request to the main branch
5. Sign the Contributor License Agreement (CLA) 

Questions and contribution inquiries:
open.source@visbo.de

------------------------------------------------------------
LICENSE
------------------------------------------------------------

This project is licensed under:

- GNU Affero General Public License v3.0 (AGPLv3)
- Commons Clause restriction
- VISBO Dual Licensing Framework

Please review:
LICENSE.md
CLA.md
COMMERCIAL-LICENSE.md

The software is provided AS IS, without warranties or guarantees.

------------------------------------------------------------
END OF README
------------------------------------------------------------
