# README VISBO Excel Add-In (Native Client)

The VISBO Excel Add-In is a native Microsoft Office extension that enables users to edit and update key VISBO project data directly from within the VISBO Web-UI invoking the SPE (Simple Project Edit) Excel Add-In.

This file contains the complete README in TXT form for distribution.

------------------------------------------------------------
OVERVIEW
------------------------------------------------------------

The VISBO Excel Add-In provides a convenient and efficient interface for project managers to update core project information without leaving the VISBO Web-Userinterface. The Add-In is invoked via the VISBO Connect Utility directly from the VISBO Web UI and relies on Microsoft’s ClickOnce deployment technology for installation and updating.

------------------------------------------------------------
FEATURES
------------------------------------------------------------

The add-in supports the editing and updating of essential project attributes, including:

- Project milestones and schedule data
- Deliverables and completion progress
- Traffic light indicators (status “red/yellow/green”)
- Resource demands, cost needs
- Secure upload of updated project data to the VISBO backend
- Ability to open projects directly from the VISBO web interface

For detailed functional descriptions, refer to:
./user-docs/

------------------------------------------------------------
ARCHITECTURE
------------------------------------------------------------

High-Level Architecture:

VISBO Web Client
    ↓ (via Connect Utility)
VISBO Excel Add-In (VSTO / VB.NET)
    ↓ REST API communication
VISBO Server REST API

Solution Structure (from VISBO SPE.sln):

VISBO SPE.sln
└── VISBO SPE.vbproj              # Main Excel Add-In project

Note: The old setup project (VISBO ProjectEditSetup.vdproj) is deprecated.

------------------------------------------------------------
REQUIREMENTS
------------------------------------------------------------

End Users:
- Microsoft Excel 2019 or later, or Microsoft 365 Desktop
- Windows 10 or later
- .NET Framework 4.6.1 or newer
- Network connectivity for VISBO REST backend usage

Developers:
- Visual Studio 2019 or later
- Office Developer Tools for Visual Studio
- .NET Framework 4.6.1 Targeting Pack
- GitHub Access
- the Simple Project Edit does require libraries from visbo-open-source/visbo-projectboard/projectboard 

------------------------------------------------------------
INSTALLATION
------------------------------------------------------------

The Add-In is deployed using Microsoft ClickOnce.

Install-Packages and installation instructions are available under:

./admin-doc/

ClickOnce installation does not require administrator privileges.

------------------------------------------------------------
BUILDING FROM SOURCE
------------------------------------------------------------

1. Clone the repository or your fork:


2. Open the solution:

VISBO SPE.sln

3. Select configuration:
- Debug or
- Release

4. Build using Visual Studio:
Build → Build Solution (Ctrl + Shift + B)

------------------------------------------------------------
USAGE GUIDE
------------------------------------------------------------

A complete user guide is included under:

./user-docs/

This includes:
- Editing project data
- Uploading updates

------------------------------------------------------------
FOLDER STRUCTURE
------------------------------------------------------------

/VISBO SPE
    Source code of the Excel Add-In

/admin-doc
    ClickOnce deployment package
    Installation guide

/user-docs
    End user documentation

README.txt
LICENSE.md
CLA.md
COMMERCIAL-LICENSE.md

------------------------------------------------------------
CONTRIBUTING
------------------------------------------------------------

Contributions are welcome.

Contribution Flow:
1. Fork the repository
2. Create a feature branch
3. Commit changes
4. Submit a Pull Request to main
5. Sign the CLA 

Contact for questions:
open.source@visbo.de

------------------------------------------------------------
LICENSE
------------------------------------------------------------

This project is released under:

- GNU Affero General Public License v3.0 (AGPLv3)
- Commons Clause restriction
- VISBO’s Dual Licensing Framework

See:
LICENSE.md
CLA.md
COMMERCIAL-LICENSE.md

The software is provided AS IS, with no warranties or guarantees.

------------------------------------------------------------
END OF README
------------------------------------------------------------
