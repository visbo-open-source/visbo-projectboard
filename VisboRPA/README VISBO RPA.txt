# VISBO RPA (Robot Process Automation Client)

The VISBO RPA Client is a Windows desktop application designed to automate recurring data-processing tasks for the VISBO Platform. It monitors a designated folder for incoming Excel files and executes automated workflows such as project creation, project updates, and batch data operations.

This file provides the complete technical README for the VISBO RPA application.

------------------------------------------------------------
OVERVIEW
------------------------------------------------------------

The VISBO RPA application automates the processing of Excel-based data imports for VISBO. When a user places an Excel file into a monitored directory, the RPA client detects the file, identifies the required processing task, and communicates with the VISBO REST backend to execute the requested operations.

Typical automated tasks include:

- Uploading VISBO project data, in so-called "Steckbriefs Format" 
- Creating multiple new projects in batch mode
- Moving, staggering planned projects in order o avoid resource Bottlenecks 
- Logging all with success/error information

The RPA client runs as an desktop application and can be configured to work with different directories and API endpoints.

------------------------------------------------------------
FEATURES
------------------------------------------------------------

The key features of the VISBO RPA client include:

- Continuous monitoring of a specified input folder
- Automatic detection of new Excel files
- Intelligent determination of required automation tasks:
  - Project creation
  - Batch job processing
- Execution of the required workflow via VISBO REST API calls
- Full logging of:
  - Processed files
  - Success states
  - Error messages and validation issues
- No administrative permissions required for installation

For functional details, refer to the User Documentation in:
./user-docs/

------------------------------------------------------------
ARCHITECTURE
------------------------------------------------------------

The VISBO RPA client is implemented as a VB.NET Windows Desktop Application targeting .NET Framework 4.6.1 or higher.

High-Level Architecture:

Excel File (Input Folder)
        ↓ (File Watcher)
VISBO RPA Desktop Client (VB.NET / .NET 4.6.1)
        ↓ REST API Communication
VISBO Server REST API
        ↓ Response Logging
RPA Logfile (Success/Error)

Solution Structure (from VisboRPA.sln):

VisboRPA.sln
└── VisboRPA\VisboRPA.vbproj            # Main Windows application:contentReference[oaicite:1]{index=1}

The solution builds in both Debug and Release configurations for Any CPU.

------------------------------------------------------------
REQUIREMENTS
------------------------------------------------------------

End Users:
- Windows 10 or later
- .NET Framework 4.6.1 or newer
- Microsoft Excel 2019 or later, or Microsoft 365 Desktop Edition
- Access to the VISBO REST API (configured in the client)

Developers:
- Microsoft Visual Studio 2019 or later
- .NET Framework 4.6.1 Developer Pack
- Office interop libraries (typically included with Office)
- GitHub access for source code and issue tracking

------------------------------------------------------------
INSTALLATION
------------------------------------------------------------

The VISBO RPA client is distributed via **Microsoft ClickOnce deployment**.

Benefits of ClickOnce:
- Installation without administrator rights
- Automatic update mechanism
- Per-user installation

Installation package and full installation guide for version 7.2.0.1 (1 December 2025) are located in:

./admin-doc/

This directory contains:
- A ZIP archive with the ClickOnce release
- Installation instructions
- All required binaries

------------------------------------------------------------
BUILDING FROM SOURCE
------------------------------------------------------------

To build the VISBO RPA application:

1. Clone the repository:

2. Open the solution:

VisboRPA.sln

3. Select your build configuration:
- Debug
- Release

4. Build the application:
There are libraries needed from the visbo-open-source/visbo-projectboard/projectboard 

Build → Build Solution (Ctrl + Shift + B)

Notes:
- The project targets Any CPU.
- Configuration settings for API endpoints are defined in the App.config file.
- ClickOnce settings are defined in the Visual Studio project properties.

------------------------------------------------------------
USAGE GUIDE
------------------------------------------------------------

End-user documentation is included under:

./user-docs/

This guide provides instructions on:

- Preparing Excel files for automation
- Understanding different automation modes


------------------------------------------------------------
FOLDER STRUCTURE
------------------------------------------------------------

/VisboRPA
    Source code of the RPA Windows application

/admin-doc
    ClickOnce deployment files
    Installation script and documentation
    Version 7.2.0.1 delivery package

/user-docs
    End user documentation (Excel templates, walkthroughs, PDF guides)

README.txt
LICENSE.md
CLA.md
COMMERCIAL-LICENSE.md

------------------------------------------------------------
CONTRIBUTING
------------------------------------------------------------

The VISBO RPA Client is open for community contributions.

Contribution Process:
1. Fork this repository
2. Create a new feature branch
3. Commit your changes
4. Open a Pull Request against the main branch
5. Sign the Contributor License Agreement (CLA) when prompted

For questions and contribution inquiries:
open.source@visbo.de

------------------------------------------------------------
LICENSE
------------------------------------------------------------

This project is licensed under:

- GNU Affero General Public License v3.0 (AGPLv3)
- Commons Clause restriction
- VISBO Dual Licensing Framework

See:
LICENSE.md
CLA.md
COMMERCIAL-LICENSE.md

The software is provided AS IS, without warranties or guarantees.

------------------------------------------------------------
END OF README
------------------------------------------------------------
