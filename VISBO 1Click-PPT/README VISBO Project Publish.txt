# VISBO Project Publish – Microsoft Project Add-In

The **VISBO Project Publish Add-In** is a native Microsoft Project extension that enables users to export selected project plan elements – such as phases and milestones – into the VISBO data model. Each export operation creates a timestamped record in the VISBO backend.  
This Add-In is part of the VISBO Open Source ecosystem and works in combination with the VISBO PowerPoint Add-In, which can generate presentation-ready reports from the synchronized VISBO project data.


## Overview

The VISBO Project Publish Add-In adds a simple but powerful capability to Microsoft Project:

- It allows project managers to **publish** the currently opened MS Project plan directly into the VISBO platform.
- Only one action is required: **Publish**.
- The VISBO Add-IN will create or update a VISBO Project using the **same name** as the MS Project plan.
- Every publish operation is written into the VISBO data store with a server-side timestamp.

This Add-In is implemented as a **.NET Framework 4.6.1+ VSTO Add-In** for Microsoft Project and packaged using Microsoft ClickOnce for easy deployment.


## Features

- **Single-click publishing** of the current MS Project plan to VISBO  
- **Automatic project name alignment** between MS Project and VISBO  
- **Timestamped Versions for historic Analysis 
- **Zero-admin installation** via ClickOnce  

## Architecture

The solution contains two main components:

VISBO Project Publish.sln
├── VISBO Project Publish.vbproj # Main VSTO Add-In project
└── VISBOProjectPublishSetup.vdproj # Setup project for ClickOnce installer

:contentReference[oaicite:1]{index=1}

### Technology Overview

- Developed in **VB.NET**
- Built using **Visual Studio 2019 or later**
- Uses the **Microsoft Office VSTO framework**
- Communicates with VISBO through the REST backend 
- Deployed via **ClickOnce** (no administrator rights required)

---

## Requirements

### End User Requirements
- **Microsoft Project 2019 or newer**   or **Microsoft 365 (Desktop version)**  
- **.NET Framework 4.6.1 or later**
- Access to a running **VISBO REST backend** (if publishing into a VISBO instance)

### Developer Requirements
- **Windows 10 or Windows 11**
- **Visual Studio 2019 or later**
- **Office Developer Tools**
- **VSTO development runtime**
- **.NET Framework 4.6.1 developer pack**
- Git + GitHub access


## Installation

> The Add-In is deployed using Microsoft **ClickOnce**, allowing end users to install it **without Administrator privileges**.

Installation packages and scripts are available in: ./admin-doc


This folder includes:

- Installation ZIP archive  
- Setup scripts  
- Build artifacts for version **7.2.0.1 (as of 1 Dec 2025)**

For installation instructions, please refer to:

**Installation Guide → located in `./admin-doc**


## Building From Source

Recommended development environment:

- **Visual Studio 2019 or later**
- .NET Framework 4.6.1 targeting pack
- Office Developer Tools installed

### Steps

1. Fork / Clone the repository:
2. in Visual Studio : open the solution file: ProjectPublish.sln 
3. Select build configuration:beware that libraries from visbo-open-source/visbo-projet-board/projectboard are necessary 
4  Build Notes:
- The Add-In targets Any CPU.
- Ensure that Office Developer Tools for Visual Studio are installed.
- ClickOnce publishing settings are defined in the project configuration.



## INSTALLATION

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

## USAGE GUIDE

End-user documentation is included in:

./user-docs/


## FOLDER STRUCTURE

/VISBO 1ClickPPT 
    VB.NET source code for the MS PRoject Add-In

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

## CONTRIBUTING

The VISBO PowerPoint Add-In is open for community contributions.

Contribution Process:
1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Submit a Pull Request to the main branch
5. Sign the Contributor License Agreement (CLA) 

Questions and contribution inquiries:
open.source@visbo.de

## LICENSE

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
