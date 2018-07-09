Case Migration Script
=====================

![Last tested in Nuix 7.4](https://img.shields.io/badge/Nuix-7.4-green.svg)

View the GitHub project [here](https://github.com/Nuix/Case-Migration-Script) or download the latest release [here](https://github.com/Nuix/Case-Migration-Script/releases).

# Overview

**Written By:** Jason Wells

Migrates a list of Nuix cases to the latest case version.

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

## Input

When ran, the script will prompt you to make some choices:
- You will be prompted whether or not to make backups
	- If you decide to make backups, you will be prompted for the directory to store the backups
- You will be prompted about how you would like to specify the cases to migrate
	- If you choose "Locate Cases in Directory", you will be prompted to select a root search directory, all cases found will be used
	- If you choose "Provide Text File with Case Paths", you will be prompted to select a text file with one case directory per line

## Output

While running, this script will produce 2 files in the same directory as the script.

- `{TIMESTAMP}_Log.txt` : This will containg a log of all output that is shown in the console
- `{TIMESTAMP}_Report.csv` : This will contain a listing of cases which succeeded and failed, with error messages for the failures

Output written to log file is also written to script console.

# License

```
Copyright 2018 Nuix

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```