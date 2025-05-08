# Working Hours Processing

A simple toolset for processing and analyzing working–hours data. Includes a standalone Windows executable for ease of use.

---

## Table of Contents

1. [Overview](#overview)  
2. [Features](#features)  
3. [Requirements](#requirements)  
4. [Installation (Windows)](#installation-windows)  
5. [Usage](#usage)  
6. [Contributing](#contributing)  
7. [License](#license)  

---

## Overview

This repository provides:

- A GUI application to import, validate, and process timesheet data  
- Export options in CSV or Excel format  
- Summary reports (total hours, overtime, anomalies)  

---

## Features

- Drag-and-drop support for CSV/Excel files  
- Configurable shift definitions  
- Built-in data validation (missing entries, overlap detection)  
- Windows standalone executable (`gui_app.exe`)

---

## Requirements

- **Windows 10 or later**  
- None of the following need to be installed on the user’s machine (all dependencies bundled in the `.exe`):
  - Python  
  - Third-party libraries  

---

## Installation (Windows)

1. Download the latest release from the **Releases** page:  
   - [`gui_app.exe`](https://github.com/your-username/working-hours-processing/releases)  
2. Place `gui_app.exe` in a folder of your choice.  

---

## Usage

1. Double-click **gui_app.exe** to launch.  
2. When prompted, select your input file (CSV or Excel).  
3. Configure any options (date format, shift rules).  
4. Click **Process**.  
5. Find your output files in the same directory as the input.

---

## License

This project is licensed under the [MIT License](LICENSE).  
