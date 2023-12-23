# VBS_File_Generator

## Overview

This VBScript generates random files with specified sizes and names, providing a simple way to test or demonstrate file-related operations.

## Features

- Generates a unique copy of the script with a random name.
- Creates random data files named "cp_i_" in the designated folder.
- Supports customization of the number of copies and desired file size.

## Prerequisites

- Windows operating system (VBScript is typically used on Windows).
- Windows Script Host (WSH) enabled.

## Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/SSL-ACTX/VBS_File_Generator.git
   ```

2. Navigate to the repository folder:

   ```bash
   cd VBS_File_Generator
   ```

## Usage

1. Double-click on the `generate_files.vbs` script.
2. Review the generated files in the specified folder.

## Configuration

- You can customize the number of copies and desired file size in the script.

## Warning and Ethical Considerations

This script generates random data files and may create a large number of files in a designated folder. Please be aware of the following:

- **Ethical Use:** Ensure that you have the right to generate and manipulate files in the specified folder. Using this script on systems or directories without proper authorization may be unethical or even illegal.

- **File System Impact:** Generating a large number of files can impact file systems. Be cautious not to overload file storage systems or interfere with critical operations.

- **Use Responsibly:** Be mindful of the impact of generating large files or quantities of files on your storage and system resources.

## Example

```vbscript
' Modify these variables as needed
numCopies = 10
String(1024*1024,"A")
```

