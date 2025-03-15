# Windows-Acquisition-Tool

## Overview
This Python script creates a Virtual Hard Disk (VHD) container on Windows systems that stores files or directories while preserving metadata and folder structures. It also records system information and maintains logs of all operations.

## Requirements

### System Requirements
- Windows operating system
- Administrator privileges
- Python 3.6 or higher
- At least 100MB of free disk space (more depending on what you're storing)

### Python Dependencies
The script uses only standard library modules:
- os, sys, time, platform, subprocess, logging, shutil, datetime, winreg, ctypes

No external packages need to be installed.

## Installation
1. Save the Python script as `Targeted_Windows_AcquisitionTool.py`
2. Ensure you have Python installed on your system

## Usage Instructions

### Running the Script
1. Right-click on Command Prompt or PowerShell and select "Run as administrator"
2. Navigate to the directory containing the script
3. Run the script:
   ```
   python Targeted_Windows_AcquisitionTool.py
   ```

### Script Workflow
1. The script will check if it has administrator privileges
2. System information will be gathered and saved to `system_info.txt`
3. You'll be prompted to enter the path to the file or directory you want to store
4. The script will calculate the required VHD size
5. You'll be prompted to enter where to save the VHD file
6. The script will create, mount, and populate the VHD with your data
7. The VHD will be detached when complete
8. A log file (`vhd_operation.log`) will be created in the same directory as the script

### Accessing Your Data
- To access the VHD after creation, right-click on the VHD file in File Explorer and select "Mount"
- Windows will mount the VHD and assign it a drive letter
- When finished, right-click on the drive in File Explorer and select "Eject"

## Customizable Parameters

You may want to modify the following lines in the script based on your needs:

### VHD Settings
- Line 138: `create vdisk file="{vhd_path}" maximum={size_mb} type=expandable`
  - Change `type=expandable` to `type=fixed` if you prefer a fixed-size VHD (better performance but uses full space immediately)

- Line 141: `format fs=ntfs quick label="VHD Storage"`
  - You can change the label "VHD Storage" to any name you prefer

- Line 142: `assign letter=V`
  - Change `letter=V` if you want to use a different drive letter

### Size Calculation
- Line 226: `return max(100, size_mb)`
  - The minimum VHD size is set to 100MB; increase this if needed

- Line 219: `size_mb = int((total_size / (1024 * 1024)) * 1.2)`
  - The 1.2 multiplier adds a 20% buffer; adjust as needed

### Log Settings
- Line 16: `filename='vhd_operation.log',`
  - Change the log filename if desired

- Line 17: `level=logging.INFO,`
  - Change to `logging.DEBUG` for more detailed logs

## Troubleshooting

### Common Issues
1. **"This script requires administrator privileges"**
   - Make sure to run Command Prompt or PowerShell as administrator

2. **Diskpart errors**
   - Check that you have sufficient permissions for disk operations
   - Verify the path for the VHD is valid and accessible

3. **"Path not found" errors**
   - Ensure you're entering correct paths for the source file/directory
   - Make sure to use full paths or proper relative paths

4. **Insufficient space errors**
   - Free up disk space or specify a different location for the VHD

### Log Files
If you encounter issues, check the `vhd_operation.log` file for detailed error information.

## Security Notes
- This script requires administrator privileges because it uses diskpart to create and manipulate VHD files
- The VHD is not encrypted by default; consider encrypting it with BitLocker for sensitive data
- The script attempts to preserve file permissions, but there may be limitations depending on the specific attributes
