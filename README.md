# Barracuda Engineering Limited - Tools Repository

This repository contains engineering tools and utilities developed by Barracuda Engineering Limited.

## Tools

### 1. edit_floppy_usb.vbs

A VBScript utility for managing and editing Gotek floppy disk images on Windows systems.

#### Description

`edit_floppy_usb.vbs` provides a streamlined workflow for editing 1.44MB floppy disk images (typically used with Gotek floppy emulators). The script automates the entire process of mounting, editing, and saving floppy images back to a USB drive.

#### Features

- **Auto-detection**: Automatically detects USB drives containing the target floppy image
- **Image creation**: Creates and formats new 1.44MB FAT floppy images if none exists
- **Virtual mounting**: Mounts images using ImDisk Virtual Disk Driver
- **User-friendly editing**: Opens Windows Explorer for intuitive file management
- **Automatic sync**: Copies modified images back to USB drive after editing
- **Safety features**: 
  - Prevents multiple instances from running simultaneously
  - Includes retry logic for unmounting
  - Maintains detailed log files for troubleshooting
  - Handles read-only attributes automatically

#### Requirements

- **Operating System**: Windows (tested on Windows 7, 8, 10, 11)
- **ImDisk Virtual Disk Driver**: Must be installed and available in PATH
  - Download from: http://www.ltr-data.se/opencode.html/#ImDisk
  - Or search for "ImDisk Virtual Disk Driver"
- **USB Drive**: Removable drive for storing floppy images

#### Configuration

Default settings (can be modified in the script):

```vbscript
imgName     = "DSKA0000.IMG"    ' Image filename to look for/create
mountLetter = "A:"              ' Virtual drive letter when mounted
useGUI      = True              ' True for GUI dialogs, False for console
```

#### Usage

1. **Insert USB drive** containing your Gotek floppy image (e.g., `DSKA0000.IMG`)
2. **Run the script** by double-clicking `edit_floppy_usb.vbs`
3. The script will:
   - Detect the USB drive automatically
   - Create a working copy in the `work` subdirectory
   - Mount the image as drive A: (or configured letter)
   - Open Windows Explorer to the mounted drive
4. **Edit files** as needed in the Explorer window
5. **Close Explorer** when finished
6. The script automatically:
   - Unmounts the virtual drive
   - Copies the modified image back to the USB drive
   - Displays a confirmation message

#### First-Time Setup

If no image exists on the USB drive:

1. Insert a USB drive
2. Run the script
3. The script will automatically:
   - Create a new 1.44MB blank image file
   - Format it as FAT filesystem
   - Mount it for editing
   - Open Explorer for you to add files

#### Log Files

The script maintains a log file (`edit_floppy_usb.log`) in the same directory as the script, recording:
- Timestamps of operations
- Error messages
- Warning conditions

This is useful for troubleshooting if issues arise.

#### Working Directory

The script creates a `work` subdirectory in its installation folder to store temporary working copies of floppy images. This ensures the original USB image is only modified after successful editing.

#### Troubleshooting

**Problem**: Script reports "ImDisk Virtual Disk Driver is not installed"
- **Solution**: Download and install ImDisk from the link above, ensure it's in your system PATH

**Problem**: Script says "Another instance is already running"
- **Solution**: Wait for the previous instance to complete, or use Task Manager to end `wscript.exe` processes

**Problem**: Cannot unmount drive or "drive still in use" error
- **Solution**: Close all Explorer windows showing the virtual drive, close any programs accessing files on the drive

**Problem**: Image fails to copy back to USB
- **Solution**: Check that the USB drive is not write-protected, has sufficient space, and the file is not read-only

#### Technical Details

- **Image Format**: 1.44MB (1,474,560 bytes) FAT filesystem
- **Drive Type Detection**: Uses WMI to identify removable drives (DriveType = 1)
- **Virtual Drive**: Created using ImDisk with removable flag (`-o rem`)
- **Process Safety**: Checks for running instances via WMI query
- **Explorer Monitoring**: Uses Shell.Application COM object to track Explorer windows

#### License

Copyright © Barracuda Engineering Limited. All rights reserved.

---

## Repository Information

**Repository URL**: https://github.com/davidnoz123/barracuda_engineering_limited

## Contributing

For questions, issues, or contributions, please contact Barracuda Engineering Limited.

## Support

For technical support or inquiries, please refer to the repository issues page or contact the development team.

---

*Last updated: November 22, 2025*
