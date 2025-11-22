' edit_floppy.vbs
' Single-script workflow for Gotek floppy image editing
' - Creates IMG if missing
' - Formats as 1.44MB FAT
' - Mounts via ImDisk
' - Opens Explorer and waits for it to close
' - Unmounts and copies back to USB
'
' Requires: ImDisk installed and available in PATH.

Option Explicit

' === GLOBAL VARIABLES ===
Dim fso, shell, SCRIPT_NAME, useGUI
Dim usbDrive, workDir, imgName, mountLetter, srcImg, workImg, scriptPath, logFile

' === INITIALIZATION ===
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Script name for messages
SCRIPT_NAME = WScript.ScriptName

' Set to True for GUI message boxes, False for console output
useGUI = True

' Image configuration
imgName     = "DSKA0000.IMG"
mountLetter = "A:"   ' Virtual floppy letter when mounted via ImDisk

' Wait times (milliseconds)
Dim wait_milliseconds
wait_milliseconds = 1000

' Log file in script's directory
Dim scriptDir
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
logFile = scriptDir & "\" & fso.GetBaseName(SCRIPT_NAME) & ".log"

' === Check if script is already running ===
Function IsScriptAlreadyRunning()
    On Error Resume Next
    Dim wmi, processes, process, scriptName
    scriptName = WScript.ScriptName
    
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'wscript.exe' OR Name = 'cscript.exe'")
    
    Dim count
    count = 0
    
    For Each process In processes
        If InStr(1, process.CommandLine, scriptName, vbTextCompare) > 0 Then
            count = count + 1
            If count > 1 Then
                IsScriptAlreadyRunning = True
                Exit Function
            End If
        End If
    Next
    
    IsScriptAlreadyRunning = False
    On Error GoTo 0
End Function

' Check for already running instance
If IsScriptAlreadyRunning() Then
    MsgBox "Another instance of " & WScript.ScriptName & " is already running." & vbCrLf & vbCrLf & _
           "Please wait for it to complete or close it before running again.", _
           vbExclamation, WScript.ScriptName
    WScript.Quit 1
End If

' === Find USB drive containing the image file ===
Function FindImageDrive(fileName)
    Dim drive, drives
    Set drives = fso.Drives
    
    For Each drive In drives
        If drive.IsReady And drive.DriveType = 1 Then ' DriveType 1 = Removable
            If fso.FileExists(drive.DriveLetter & ":\" & fileName) Then
                FindImageDrive = drive.DriveLetter & ":"
                Exit Function
            End If
        End If
    Next
    
    FindImageDrive = "" ' Not found
End Function

' === Find first available USB drive ===
Function FindFirstUSBDrive()
    Dim drive, drives
    Set drives = fso.Drives
    
    For Each drive In drives
        If drive.IsReady And drive.DriveType = 1 Then ' DriveType 1 = Removable
            FindFirstUSBDrive = drive.DriveLetter & ":"
            Exit Function
        End If
    Next
    
    FindFirstUSBDrive = "" ' No USB found
End Function

' Scan for USB drive with the image file
usbDrive = FindImageDrive(imgName)

' If not found, try to find any USB drive and we'll create the image on it
If usbDrive = "" Then
    usbDrive = FindFirstUSBDrive()
    
    If usbDrive = "" Then
        ShowMessage "No USB drive detected." & vbCrLf & vbCrLf & _
                    "Please insert a USB drive and try again.", _
                    vbCritical, SCRIPT_NAME
        WScript.Quit 1
    End If
    
    ' Inform user we'll create a new image on this USB
    ShowMessage "No " & imgName & " found." & vbCrLf & vbCrLf & _
                "A new image will be created on USB drive " & usbDrive, _
                vbInformation, SCRIPT_NAME
End If

' Get the directory where this script is located
scriptPath = WScript.ScriptFullName
workDir = Left(scriptPath, InStrRev(scriptPath, "\") - 1) & "\work"

' === Check if ImDisk is installed ===
Function IsImDiskInstalled()
    On Error Resume Next
    Dim result
    result = shell.Run("imdisk.exe", 0, True)
    ' If imdisk is found, it will return an error code (usage error since no params)
    ' If not found, Err.Number will be set
    IsImDiskInstalled = (Err.Number = 0)
    On Error GoTo 0
End Function

If Not IsImDiskInstalled() Then
    ShowMessage "ImDisk Virtual Disk Driver is not installed or not in PATH." & vbCrLf & vbCrLf & _
                "This script requires ImDisk to mount floppy images." & vbCrLf & vbCrLf & _
                "Download from: http://www.ltr-data.se/opencode.html/#ImDisk" & vbCrLf & _
                "Or search for 'ImDisk Virtual Disk Driver'", _
                vbCritical, SCRIPT_NAME
    WScript.Quit 1
End If

' === MESSAGE/PROMPT FUNCTIONS ===
srcImg  = usbDrive & "\" & imgName
workImg = workDir  & "\" & imgName

' === Log Function ===
' Writes a timestamped message to the log file
Sub PrintLog(msg)
    On Error Resume Next
    Dim logStream, timestamp, cleanMsg
    
    ' Get current timestamp
    timestamp = Now()
    
    ' Replace vbCrLf with \n for cleaner log lines
    cleanMsg = Replace(msg, vbCrLf, "\n")
    
    ' Open log file for appending (create if doesn't exist)
    Set logStream = fso.OpenTextFile(logFile, 8, True)
    
    ' Write timestamped message
    logStream.WriteLine timestamp & " - " & cleanMsg
    
    ' Close the file
    logStream.Close
    Set logStream = Nothing
    On Error GoTo 0
End Sub

' === Message Output Routine ===
' Shows message via MsgBox (GUI) or WScript.Echo (console) based on useGUI setting
' msgType: vbInformation, vbCritical, vbExclamation (ignored for console output)
Sub ShowMessage(msg, msgType, title)
    If msgType = vbCritical Then
        PrintLog "ERROR:" & msg
    End If
    If useGUI Then
        MsgBox msg, msgType, title
    Else
        WScript.Echo title & ": " & msg
    End If
End Sub

' === Prompt for user input in console mode ===
Sub WaitForUser(msg)
    If useGUI Then
        MsgBox msg, vbInformation, SCRIPT_NAME
    Else
        WScript.Echo msg
        WScript.Echo "Press Enter to continue..."
        WScript.StdIn.ReadLine
    End If
End Sub

' === Safe cleanup routine - unmounts drive and logs any errors ===
Sub SafeUnmount(driveLetter)
    On Error Resume Next
    Dim cmd, rc
    cmd = "imdisk -d -m " & driveLetter
    rc = shell.Run(cmd, 0, True)
    If Err.Number <> 0 Or rc <> 0 Then
        PrintLog "WARNING: Unmount cleanup failed for " & driveLetter & " - Error: " & Err.Description & " (rc=" & rc & ")"
    End If
    On Error GoTo 0
End Sub

' === Check if Explorer windows are open for a specific drive ===
Function HasExplorerWindows(driveLetter)
    On Error Resume Next
    Dim shellApp, window, locationURL, searchPattern, foundWindow
    Set shellApp = CreateObject("Shell.Application")
    
    foundWindow = False
    searchPattern = UCase(driveLetter) & "/"
    
    For Each window In shellApp.Windows
        locationURL = ""
        locationURL = window.LocationURL
        If Len(locationURL) > 0 Then
            If InStr(1, UCase(locationURL), searchPattern, vbTextCompare) > 0 Then
                foundWindow = True
                Exit For
            End If
        End If
    Next
    
    Set shellApp = Nothing
    HasExplorerWindows = foundWindow
    On Error GoTo 0
End Function

' === Ensure work dir exists ===
If Not fso.FolderExists(workDir) Then
    On Error Resume Next
    fso.CreateFolder(workDir)
    If Err.Number <> 0 Then
        ShowMessage "Failed to create work folder: " & workDir & vbCrLf & _
                    "Error: " & Err.Description, vbCritical, SCRIPT_NAME
        WScript.Quit 1
    End If
    On Error GoTo 0
End If

' === Unmount any existing mount on the target drive letter ===
' This prevents errors if a previous run didn't complete properly
SafeUnmount mountLetter
' Ignore errors - if nothing was mounted, that's fine

' Wait a moment for the unmount to complete
WScript.Sleep wait_milliseconds

' Delete the work image if it exists to ensure we have a clean copy
If fso.FileExists(workImg) Then
    On Error Resume Next
    fso.DeleteFile workImg, True
    If Err.Number <> 0 Then
        ShowMessage "Cannot delete old work file (may still be in use):" & vbCrLf & _
                    workImg & vbCrLf & vbCrLf & _
                    "Error: " & Err.Description & vbCrLf & vbCrLf & _
                    "Try closing all Explorer windows and run the script again.", _
                    vbCritical, SCRIPT_NAME
        WScript.Quit 1
    End If
    On Error GoTo 0
End If

' === If image does not exist on USB, create and format a new one ===
If Not fso.FileExists(srcImg) Then
    Dim cmd, rc

    ' Create blank 1.44MB image in workDir 
    cmd = "cmd /c fsutil file createnew " & Chr(34) & workImg & Chr(34) & " 1474560"
    rc = shell.Run(cmd, 0, True)
    If rc <> 0 Then
        ShowMessage "fsutil failed creating image:" & vbCrLf & cmd, vbCritical, SCRIPT_NAME
        WScript.Quit 1
    End If

    ' Wait for filesystem to commit the new file
    WScript.Sleep wait_milliseconds

    ' Mount it via ImDisk
    cmd = "imdisk -a -t file -f " & Chr(34) & workImg & Chr(34) & " -m " & mountLetter & " -o rem"
    rc = shell.Run(cmd, 0, True)
    If rc <> 0 Then
        ShowMessage "ImDisk failed mounting new image (exit code: " & rc & "):" & vbCrLf & cmd, vbCritical, SCRIPT_NAME
        WScript.Quit 1
    End If

    ' Format it as FAT (non-interactive: echo Y | format ...)
    cmd = "cmd /c echo Y|format " & mountLetter & " /FS:FAT /V:FLOPPY /Q"
    rc = shell.Run(cmd, 0, True)
    If rc <> 0 Then
        ShowMessage "format failed on " & mountLetter & ":" & vbCrLf & cmd, vbCritical, SCRIPT_NAME
        ' Try to unmount before exiting
        SafeUnmount mountLetter
        WScript.Quit 1
    End If
    
    WScript.Sleep wait_milliseconds

Else
    ' If image exists on USB, copy it into work dir
    On Error Resume Next
    fso.CopyFile srcImg, workImg, True
    If Err.Number <> 0 Then
        ShowMessage "Failed copying image from USB to work folder:" & vbCrLf & _
                    "From: " & srcImg & vbCrLf & "To:   " & workImg & vbCrLf & _
                    "Error: " & Err.Description & vbCrLf & vbCrLf & _
                    "Check if the file is read-only or in use.", vbCritical, SCRIPT_NAME
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Remove read-only attribute if present
    If Not fso.FileExists(workImg) Then
      ShowMessage "Not fso.FileExists(workImg):" & vbCrLf & workImg, vbCritical, SCRIPT_NAME
      WScript.Quit 1      
    Else:
        Dim workFile
        Set workFile = fso.GetFile(workImg)
        If workFile.Attributes And 1 Then ' 1 = ReadOnly
            workFile.Attributes = workFile.Attributes And Not 1
        End If
    End If
    
      
    ' === Mount image from work dir for editing ===
    Dim cmdMount, rcMount
    cmdMount = "imdisk -a -t file -f " & Chr(34) & workImg & Chr(34) & " -m " & mountLetter & " -o rem"
    rcMount  = shell.Run(cmdMount, 0, True)
    If rcMount <> 0 Then
        ShowMessage "ImDisk failed mounting image for edit (exit code: " & rcMount & "):" & vbCrLf & cmdMount, vbCritical, SCRIPT_NAME
        WScript.Quit 1
    End If    
    
    
End If



' === Open Explorer on the mounted floppy and wait for it to close ===
'WaitForUser "Floppy image is now mounted at " & mountLetter & vbCrLf & "Explorer will open. The script will wait until you close all Explorer windows showing " & mountLetter & "."

' Open Explorer (don't wait, as it may return immediately)
shell.Run "explorer.exe " & mountLetter, 1, False

' Wait for Explorer window to actually appear (check multiple times)
Dim waitForOpen, openWaitCount
openWaitCount = 0
Do While Not HasExplorerWindows(mountLetter) And openWaitCount < 10
    WScript.Sleep 500 ' Check every half second
    openWaitCount = openWaitCount + 1
Loop

' Wait in a loop until all Explorer windows for this drive are closed
Dim checkCount, maxChecks
checkCount = 0
maxChecks = 1800 ' 30 minutes max (1800 * 1 second)

' Only start waiting if Explorer window actually opened
If HasExplorerWindows(mountLetter) Then
    Do While HasExplorerWindows(mountLetter) And checkCount < maxChecks
        WScript.Sleep 1000 ' Check every second
        checkCount = checkCount + 1
    Loop
    
    ' If we hit the timeout, inform the user
    If checkCount >= maxChecks Then
        ShowMessage "Timeout waiting for Explorer windows to close." & vbCrLf & _
                    "The script will now attempt to close them and continue.", _
                    vbExclamation, SCRIPT_NAME
    End If
Else
    ' Explorer didn't open or closed too quickly
    ' Give user a chance to work with drive manually if needed
    WaitForUser "Explorer window for " & mountLetter & " not detected or already closed." & vbCrLf & _
                "If you need to edit files, open " & mountLetter & " manually, make changes, close it, then click OK."
End If

' === Unmount image ===
Dim cmdUnmount, rcUnmount, retryCount
cmdUnmount = "imdisk -d -m " & mountLetter

' Try to unmount with retries (Explorer may still have handles)
retryCount = 0
Do While retryCount < 3
    rcUnmount = shell.Run(cmdUnmount, 0, True)
    If rcUnmount = 0 Then
        Exit Do ' Success
    End If
    retryCount = retryCount + 1
    If retryCount < 3 Then
        WScript.Sleep 2000 ' Wait 2 seconds before retry
    End If
Loop

If rcUnmount <> 0 Then
    ' Unmount failed - likely Explorer still has handles open
    ' This is OK - we'll clean it up on the next run
    ' Don't show error message, just note in console mode
    If Not useGUI Then
        WScript.Echo "Note: Image still mounted at " & mountLetter & " (will be cleaned up on next run)"
    End If
End If

' === Copy updated image back to USB ===
On Error Resume Next
fso.CopyFile workImg, srcImg, True
If Err.Number <> 0 Then

    ' Try to unmount before exiting
    SafeUnmount mountLetter
    
    ShowMessage "Failed copying updated image back to USB:" & vbCrLf & "From: " & workImg & vbCrLf & "To:   " & srcImg & vbCrLf &  "Error: " & Err.Description, vbCritical, SCRIPT_NAME
    WScript.Quit 1
End If
On Error GoTo 0

ShowMessage "Floppy image updated and copied back to " & srcImg & ".", _
            vbInformation, SCRIPT_NAME
