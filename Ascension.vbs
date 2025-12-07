' Silent version - no visible window
Option Explicit

Dim objFSO, objShell, objWMIService
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

' Constants
Const ForReading = 1, ForWriting = 2, ForAppending = 8

' Settings file path (in user's AppData folder)
Dim settingsPath, appDataFolder
appDataFolder = objShell.ExpandEnvironmentStrings("%APPDATA%")
settingsPath = appDataFolder & "\AscensionLauncherHelper\settings.ini"

' Variables for paths
Dim filePath, launcherPath, gamePath, settingsFileExists

' Settings
Dim maxRuntime, normalLaunchTimeout, checkInterval
maxRuntime = 7200 ' 2 hours in seconds (safety timeout)
normalLaunchTimeout = 30 ' Normal launch should start Ascension within 30 seconds
checkInterval = 30 ' Check every 30 seconds during monitoring

' -------------------------------------------------------------------
' FUNCTIONS
' -------------------------------------------------------------------

' Create the settings directory if it doesn't exist
Sub CreateSettingsDirectory()
    Dim dirPath
    dirPath = objFSO.GetParentFolderName(settingsPath)
    If Not objFSO.FolderExists(dirPath) Then
        objFSO.CreateFolder(dirPath)
    End If
End Sub

' Save paths to settings file
Sub SaveSettings(launcherPath, gamePath)
    CreateSettingsDirectory
    
    Dim objFile
    Set objFile = objFSO.CreateTextFile(settingsPath, True)
    
    objFile.WriteLine("[Paths]")
    objFile.WriteLine("LauncherPath=" & launcherPath)
    objFile.WriteLine("GamePath=" & gamePath)
    objFile.WriteLine("LastUpdated=" & Now())
    
    objFile.Close
End Sub

' Load paths from settings file
Function LoadSettings()
    Dim launcherPath, gamePath, objFile, line, parts
    
    launcherPath = ""
    gamePath = ""
    
    If objFSO.FileExists(settingsPath) Then
        Set objFile = objFSO.OpenTextFile(settingsPath, ForReading)
        
        Do While Not objFile.AtEndOfStream
            line = Trim(objFile.ReadLine)
            
            If Left(line, 1) <> ";" And InStr(line, "=") > 0 Then
                parts = Split(line, "=")
                If UBound(parts) >= 1 Then
                    Select Case Trim(parts(0))
                        Case "LauncherPath"
                            launcherPath = Trim(parts(1))
                        Case "GamePath"
                            gamePath = Trim(parts(1))
                    End Select
                End If
            End If
        Loop
        
        objFile.Close
    End If
    
    LoadSettings = Array(launcherPath, gamePath)
End Function

' Browse for folder starting from My Computer
Function BrowseForFolderSimple(title)
    Dim shellApp, folder, folderPath
    
    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    
    If Err.Number = 0 Then
        ' Start from Desktop (most flexible)
        Set folder = shellApp.BrowseForFolder(0, title, 0, 0)
        
        If Not folder Is Nothing Then
            folderPath = folder.Self.Path
            BrowseForFolderSimple = folderPath
        Else
            BrowseForFolderSimple = ""
        End If
    Else
        ' If Shell.Application fails, fall back to input box
        Err.Clear
        folderPath = InputBox("Please enter the folder path:", title)
        BrowseForFolderSimple = folderPath
    End If
    
    On Error GoTo 0
End Function

' Search for executable in folder and subfolders
Function SearchForExecutable(folderPath, exeName)
    Dim fso, folder, subfolder, exePath
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        SearchForExecutable = ""
        Exit Function
    End If
    
    ' Check the main folder first
    exePath = folderPath & "\" & exeName
    If fso.FileExists(exePath) Then
        SearchForExecutable = exePath
        Exit Function
    End If
    
    ' Check immediate subfolders
    On Error Resume Next ' In case of permission errors
    
    Set folder = fso.GetFolder(folderPath)
    
    For Each subfolder In folder.SubFolders
        exePath = subfolder.Path & "\" & exeName
        If fso.FileExists(exePath) Then
            SearchForExecutable = exePath
            Exit Function
        End If
    Next
    
    On Error GoTo 0
    
    SearchForExecutable = ""
End Function

' Get launcher path - check defaults first
Function GetLauncherPath()
    Dim defaultPaths, path, launcherPath
    
    ' Common default locations for Ascension Launcher
    defaultPaths = Array( _
        "C:\Program Files\Ascension Launcher\Ascension Launcher.exe", _
        "C:\Program Files (x86)\Ascension Launcher\Ascension Launcher.exe", _
        "D:\Program Files\Ascension Launcher\Ascension Launcher.exe", _
        "E:\Program Files\Ascension Launcher\Ascension Launcher.exe", _
        "F:\Program Files\Ascension Launcher\Ascension Launcher.exe" _
    )
    
    ' Check if launcher exists in any default location
    For Each path In defaultPaths
        If objFSO.FileExists(path) Then
            ' Ask user if this is correct
            Dim response
            response = MsgBox("Found Ascension Launcher at:" & vbCrLf & path & vbCrLf & vbCrLf & _
                             "Is this correct?", vbYesNo + vbQuestion, "Confirm Launcher Location")
            
            If response = vbYes Then
                GetLauncherPath = path
                Exit Function
            End If
        End If
    Next
    
    ' If not found in defaults or user said no, browse for it
    launcherPath = GetPathWithBrowse("Ascension Launcher", "Ascension Launcher.exe")
    GetLauncherPath = launcherPath
End Function

' Get game path - check default resources folder first
Function GetGamePath(launcherPath)
    Dim resourcesPath, gamePath, clientFolders(), folder, exePath
    Dim i, folderList, selectedFolder, folderName
    
    ' Check if resources folder exists in default location
    If launcherPath <> "" Then
        resourcesPath = objFSO.GetParentFolderName(launcherPath) & "\resources"
        
        If objFSO.FolderExists(resourcesPath) Then
            ' Look for client folders
            ReDim clientFolders(-1) ' Empty array
            
            On Error Resume Next
            Set folder = objFSO.GetFolder(resourcesPath)
            
            For Each subfolder In folder.SubFolders
                ' Check if this folder contains Ascension.exe
                exePath = subfolder.Path & "\Ascension.exe"
                If objFSO.FileExists(exePath) Then
                    ' Add to array
                    ReDim Preserve clientFolders(UBound(clientFolders) + 1)
                    clientFolders(UBound(clientFolders)) = subfolder.Name
                End If
            Next
            On Error GoTo 0
            
            ' If we found client folders
            If UBound(clientFolders) >= 0 Then
                ' Build folder list for message box
                folderList = ""
                For i = 0 To UBound(clientFolders)
                    folderList = folderList & (i + 1) & ". " & clientFolders(i) & vbCrLf
                Next
                
                ' Ask user to select a folder
                selectedFolder = InputBox( _
                    "Found Ascension client folders in:" & vbCrLf & _
                    resourcesPath & vbCrLf & vbCrLf & _
                    "Available clients:" & vbCrLf & _
                    folderList & vbCrLf & _
                    "Enter the number of the client you want to use:", _
                    "Select Ascension Client", "1")
                
                If selectedFolder <> "" Then
                    ' Convert input to number
                    Dim folderIndex
                    folderIndex = CInt(selectedFolder) - 1
                    
                    If folderIndex >= 0 And folderIndex <= UBound(clientFolders) Then
                        folderName = clientFolders(folderIndex)
                        gamePath = resourcesPath & "\" & folderName & "\Ascension.exe"
                        
                        If objFSO.FileExists(gamePath) Then
                            GetGamePath = gamePath
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    ' If not found in resources folder or user cancelled, browse for it
    gamePath = GetPathWithBrowse("Ascension Game", "Ascension.exe")
    GetGamePath = gamePath
End Function

' Get path with browsing support (fallback method)
Function GetPathWithBrowse(title, exeName)
    Dim folderPath, exePath, attempts, maxAttempts, retryLoop
    
    attempts = 0
    maxAttempts = 3
    
    ' Show helpful instructions first
    MsgBox "Please browse to the folder containing " & exeName & _
           vbCrLf & vbCrLf & "You can navigate to any drive (C:, D:, E:, etc.)", _
           vbInformation, title
    
    Do While attempts < maxAttempts
        retryLoop = False
        
        ' Browse starting from Desktop
        folderPath = BrowseForFolderSimple("Select folder with " & exeName)
        
        If folderPath = "" Then
            ' User cancelled
            GetPathWithBrowse = ""
            Exit Function
        End If
        
        ' Check for executable in the selected folder
        exePath = folderPath & "\" & exeName
        
        If objFSO.FileExists(exePath) Then
            ' Found it!
            GetPathWithBrowse = exePath
            Exit Function
        Else
            ' Not found, try subfolders
            Dim foundPath
            foundPath = SearchForExecutable(folderPath, exeName)
            
            If foundPath <> "" Then
                GetPathWithBrowse = foundPath
                Exit Function
            End If
            
            ' Ask user what to do
            attempts = attempts + 1
            
            Dim response
            response = MsgBox(exeName & " was not found in:" & vbCrLf & _
                              folderPath & vbCrLf & vbCrLf & _
                              "Choose an option:" & vbCrLf & _
                              "Retry = Browse again" & vbCrLf & _
                              "Ignore = Enter path manually" & vbCrLf & _
                              "Cancel = Exit", _
                              vbAbortRetryIgnore + vbExclamation + vbDefaultButton1, _
                              "File Not Found")
            
            If response = vbRetry Then
                ' Try again - loop will continue
                retryLoop = True
            ElseIf response = vbIgnore Then
                ' Manual entry
                exePath = InputBox("Enter the full path to " & exeName & ":", _
                                   "Manual Entry", folderPath & "\" & exeName)
                exePath = Replace(exePath, """", "")
                
                If exePath <> "" And objFSO.FileExists(exePath) Then
                    GetPathWithBrowse = exePath
                    Exit Function
                ElseIf exePath <> "" Then
                    MsgBox "File not found: " & exePath, vbExclamation, "Error"
                    ' Let loop continue for another attempt
                    retryLoop = True
                Else
                    GetPathWithBrowse = ""
                    Exit Function
                End If
            Else
                ' vbAbort
                GetPathWithBrowse = ""
                Exit Function
            End If
        End If
        
        ' If we need to retry, continue the loop
        If Not retryLoop Then
            Exit Do
        End If
    Loop
    
    GetPathWithBrowse = ""
End Function

' Delete file if exists (with error handling)
Function DeleteFileIfExists(path)
    On Error Resume Next
    If objFSO.FileExists(path) Then
        objFSO.DeleteFile path, True
        If Err.Number = 0 Then
            DeleteFileIfExists = True
        Else
            ' File is in use or permission error
            DeleteFileIfExists = False
        End If
        Err.Clear
    Else
        DeleteFileIfExists = False ' File doesn't exist
    End If
    On Error GoTo 0
End Function

' Check if DLL exists
Function DLLExists()
    DLLExists = objFSO.FileExists(filePath)
End Function

' Check if process is running
Function IsProcessRunning(processName)
    Dim colProcesses
    IsProcessRunning = False
    
    On Error Resume Next
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & processName & "'")
    If Not colProcesses Is Nothing Then
        If colProcesses.Count > 0 Then
            IsProcessRunning = True
        End If
    End If
    On Error GoTo 0
End Function

' Try to delete DLL with retries (for update scenario)
Function DeleteDLLWithRetries()
    Dim attempts, maxAttempts, success
    attempts = 0
    maxAttempts = 3 ' Try 3 times
    success = False
    
    Do While attempts < maxAttempts And Not success
        success = DeleteFileIfExists(filePath)
        
        If Not success And DLLExists() Then
            ' File exists but couldn't delete - might be in use during update
            attempts = attempts + 1
            If attempts < maxAttempts Then
                ' Wait before retrying (longer wait for updates)
                WScript.Sleep 60000 ' Wait 1 minute before retry
            End If
        End If
    Loop
    
    DeleteDLLWithRetries = success
End Function

' -------------------------------------------------------------------
' MAIN LOGIC - PATH SETUP
' -------------------------------------------------------------------

' Try to load saved settings first
Dim savedPaths
savedPaths = LoadSettings()
launcherPath = savedPaths(0)
gamePath = savedPaths(1)

' Check if we need to ask for paths
settingsFileExists = objFSO.FileExists(settingsPath)

If Not settingsFileExists Or launcherPath = "" Or Not objFSO.FileExists(launcherPath) Then
    ' Get launcher path (checks defaults first)
    launcherPath = GetLauncherPath()
    
    If launcherPath = "" Then
        MsgBox "Launcher path is required. The script will now exit.", vbCritical, "Path Required"
        WScript.Quit
    End If
End If

If Not settingsFileExists Or gamePath = "" Or Not objFSO.FileExists(gamePath) Then
    ' Get game path (checks resources folder first)
    gamePath = GetGamePath(launcherPath)
    
    If gamePath = "" Then
        MsgBox "Game path is required. The script will now exit.", vbCritical, "Path Required"
        WScript.Quit
    End If
End If

' Save the paths for future use
If Not settingsFileExists Or launcherPath <> savedPaths(0) Or gamePath <> savedPaths(1) Then
    SaveSettings launcherPath, gamePath
    MsgBox "Paths have been saved for future use." & vbCrLf & _
           "Launcher: " & launcherPath & vbCrLf & _
           "Game: " & gamePath, vbInformation, "Settings Saved"
End If

' Set the DLL path based on game path
filePath = objFSO.GetParentFolderName(gamePath) & "\DivxTac.dll"

' Show DLL location for confirmation
'MsgBox "DLL monitoring location set to:" & vbCrLf & filePath, vbInformation, "Ready"

' -------------------------------------------------------------------
' MAIN LOGIC - DLL MONITORING AND CLEANUP
' -------------------------------------------------------------------

Dim launcherRunning, ascensionRunning, dllExistsAtStart, startTime, scriptStartTime
scriptStartTime = Timer ' Track total script runtime

' Check initial state
dllExistsAtStart = DLLExists()
launcherRunning = IsProcessRunning("Ascension Launcher.exe")
ascensionRunning = IsProcessRunning("Ascension.exe")

' Phase 1: Initial cleanup and launcher start
If dllExistsAtStart Then
    ' Try to delete DLL before starting launcher
    If Not DeleteFileIfExists(filePath) Then
        ' If can't delete, DLL might be in use (launcher already running or previous error)
        ' We'll handle this in monitoring phase
    End If
End If

' Start launcher if not already running
If Not launcherRunning Then
    objShell.Run """" & launcherPath & """", 1, False
    WScript.Sleep 5000 ' Wait 5 seconds for launcher to start
End If

' Phase 2: Wait for normal launch (30 seconds)
startTime = Timer
ascensionRunning = IsProcessRunning("Ascension.exe")

Do While (Timer - startTime) < normalLaunchTimeout And Not ascensionRunning
    WScript.Sleep 2000 ' Check every 2 seconds
    ascensionRunning = IsProcessRunning("Ascension.exe")
    
    ' Check if we should exit early
    If ascensionRunning And Not DLLExists() Then
        WScript.Quit ' Exit condition met
    End If
Loop

' Phase 3: Main monitoring loop (handles normal and update scenarios)
Dim monitoringStartTime, exitLoop
monitoringStartTime = Timer
exitLoop = False

Do While Not exitLoop And (Timer - scriptStartTime) < maxRuntime
    ' Check exit conditions
    If IsProcessRunning("Ascension.exe") And Not DLLExists() Then
        ' Condition 1: Ascension running AND DLL doesn't exist
        exitLoop = True
        Exit Do
    End If
    
    If Not IsProcessRunning("Ascension Launcher.exe") And Not DLLExists() Then
        ' Condition 2: Launcher closed AND DLL doesn't exist
        exitLoop = True
        Exit Do
    End If
    
    ' Only try to delete if DLL exists
    If DLLExists() Then
        ' Check current process states
        launcherRunning = IsProcessRunning("Ascension Launcher.exe")
        ascensionRunning = IsProcessRunning("Ascension.exe")
        
        If ascensionRunning Then
            ' Ascension is running - safe to delete DLL
            If DeleteDLLWithRetries() Then
                ' Successfully deleted
                If Not DLLExists() Then
                    exitLoop = True
                    Exit Do
                End If
            End If
        ElseIf launcherRunning Then
            ' Only launcher is running (update in progress)
            ' Don't attempt deletion - will fail during update
            ' Just wait and check again later
        Else
            ' Neither process running - safe to delete
            DeleteFileIfExists(filePath)
        End If
    End If
    
    ' If we haven't exited, wait before next check
    If Not exitLoop Then
        WScript.Sleep checkInterval * 1000 ' Convert to milliseconds
    End If
Loop

' Phase 4: Final cleanup attempt (if still needed)
If DLLExists() Then
    ' One last attempt if conditions allow
    launcherRunning = IsProcessRunning("Ascension Launcher.exe")
    ascensionRunning = IsProcessRunning("Ascension.exe")
    
    If ascensionRunning Then
        ' Ascension running - final delete attempt
        DeleteDLLWithRetries()
    ElseIf Not launcherRunning Then
        ' Launcher not running - safe to delete
        DeleteFileIfExists(filePath)
    End If
End If

' Check final exit condition
If IsProcessRunning("Ascension.exe") And Not DLLExists() Then
    ' Clean exit
    WScript.Quit
End If

' If we reach here, either:
' 1. 2-hour timeout reached
' 2. Some unexpected state
' Script will exit naturally