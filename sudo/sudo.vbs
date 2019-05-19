'********************************************************************************
' Script Name: sudo.vbs
' @(#) Purpose: This script allows you to 'sudo' commands similar to Linux
' @(#) by involking the Windows UAC window on the command if needed.
' @(#) Useful when you want to elivate one command quickly from an existing
' @(#) Command Prompt. Copy it to a folder in your %PATH% for easy use
'**********************************************************************************
' Written By: Brad Detchevery
' Created: Nov 19, 2018
'*********************************************************************************
Dim TheArgs
TheArgs=""
For Args=0 to WSCript.Arguments.length-1
TheArgs=TheArgs & " " & WScript.Arguments(Args)
Next

Set WshShell = WScript.CreateObject("WScript.Shell")
CurDir = WshShell.CurrentDirectory

Set objFSO=CreateObject("Scripting.FileSystemObject")
strDrive = objFSO.GetDriveName(CurDir)
Set d=objFSO.GetDrive(strDrive)
If d.ShareName <> "" THen
CurDir = Replace(CurDir,strDrive,d.ShareName)
End If
Dim tempFolder: tempFolder = objFSO.GetSpecialFolder(2)
Randomize
Dim RandomNum
RandomNum=(Int((100)*Rnd+100))

FileName=tempFolder & "\elevateme" & RandomNum & ".bat"

Set objFile = objFSO.CreateTextFile(FileName,True)
objFile.Write "@echo off" & vbCrLf
objFile.Write "pushd """ & CurDir & """" & vbCrLf
objFile.Write ltrim(TheArgs) & vbCrLf
objFile.Write "pause" & vbCrLf
objFile.Close

Set objShell = CreateObject("Shell.Application")

Dim TheCmd
TheCmd="cmd"
TheARgs = "/c " & FileName
objShell.ShellExecute TheCmd,TheArgs,"","runas",1
wscript.sleep(5000+RandomNum)
On Error Resume Next
objFSO.DeleteFile FileName
