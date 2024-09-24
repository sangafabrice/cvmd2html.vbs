''' <summary>
''' Launch the shortcut target PowerShell script with the selected markdown as an argument.
''' It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
''' </summary>
''' <version>0.0.1.3</version>
Option Explicit

RequestAdminPrivileges

Imports "src\parameters.vbs"
Imports "src\package.vbs"

Dim objParam: Set objParam = New Parameters
Dim objPackage: Set objPackage = New Package

''' The application execution.
If Not IsEmpty(objParam.Markdown) Then
  Imports "src\errorLog.vbs"
  Dim objErrorLog: Set objErrorLog = New ErrorLogHandler
  Const WINDOW_STYLE_HIDDEN = &HC
  Dim objStartInfo: Set objStartInfo = GetObject("winmgmts:Win32_ProcessStartup").SpawnInstance_
  objStartInfo.ShowWindow = WINDOW_STYLE_HIDDEN
  Dim intCmdExeId
  objPackage.CreateIconLink objParam.Markdown
  GetObject("winmgmts:Win32_Process").Create Format("C:\Windows\System32\cmd.exe /d /c """"{0}"" 2> ""{1}""""", Array(objPackage.IconLink.Path, objErrorLog.Path)),, objStartInfo, intCmdExeId
  If WaitForExit(intCmdExeId) Then
    With objErrorLog
      .Read
      .Delete
    End With
  End If
  objPackage.DeleteIconLink
  Set objStartInfo = Nothing
  Set objErrorLog = Nothing
  Quit
End If

''' Configuration and settings.
If objParam.Install Or objParam.Unset Then
  Imports "src\setup.vbs"
  Dim objSetup: Set objSetup = New Setup
  If objParam.Install Then
    objSetup.Install objParam.NoIcon, objPackage.MenuIconPath
  ElseIf objParam.Unset Then
    objSetup.Unset
  End If
  Set objSetup = Nothing
End If

Quit

''' <summary>
''' Wait for the process executing the link to exit.
''' </summary>
''' <param name="intProcessId">The identifier of the process.</param>
''' <returns>The process exit code.</returns>
Function WaitForExit(ByVal intProcessId)
  ' The process termination event query.
  Dim strWmiQuery: strWmiQuery = "SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName='cmd.exe' AND ProcessId=" & intProcessId
  ' Wait for the process to exit.
  WaitForExit = GetObject("winmgmts:").ExecNotificationQuery(strWmiQuery).NextEvent().ExitStatus
End Function

''' <summary>
''' Replace "{n}" by the nth input argument recursively.
''' </summary>
''' <param name="strFormat">The pattern format.</param>
''' <param name="astrArgs">The replacement texts.</param>
''' <returns>A text string.</returns>
Function Format(ByVal strFormat, ByVal astrArgs)
  If Not IsArray(astrArgs) Then
    Format = Replace(strFormat, "{0}", astrArgs)
    Exit Function
  End If
  Dim intBound: intBound = UBound(astrArgs)
  If intBound > -1 Then
    Dim strReplaceWith: strReplaceWith = astrArgs(intBound)
    Redim Preserve astrArgs(intBound - 1)
    Format = Format(Replace(strFormat, "{" & intBound &"}", strReplaceWith), astrArgs)
    Exit Function
  End If
  Format = strFormat
End Function

''' <summary>
''' Import the specified vbscript source file.
''' </summary>
''' <param name="strLibraryPath">the source file path.</param>
Sub Imports(ByVal strLibraryPath)
  On Error Resume Next
  Const FOR_READING = 1
  With CreateObject("Scripting.FileSystemObject")
    With .OpenTextFile(.BuildPath(.GetParentFolderName(WScript.ScriptFullName), strLibraryPath), FOR_READING)
      ExecuteGlobal .ReadAll
      .Close
    End With
  End With
End Sub

''' <summary>
''' Clean up and quit.
''' </summary>
Sub Quit
  Set objParam = Nothing
  Set objPackage = Nothing
  WScript.Quit
End Sub

''' <summary>
''' Request administrator privileges if standard user.
''' </summary>
Sub RequestAdminPrivileges
  Const HKU = &H80000003
  Dim blnIsAdmin
  GetObject("winmgmts:StdRegProv").CheckAccess HKU, "S-1-5-19\Environment",, blnIsAdmin
  If blnIsAdmin Then
    Exit Sub
  End If
  Const WINDOW_STYLE_HIDDEN = 0
  Dim strInputCommand: strInputCommand = Format("""{0}""", WScript.ScriptFullName)
  Dim strArgument
  For Each strArgument In WScript.Arguments
    strInputCommand = strInputCommand & Format(" ""{0}""", strArgument)
  Next
  CreateObject("Shell.Application").ShellExecute WScript.FullName, strInputCommand,, "runas", WINDOW_STYLE_HIDDEN
  WScript.Quit
End Sub