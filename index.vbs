''' <summary>
''' Launch the shortcut target PowerShell script with the selected markdown as an argument.
''' It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
''' </summary>
''' <version>0.0.1.1</version>
Option Explicit

Imports "src\parameters.vbs"
Imports "src\package.vbs"

Dim objParam: Set objParam = New Parameters
Dim objPackage: Set objPackage = New Package

''' The application execution.
If Not IsEmpty(objParam.Markdown) Then
  If Not objPackage.IsIconLinkValid Then
    Quit
  End If
  Imports "src\errorLog.vbs"
  Dim objErrorLog: Set objErrorLog = New ErrorLogHandler
  Const WINDOW_STYLE_HIDDEN = &HC
  Dim objStartInfo: Set objStartInfo = GetObject("winmgmts:Win32_ProcessStartup").SpawnInstance_
  objStartInfo.ShowWindow = WINDOW_STYLE_HIDDEN
  Dim intCmdExeId
  GetObject("winmgmts:Win32_Process").Create Format("C:\Windows\System32\cmd.exe /d /c """"{0}"" ""{1}"" 2> ""{2}""""", Array(objPackage.IconLink.Path, objParam.Markdown, objErrorLog.Path)),, objStartInfo, intCmdExeId
  On Error Resume Next
  WaitForExit intCmdExeId
  On Error Goto 0
  With objErrorLog
    .Read
    .Delete
  End With
  Set objStartInfo = Nothing
  Set objErrorLog = Nothing
  Quit
End If

''' Configuration and settings.
If objParam.Install Or objParam.Unset Then
  Imports "src\setup.vbs"
  Dim objSetup: Set objSetup = New Setup
  If objParam.Install Then
    objPackage.CreateIconLink
    objSetup.Install objParam.NoIcon, objPackage.MenuIconPath
  ElseIf objParam.Unset Then
    objSetup.Unset
    objPackage.DeleteIconLink
  End If
  Set objSetup = Nothing
End If

Quit

''' <summary>
''' Wait for the process executing the link to exit.
''' </summary>
''' <param name="intProcessId">The identifier of the process.</param>
Sub WaitForExit(ByVal intProcessId)
  Dim strMoniker: strMoniker = "winmgmts:Win32_Process.Handle=" & intProcessId
  While GetObject(strMoniker).Name = "cmd.exe" : Wend
End Sub

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