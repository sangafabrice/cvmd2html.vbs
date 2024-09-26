''' <summary>
''' Launch the shortcut target PowerShell script with the selected markdown as an argument.
''' It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
''' </summary>
''' <version>0.0.1.5</version>
Option Explicit

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
  Dim objSink: Set objSink = WScript.CreateObject("WbemScripting.SWbemSink", "PwshProcess_")
  WaitForChildExit intCmdExeId, objSink
  objPackage.DeleteIconLink
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
''' <param name="intParentProcessId">The identifier of the parent process.</param>
''' <param name="objSink">The SWbemSink object for async execution.</param>
Sub WaitForChildExit(ByVal intParentProcessId, objSink)
  ' The process termination event query.
  ' Select the process whose parent is the intermediate process used for executing the link.
  Dim strWmiQuery: strWmiQuery = "SELECT * FROM __InstanceDeletionEvent WITHIN 0.1 WHERE TargetInstance ISA 'Win32_Process' AND TargetInstance.Name='pwsh.exe' AND TargetInstance.ParentProcessId=" & intParentProcessId
  ' Wait for the child process to exit.
  GetObject("winmgmts:").ExecNotificationQueryAsync objSink, strWmiQuery
  While Not objSink Is Nothing
    WScript.Sleep 1
  Wend
End Sub

''' <summary>
''' Expected to be called when the child process exits.
''' </summary>
Sub PwshProcess_OnObjectReady(ByVal objWbem, ByVal objAsyncContext)
  Set objWbem = Nothing
  Set objAsyncContext = Nothing
  objSink.Cancel
  Set objSink = Nothing
End Sub

''' <summary>
''' Expected to be called when objSink.Cancel is called.
''' </summary>
Sub PwshProcess_OnCompleted(ByVal intHResult, ByVal objWbem, ByVal objAsyncContext)
  ' If the HRresult message is not WBEM_E_CALL_CANCELLED.
  If &H80041032 <> intHResult Then
    MsgBox "An unhandled exception occured.", vbOKOnly + vbCritical, "Convert to HTML"
  End If
  Set objWbem = Nothing
  Set objAsyncContext = Nothing
  Set objSink = Nothing
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