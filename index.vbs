''' <summary>
''' Launch the shortcut target PowerShell script with the selected markdown as an argument.
''' It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
''' </summary>
''' <version>0.0.1.7</version>
Option Explicit

Dim objFs, objWShell, objTypeLib
Set objFs = CreateObject("Scripting.FileSystemObject")
Set objWShell = CreateObject("WScript.Shell")
Set objTypeLib = CreateObject("Scriptlet.TypeLib")

Imports "src\parameters.vbs"
Imports "src\package.vbs"

Dim objParam: Set objParam = New Parameters
Dim objPackage: Set objPackage = New Package

''' The application execution.
If Not IsEmpty(objParam.Markdown) Then
  Imports "src\errorLog.vbs"
  Dim objErrorLog: Set objErrorLog = New ErrorLogHandler
  objPackage.CreateIconLink objParam.Markdown
  Const WINDOW_STYLE_HIDDEN = 0
  Const WAIT_ON_RETURN = True
  If objWShell.Run(Format("C:\Windows\System32\cmd.exe /d /c """"{0}"" 2> ""{1}""""", Array(objPackage.IconLink.Path, objErrorLog.Path)), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN) Then
    With objErrorLog
      .Read
      .Delete
    End With
  End If
  objPackage.DeleteIconLink
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
  With objFs
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
  Set objFs = Nothing
  Set objWShell = Nothing
  Set objTypeLib = Nothing
  Set objParam = Nothing
  Set objPackage = Nothing
  WScript.Quit
End Sub