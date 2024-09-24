''' <summary>
''' Launch the shortcut target PowerShell script with the selected markdown as an argument.
''' It aims to eliminate the flashing console window when the user clicks on the shortcut menu.
''' </summary>
''' <version>0.0.1.8</version>
Option Explicit

Dim objFs, objWShell
Set objFs = CreateObject("Scripting.FileSystemObject")
Set objWShell = CreateObject("WScript.Shell")

Imports "src\parameters.vbs"
Imports "src\package.vbs"

Dim objParam: Set objParam = New Parameters
Dim objPackage: Set objPackage = New Package

''' The application execution.
If objParam.RunLink Then
  Const WINDOW_STYLE_HIDDEN = 0
  Const WAIT_ON_RETURN = True
  objPackage.CreateIconLink objParam.Markdown
  objWShell.Run Format("""{0}""", objPackage.IconLink.Path), WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN
  objPackage.DeleteIconLink
  Quit
End If

If Not IsEmpty(objParam.Markdown) Then
  Imports "src\conhost.vbs"
  ConsoleHost.SetConsoleHostProperties objPackage.PwshExePath, objPackage.PwshScriptPath
  ConsoleHost.StartWith objParam.Markdown
  Set ConsoleHost = Nothing
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
''' Get the WSH runtime in GUI mode (wscript.exe).
''' </summary>
''' <returns>The WScript.Exe path.</returns>
Function GetDefaultCustomIconLinkTarget
  With objFs
    GetDefaultCustomIconLinkTarget = .BuildPath(.GetParentFolderName(WScript.FullName), "wscript.exe")
  End With
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
  Set objParam = Nothing
  Set objPackage = Nothing
  WScript.Quit
End Sub