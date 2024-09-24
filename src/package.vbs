''' <summary>
''' Returns information about the resource files used by the project.
''' It also provides a way to manage the custom icon link that can be installed and uninstalled.
''' </summary>
''' <version>0.0.1.2</version>

''' <summary>
''' Represents the package files used by the project.
''' </summary>
Class Package

  ''' <summary>
  ''' The package hashtable.
  ''' </summary>
  Private objPackage

  ''' <summary>
  ''' The project root path string.
  ''' </summary>
  Property Get Root
    Root = objPackage("Root")
  End Property

  ''' <summary>
  ''' The project resources directory path string.
  ''' </summary>
  Property Get ResourcePath
    ResourcePath = objPackage("ResourcePath")
  End Property

  ''' <summary>
  ''' The shortcut target powershell script path string.
  ''' </summary>
  Property Get PwshScriptPath
    PwshScriptPath = objPackage("PwshScriptPath")
  End Property

  ''' <summary>
  ''' The shortcut menu icon path string.
  ''' </summary>
  Property Get MenuIconPath
    MenuIconPath = objPackage("MenuIconPath")
  End Property

  ''' <summary>
  ''' The powershell core runtime path string.
  ''' </summary>
  Property Get PwshExePath
    PwshExePath = objPackage("PwshExePath")
  End Property

  ''' <summary>
  ''' The adapted custom icon link object.
  ''' </summary>
  Property Get IconLink
    Set IconLink = objPackage("IconLink")
  End Property

  Private Sub Class_Initialize
    Set objPackage = CreateObject("Scripting.Dictionary")
    With objPackage
      .Add "Root", objFs.GetParentFolderName(WScript.ScriptFullName)
      .Add "ResourcePath", objFs.BuildPath(.Item("Root"), "rsc")
      .Add "PwshScriptPath", objFs.BuildPath(.Item("ResourcePath"), "cvmd2html.ps1")
      .Add "MenuIconPath", objFs.BuildPath(.Item("ResourcePath"), "menu.ico")
      .Add "PwshExePath", GetPwshPath
      .Add "IconLink", New IconLinkResource
    End With
  End Sub

  ''' <summary>
  ''' Create the custom icon link file.
  ''' </summary>
  ''' <param name="strMarkdownPath">The input markdown file path.</param>
  Sub CreateIconLink(ByVal strMarkdownPath)
    With GetCustomIconLink
      .TargetPath = Me.PwshExePath 
      .Arguments = Format("-ep Bypass -nop -w Hidden -f ""{0}"" -Markdown ""{1}""", Array(Me.PwshScriptPath, strMarkdownPath))
      .IconLocation = Me.MenuIconPath
      .Save
    End With
  End Sub

  ''' <summary>
  ''' Delete the custom icon link file.
  ''' </summary>
  Sub DeleteIconLink
    On Error Resume Next
    objFs.DeleteFile Me.IconLink.Path
  End Sub

  ''' <summary>
  ''' Get the custom icon link object.
  ''' </summary>
  ''' <returns>The specified link file object.</returns>
  Private Function GetCustomIconLink
    Set GetCustomIconLink = objWShell.CreateShortcut(Me.IconLink.Path)
  End Function

  ''' <summary>
  ''' Get the PowerShell Core application path from the registry.
  ''' </summary>
  ''' <returns>The pwsh.exe full path.</returns>
  Private Function GetPwshPath
    ' The HKLM registry subkey stores the PowerShell Core application path.
    GetPwshPath = objWShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pwsh.exe\")
  End Function

  Private Sub Class_Terminate
    If Not IsEmpty(objPackage) Then
      Set objPackage("IconLink") = Nothing
      objPackage.RemoveAll
    End If
    Set objPackage = Nothing
  End Sub

End Class

''' <summary>
''' Represents an adapted custom icon link object.
''' </summary>
Class IconLinkResource

  ''' <summary>
  ''' The segments of the icon link full path.
  ''' </summary>
  Private strDirName, strName, strPath

  ''' <summary>
  ''' The custom icon parent directory path string.
  ''' </summary>
  Property Get DirName
    DirName = strDirName
  End Property

  ''' <summary>
  ''' The custome icon file name string.
  ''' </summary>
  Property Get Name
    Name = strName
  End Property

  ''' <summary>
  ''' The custom icon file full path string.
  ''' </summary>
  Property Get Path
    Path = strPath
  End Property

  Private Sub Class_Initialize()
    strDirName = objWShell.ExpandEnvironmentStrings("%TEMP%")
    strName = LCase(Mid(objTypeLib.Guid, 2, 36)) & ".tmp.lnk"
    strPath = objFs.BuildPath(strDirName, strName)
  End Sub

End Class