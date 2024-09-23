''' <summary>
''' Returns information about the resource files used by the project.
''' It also provides a way to manage the custom icon link that can be installed and uninstalled.
''' </summary>
''' <version>0.0.1</version>

''' <summary>
''' Represents the package files used by the project.
''' </summary>
Class Package

  ''' <summary>
  ''' The package hashtable.
  ''' </summary>
  Private objPackage

  ''' <summary>
  ''' The file system object.
  ''' </summary>
  Private objFs

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
    Set objFs = CreateObject("Scripting.FileSystemObject")
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
  Sub CreateIconLink
    objFs.CreateTextFile(Me.IconLink.Path).Close
    With CreateObject("Shell.Application").NameSpace(Me.IconLink.DirName).ParseName(Me.IconLink.Name).GetLink
      .Path = Me.PwshExePath 
      .Arguments = Format("-ep Bypass -nop -w Hidden -f ""{0}"" -Markdown", Me.PwshScriptPath)
      .SetIconLocation Me.MenuIconPath, 0
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
  ''' Validate the link properties.
  ''' </summary>
  ''' <returns>True if the link properties are as expected.</returns>
  Function IsIconLinkValid
    Dim objLinkItem
    For Each objLinkItem In GetObject("winmgmts:Win32_ShortcutFile").Instances_
      With objLinkItem
        If Not StrComp(.Name, Me.IconLink.Path, vbTextCompare) Then
          IsIconLinkValid = Not StrComp(.Target, Me.PwshExePath, vbTextCompare)
          Set objLinkItem = Nothing
          Exit Function
        End If
      End With
      Set objLinkItem = Nothing
    Next
    Set objLinkItem = Nothing
    IsIconLinkValid = False
  End Function

  ''' <summary>
  ''' Get the PowerShell Core application path from the registry.
  ''' </summary>
  ''' <returns>The pwsh.exe full path.</returns>
  Private Function GetPwshPath
    ' The HKLM registry subkey stores the PowerShell Core application path.
    GetObject("winmgmts:StdRegProv").GetStringValue , "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pwsh.exe",, GetPwshPath
  End Function

  Private Sub Class_Terminate
    If Not IsEmpty(objPackage) Then
      Set objPackage("IconLink") = Nothing
      objPackage.RemoveAll
    End If
    Set objPackage = Nothing
    Set objFs = Nothing
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
    strDirName = CreateObject("WScript.Shell").SpecialFolders("StartMenu")
    strName = "cvmd2html.lnk"
    strPath = CreateObject("Scripting.FileSystemObject").BuildPath(strDirName, strName)
  End Sub

End Class