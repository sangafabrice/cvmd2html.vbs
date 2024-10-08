''' <summary>
''' Returns the methods for managing the shortcut menu option: install and uninstall.
''' </summary>
''' <version>0.0.1.3</version>

''' <summary>
''' Represents the setup methods for managing the shortcut records in the registry.
''' </summary>
Class Setup

  ''' <summary>
  ''' The parameters hashtable.
  ''' </summary>
  Private objRegistry

  ''' <summary>
  ''' HKCU registry hive.
  ''' </summary>
  Private HKCU

  ''' <summary>
  ''' The shortcut menu verb subkey.
  ''' </summary>
  Private VERB_KEY

  Private Sub Class_Initialize
    Set objRegistry = GetObject("winmgmts:StdRegProv")
    HKCU = &H80000001
    VERB_KEY = "SOFTWARE\Classes\SystemFileAssociations\.md\shell\cthtml"
  End Sub

  ''' <summary>
  ''' Configure the shortcut menu in the registry.
  ''' </summary>
  ''' <param name="blnParamNoIcon">Specifies that the custom menu icon should not be set.</param>
  ''' <param name="strMenuIconPath">The shortcut menu icon path.</param>
  Sub Install(ByVal blnParamNoIcon, ByVal strMenuIconPath)
    Dim strCommandKey: strCommandKey = VERB_KEY & "\command"
    Dim strCommand: strCommand = Format("{0} ""{1}"" /Markdown:""%1"" /RunLink", Array(GetDefaultCustomIconLinkTarget, WScript.ScriptFullName))
    With objRegistry
      .CreateKey HKCU, strCommandKey
      .SetStringValue HKCU, strCommandKey,, strCommand
      .SetStringValue HKCU, VERB_KEY,, "Convert to &HTML"
      Dim strIconValueName: strIconValueName = "Icon"
      If objParam.NoIcon Then
        .DeleteValue HKCU, VERB_KEY, strIconValueName
      Else
        .SetStringValue HKCU, VERB_KEY, strIconValueName, strMenuIconPath
      End If
    End With
  End Sub

  ''' <summary>
  ''' Remove the shortcut menu.
  ''' </summary>
  Sub Unset
    DeleteSubkeyTree VERB_KEY
  End Sub

  ''' <summary>
  ''' Remove the key and subkeys.
  ''' </summary>
  ''' <remarks>
  ''' Recursion is used because a key with subkeys cannot be deleted.
  ''' Recursion helps removing the leaf keys first.
  ''' </remarks>
  ''' <param name="strKey">A registry key.</param>
  Private Sub DeleteSubkeyTree(ByVal strKey)
    Dim astrSNames, strSName
    With objRegistry
      .EnumKey HKCU, strKey, astrSNames
      If IsArray(astrSNames) Then
        For Each strSName In astrSNames
          DeleteSubkeyTree Format("{0}\{1}", Array(strKey, strSName))
        Next
      End If
      .DeleteKey HKCU, strKey
    End With
  End Sub

  Private Sub Class_Terminate
    Set objRegistry = Nothing
  End Sub

End Class