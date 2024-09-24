''' <summary>
''' Returns the methods for managing the shortcut menu option: install and uninstall.
''' </summary>
''' <version>0.0.1.1</version>

''' <summary>
''' Represents the setup methods for managing the shortcut records in the registry.
''' </summary>
Class Setup

  ''' <summary>
  ''' The format of the HKCU key.
  ''' </summary>
  Private KEY_FORMAT

  ''' <summary>
  ''' HKCU registry hive.
  ''' </summary>
  Private HKCU

  ''' <summary>
  ''' The shortcut menu verb subkey.
  ''' </summary>
  Private VERB_KEY

  Private Sub Class_Initialize
    KEY_FORMAT = "HKCU\{0}\"
    HKCU = &H80000001
    VERB_KEY = "SOFTWARE\Classes\SystemFileAssociations\.md\shell\cthtml"
  End Sub

  ''' <summary>
  ''' Configure the shortcut menu in the registry.
  ''' </summary>
  ''' <param name="blnParamNoIcon">Specifies that the custom menu icon should not be set.</param>
  ''' <param name="strMenuIconPath">The shortcut menu icon path.</param>
  Sub Install(ByVal blnParamNoIcon, ByVal strMenuIconPath)
    Dim strVerbKey: strVerbKey = Format(KEY_FORMAT, VERB_KEY)
    Dim strCommandKey: strCommandKey = strVerbKey & "command\"
    With New RegExp
      .Pattern = "\\cscript\.exe$"
      .IgnoreCase = True
      Dim strCommand: strCommand = Format("{0} ""{1}"" /Markdown:""%1""", Array(.Replace(WScript.FullName, "\wscript.exe"), WScript.ScriptFullName))
    End With
    With objWShell
      .RegWrite strCommandKey, strCommand
      .RegWrite strVerbKey, "Convert to &HTML"
      Dim strIconValueName: strIconValueName = strVerbKey & "Icon"
      If blnParamNoIcon Then
        On Error Resume Next
        .RegDelete strIconValueName
        On Error Goto 0
      Else
        .RegWrite strIconValueName, strMenuIconPath
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
    With GetObject("winmgmts:StdRegProv")
      .EnumKey HKCU, strKey, astrSNames
      If IsArray(astrSNames) Then
        For Each strSName In astrSNames
          DeleteSubkeyTree Format("{0}\{1}", Array(strKey, strSName))
        Next
      End If
      On Error Resume Next
      objWShell.RegDelete Format(KEY_FORMAT, strKey)
      On Error Goto 0
    End With
  End Sub

End Class