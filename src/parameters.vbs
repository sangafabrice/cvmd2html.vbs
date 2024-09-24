''' <summary>
''' Returns the parsed parameters.
''' </summary>
''' <version>0.0.1</version>

''' <summary>
''' Represents the input arguments and parameters.
''' </summary>
Class Parameters

  ''' <summary>
  ''' The parameters hashtable.
  ''' </summary>
  Private objParam

  ''' <summary>
  ''' The input markdown file path string.
  ''' </summary>
  Property Get Markdown
    Markdown = objParam("Markdown")
  End Property

  ''' <summary>
  ''' Installs the shortcut menu if it is true.
  ''' </summary>
  Property Get Install
    Install = objParam("Set")
  End Property

  ''' <summary>
  ''' Installs the shortcut menu without icon if it is true.
  ''' </summary>
  Property Get NoIcon
    NoIcon = objParam("NoIcon")
  End Property

  ''' <summary>
  ''' Uninstalls the shortcut menu if it is true.
  ''' </summary>
  Property Get Unset
    Unset = objParam("Unset")
  End Property

  Private Sub Class_Initialize
    Set objParam = GetParameters
  End Sub

  ''' <summary>
  ''' Get the input arguments and parameters.
  ''' </summary>
  ''' <returns>A hashtable of arguments.</returns>
  Private Function GetParameters
    Dim objWshArguments: Set objWshArguments = WScript.Arguments
    Dim objWshNamed: Set objWshNamed = objWshArguments.Named
    Dim intParamCount: intParamCount = objWshArguments.Count()
    Set GetParameters = CreateObject("Scripting.Dictionary")
    With GetParameters
      If intParamCount = 1 Then
        Dim strParamMarkdown: strParamMarkdown = objWshNamed("Markdown")
        If Len(strParamMarkdown) Then
          .Add "Markdown", strParamMarkdown
          Set objWshNamed = Nothing
          Set objWshArguments = Nothing
          Exit Function
        End If
        .Add "Set", objWshNamed.Exists("Set")
        If .Item("Set") Then
          Dim strNoIconParam: strNoIconParam = objWshNamed("Set")
          Dim blnIsNoIconParam: blnIsNoIconParam = CBool(Not StrComp(strNoIconParam, "NoIcon", vbTextCompare))
          If IsEmpty(strNoIconParam) Or blnIsNoIconParam Then
            .Add "NoIcon", blnIsNoIconParam
            Set objWshNamed = Nothing
            Set objWshArguments = Nothing
            Exit Function
          End If
        End If
        .RemoveAll
        .Add "Unset", objWshNamed.Exists("Unset") And IsEmpty(objWshNamed("Unset"))
        If .Item("Unset") Then
          Set objWshNamed = Nothing
          Set objWshArguments = Nothing
          Exit Function
        End If
        .RemoveAll
      ElseIf intParamCount = 0 Then
        .Add "Set", True
        .Add "NoIcon", False
        Set objWshNamed = Nothing
        Set objWshArguments = Nothing
        Exit Function
      End If
    End With
    Set GetParameters = Nothing
    Set objWshNamed = Nothing
    Set objWshArguments = Nothing
    ShowHelp
  End Function

  ''' <summary>
  ''' Show help and quit.
  ''' </summary>
  Private Sub ShowHelp
    Dim strHelpText: strHelpText = ""
    strHelpText = strHelpText & "The MarkdownToHtml shortcut launcher." & vbCrLf
    strHelpText = strHelpText & "It starts the shortcut menu target script in a hidden window." & vbCrLf & vbCrLf
    strHelpText = strHelpText & "Syntax:" & vbCrLf
    strHelpText = strHelpText & "  Convert-MarkdownToHtml.vbs /Markdown:<markdown file path>" & vbCrLf
    strHelpText = strHelpText & "  Convert-MarkdownToHtml.vbs [/Set[:NoIcon]]" & vbCrLf
    strHelpText = strHelpText & "  Convert-MarkdownToHtml.vbs /Unset" & vbCrLf
    strHelpText = strHelpText & "  Convert-MarkdownToHtml.vbs /Help" & vbCrLf & vbCrLf
    strHelpText = strHelpText & "<markdown file path>  The selected markdown's file path." & vbCrLf
    strHelpText = strHelpText & "                 Set  Configure the shortcut menu in the registry." & vbCrLf
    strHelpText = strHelpText & "              NoIcon  Specifies that the icon is not configured." & vbCrLf
    strHelpText = strHelpText & "               Unset  Removes the shortcut menu." & vbCrLf
    strHelpText = strHelpText & "                Help  Show the help doc." & vbCrLf
    With WScript
      .Echo strHelpText
      .Quit
    End With
  End Sub

  Private Sub Class_Terminate
    If Not IsEmpty(objParam) Then
      objParam.RemoveAll
    End If
    Set objParam = Nothing
  End Sub

End Class