''' <summary>
''' Exports the shortcut menu message box.
''' </summary>
''' <version>0.0.1</version>

Dim MessageBox: Set MessageBox = New MessageBoxType

''' <summary>
''' Represents the conversion watcher message box.
''' </summary>
Class MessageBoxType

  ''' <summary>
  ''' Show a warning message or an error message box.
  ''' </summary>
  ''' <param name="strMessage">The message text.</param>
  ''' <param name="varMessageType">The message box type (Warning/Error).</param>
  ''' <returns>"Yes" or "No" depending on the user's click when the message box is a warning.</returns>
  Function Show(ByVal strMessage, ByVal varMessageType)
    ' The default message box type is vbCritical for the error message.
    If varMessageType <> vbExclamation  And varMessageType <> vbCritical Then
      varMessageType = vbCritical
    End If
    ' The error message box shows the OK button alone.
    Dim varButton : varButton = vbOKOnly
    ' The warning message box shows the alternative Yes or No buttons.
    If varMessageType = vbExclamation Then
      varButton = vbYesNo
    End If
    ' Match the button clicked with its name string.
    Select Case MsgBox(strMessage, varMessageType + varButton, "Convert to HTML")
      Case vbYes
        Show = "Yes"
      Case vbNo
        Show = "No"
    End Select 
  End Function

End Class