''' <summary>
''' Exports the shortcut menu message box.
''' </summary>
''' <version>0.0.1.1</version>

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
    ' Quit the application when the user presses Ok or No buttons.
    Select Case MsgBox(strMessage, varMessageType + varButton, "Convert to HTML")
      Case vbOK, vbNo
        Quit
    End Select 
  End Function

End Class