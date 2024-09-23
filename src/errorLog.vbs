''' <summary>
''' Manage the error log file and content.
''' </summary>
''' <version>0.0.1</version>

''' <summary>
''' Represents the input arguments and parameters.
''' </summary>
Class ErrorLogHandler

  ''' <summary>
  ''' The file system object.
  ''' </summary>
  Private objFs

  ''' <summary>
  ''' The error log file path.
  ''' </summary>
  Private strPath

  ''' <summary>
  ''' The error log file path.
  ''' </summary>
  Property Get Path
    Path = strPath
  End Property

  Private Sub Class_Initialize
    Set objFs = CreateObject("Scripting.FileSystemObject")
    strPath = objFs.BuildPath(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%"), LCase(Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)) & ".tmp.log")
  End Sub

  ''' <summary>
  ''' Display the content of the error log file in a message box if it is not empty.
  ''' </summary>
  Sub Read
    On Error Resume Next
    Const FOR_READING = 1
    With objFs.OpenTextFile(Me.Path, FOR_READING)
      Dim strErrorMessage: strErrorMessage = .ReadAll
      .Close
    End With
    If Len(strErrorMessage) Then
      ' Remove the ANSI escaped character for red coloring.
      With New RegExp
        .Pattern = "(\x1B\[31;1m)|(\x1B\[0m)"
        .Global = True
        MsgBox .Replace(strErrorMessage, ""), vbOKOnly + vbCritical, "Convert to HTML"
      End With
    End If
  End Sub

  ''' <summary>
  ''' Delete the error log file.
  ''' </summary>
  Sub Delete
    On Error Resume Next
    objFs.DeleteFile Me.Path
  End Sub

  Private Sub Class_Terminate
    Set objFs = Nothing
  End Sub

End Class