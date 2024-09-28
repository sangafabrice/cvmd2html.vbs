''' <summary>
''' The markdown to html converter.
''' </summary>
''' <version>0.0.1.3</version>

Imports "src\msgbox.vbs"

Dim MarkdownToHtml: Set MarkdownToHtml = New MarkdownToHtmlType

''' <summary>
''' Represents the markdown to html converter.
''' </summary>
Class MarkdownToHtmlType

  ''' <summary>
  ''' The javascript library path string.
  ''' </summary>
  Private strJsLibraryPath

  ''' <summary>
  ''' The path string of the html loading the library.
  ''' </summary>
  Private strHtmlLibraryPath

  ''' <summary>
  ''' Set the properties of the converter.
  ''' </summary>
  ''' <param name="strHtmlLibraryPathValue">The path string of the html loading the library.</param>
  ''' <param name="strJsLibraryPathValue">The javascript library path string.</param>
  Sub SetProperties(ByVal strHtmlLibraryPathValue, ByVal strJsLibraryPathValue)
    If IsEmpty(strHtmlLibraryPath) Then
      strHtmlLibraryPath = strHtmlLibraryPathValue
    End If
    If IsEmpty(strJsLibraryPath) Then
      strJsLibraryPath = strJsLibraryPathValue
    End If
  End Sub

  Sub ConvertFrom(strMarkdownPath)
    ' Validate the input markdown path string.
    If StrComp(objFs.GetExtensionName(strMarkdownPath), "md", vbTextCompare)  Then
      MessageBox.Show Format("""{0}"" is not a markdown (.md) file.", strMarkdownPath), vbCritical
    End If
    SetHtmlContent GetHtmlPath(strMarkdownPath), ConvertToHtml(GetContent(strMarkdownPath))
  End Sub

  ''' <summary>
  ''' This function returns the output path when it is unique without prompts or when the
  ''' user accepts to overwrite an existing HTML file. Otherwise, it exits the script.
  ''' </summary>
  ''' <param name="strMarkdownPath">The input markdown path argument.</param>
  ''' <returns>The output html path string.</returns>
  Private Function GetHtmlPath(ByVal strMarkdownPath)
    With objFs
      GetHtmlPath = Format("{0}\{1}.html", Array(.GetParentFolderName(strMarkdownPath), .GetBaseName(strMarkdownPath)))
      If .FileExists(GetHtmlPath) Then
        MessageBox.Show Format("The file ""{0}"" already exists.{1}{1}Do you want to overwrite it?", Array(GetHtmlPath, vbCrLf)), vbExclamation
      ElseIf .FolderExists(GetHtmlPath) Then
        MessageBox.Show Format("""{0}"" cannot be overwritten because it is a directory.", GetHtmlPath), vbCritical
      End If
    End With
  End Function

  ''' <summary>
  ''' Get the content of a file.
  ''' </summary>
  ''' <param name="strFilePath">The path that is read.</param>
  ''' <returns>The content of the file.</returns>
  Private Function GetContent(ByVal strFilePath)
    On Error Resume Next
    Const FOR_READING = 1
    Const PERMISSION_DENIED = 70
    Const FILE_NOT_FOUND = 53
    Const NO_ERROR = 0
    Err.Clear
    With objFs.OpenTextFile(strFilePath, FOR_READING)
      Select Case Err.Number
        Case NO_ERROR
          GetContent = .ReadAll
          .Close
        Case PERMISSION_DENIED, FILE_NOT_FOUND
          If Err.Number = PERMISSION_DENIED And Not objFs.FolderExists(strFilePath) Then
            MessageBox.Show Format("Access to the path ""{0}"" is denied.", strFilePath), vbCritical
          End If
          MessageBox.Show Format("File ""{0}"" is not found.", strFilePath), vbCritical
        Case Else
          MessageBox.Show format("Unspecified error trying to read from ""{0}"".", strFilePath), vbCritical
      End Select
    End With
  End Function

  ''' <summary>
  ''' Write the html text to the output HTML file.
  ''' It notifies the user when the operation did not complete with success.
  ''' </summary>
  ''' <param name="strHtmlPath">The output html path.</param>
  ''' <param name="strHtmlContent">The content of the html file.</param>
  Private Sub SetHtmlContent(ByVal strHtmlPath, ByVal strHtmlContent)
    On Error Resume Next
    Const FOR_WRITING = 2
    Err.Clear
    With objFs.OpenTextFile(strHtmlPath, FOR_WRITING, True)
      If Err.Number = 0 Then
        .Write strHtmlContent
        .Close
      ElseIf Err.Description = "Permission denied" Then
        MessageBox.Show Format("Access to the path ""{0}"" is denied.", strHtmlPath), vbCritical
      Else
        MessageBox.Show Format("Unspecified error trying to write to ""{0}"".", strHtmlPath), vbCritical
      End If
    End With
  End Sub

  ''' <summary>
  ''' Convert the content of the markdown file to html.
  ''' </summary>
  ''' <param name="strMarkdownContent">The markdown file content.</param>
  ''' <returns>The HTML text from converting the markdown text.</returns>
  Private Function ConvertToHtml(ByVal strMarkdownContent)
    On Error Resume Next
    With CreateObject("htmlFile")
      .Open
      .Write Format(GetContent(strHtmlLibraryPath), strJsLibraryPath)
      .Close
      While IsEmpty(.parentWindow.showdown)
        WScript.Sleep 1
      Wend
      ConvertToHtml = .parentWindow.convertMarkdown(strMarkdownContent)
    End With
  End Function

  Private Sub Class_Terminate
    Set MessageBox = Nothing
  End Sub

End Class