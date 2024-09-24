''' <summary>
''' The shortcut target script runner host.
''' </summary>
''' <version>0.0.1</version>

Imports "src\msgbox.vbs"

Dim ConsoleHost: Set ConsoleHost = New ConsoleHostType

''' <summary>
''' Represents the shortcut target script runner host.
''' </summary>
Class ConsoleHostType

  ''' <summary>
  ''' The powershell core path string.
  ''' </summary>
  Private strPwshExePath

  ''' <summary>
  ''' The target powershell script path.
  ''' </summary>
  Private strPwshScriptPath

  ''' <summary>
  ''' Set the private properties once.
  ''' </summary>
  ''' <param name="strPwshExePathValue">The powershell core path.</param>
  ''' <param name="strPwshScriptPathValue">The target powershell script path.</param>
  Sub SetConsoleHostProperties(ByVal strPwshExePathValue, ByVal strPwshScriptPathValue)
    If IsEmpty(strPwshExePath) Then
      strPwshExePath = strPwshExePathValue
    End If
    If IsEmpty(strPwshScriptPath) Then
      strPwshScriptPath = strPwshScriptPathValue
    End If
  End Sub

  ''' <summary>
  ''' Execute the shortcut target script runner and wait for its exit.
  ''' </summary>
  ''' <param name="strMarkdownPath">The input markdown file path.</param>
  Sub StartWith(ByVal strMarkdownPath)
    With objWShell
      WaitForExit .Exec(Format("""{0}"" -nop -ep Bypass -w Hidden -cwa ""try { Import-Module $args[0]; {3} -MarkdownPath $args[1] } catch { Write-Error $_.Exception.Message }"" ""{1}"" ""{2}""", Array(strPwshExePath, strPwshScriptPath, strMarkdownPath, objFs.GetBaseName(strPwshScriptPath))))
    End With
  End Sub

  ''' <summary>
  ''' Observe when the child process exits with or without an error.
  ''' Call the appropriate handler for each outcome.
  ''' </summary>
  ''' <param name="objPwshExe">The PowerShell Core process or child process.</param>
  Private Sub WaitForExit(ByVal objPwshExe)
    Dim objConsoleData: Set objConsoleData = New ConsoleData
    ' Wait for the process to complete.
    While objPwshExe.Status = 0 And objPwshExe.ExitCode = 0
      objConsoleData.objPwshExe_OutputDataReceived objPwshExe, objPwshExe.StdOut.ReadLine
    Wend
    ' When the process terminated with an error.
    If objPwshExe.ExitCode Then
      objPwshExe_ErrorDataReceived objPwshExe.StdErr.ReadAll
    End If
    Set objConsoleData = Nothing
  End Sub

  ''' <summary>
  ''' Show the error message that the child process writes on the console host.
  ''' </summary>
  ''' <param name="strErrData">The error message text.</param>
  Private Sub objPwshExe_ErrorDataReceived(ByVal strErrData)
    If Len(strErrData) > 0 Then
      ' Remove the ANSI escaped character for red coloring.
      With New RegExp
        .Pattern = "(\x1B\[31;1m)|(\x1B\[0m)"
        .Global = True
        strErrData = .Replace(strErrData, "")
      End With
      MessageBox.Show Mid(strErrData, InStr(strErrData, ":") + 2), vbCritical
    End If
  End Sub

  Private Sub Class_Terminate
    Set MessageBox = Nothing
  End Sub

End Class

Class ConsoleData

  ''' <summary>
  ''' The expected prompt from the console host.
  ''' </summary>
  Private strOverwritePromptText

  Private Sub Class_Initialize()
    strOverwritePromptText = ""
  End Sub

  ''' <summary>
  ''' Show the overwrite prompt that the child process sends. Handle the event when the
  ''' PowerShell Core (child) process redirects output to the parent Standard Output stream.
  ''' </summary>
  ''' <param name="objPwshExe">The sender child process.</param>
  ''' <param name="strOutData">The output text line sent.</param>
  Sub objPwshExe_OutputDataReceived(ByVal objPwshExe, ByVal strOutData)
    If Len(strOutData) > 0 Then
      ' Show the message box when the text line is a question.
      ' Otherwise, append the text line to the overall message text variable.
      If Right(RTrim(strOutData), 1) = "?" Then
        strOverwritePromptText = strOverwritePromptText & vbCrLf & strOutData
        ' Write the user's choice to the child process console host.
        objPwshExe.StdIn.WriteLine MessageBox.Show(strOverwritePromptText, vbExclamation)
        strOverwritePromptText = ""
      Else
        strOverwritePromptText = strOverwritePromptText & strOutData & vbCrLf
      End If
    End If
  End Sub

End Class