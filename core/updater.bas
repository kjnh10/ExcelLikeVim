Attribute VB_Name = "updater"

Public Function UpdateSourcesFromGithub()
  Dim instruction As String
  Dim result As String
  ' instruction = "cd " & ThisWorkbook.Path & " & git clone https://github.com/kojinho10/vimx"
  instruction = "cd " & ThisWorkbook.Path & " & git pull origin master"
  Call ExecCommand(instruction, result)
  Msgbox result
End Function
