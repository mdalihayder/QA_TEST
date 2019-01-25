Function CloseProcess(strProgramName)
   Dim objshell
   Set objshell=CreateObject("WScript.Shell")
   objshell.Run "TASKKILL /F /IM "& strProgramName
   Set objshell=nothing
End Function

'close excel
 CloseProcess("Excel.EXE")

'Close word
'CloseProcess("WINWORD.exe")