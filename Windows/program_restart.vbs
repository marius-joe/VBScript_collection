' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Restart Program
'  v1.0.2
' ******************************************

Const C_RunWindowVisibility = 5		' 5 - Open the application with its window at its current size and position
									' 2 - Open the application with a minimized window
Dim objProcessList
Dim objShell_WS


    Set objShell_WS = CreateObject("WScript.Shell")

    isInputName = False
	Set objArgs = WScript.Arguments
	If objArgs.Count = 0 Then
        strMsg = "Please enter full program path !"
        path_program = LCase(InputBox(strMsg, "Restart Program"))
        isInputName = True
	Else
        path_program = LCase(objArgs(0))
	End If

    If path_program <> "" Then
        ' Exit program processes
        Call exit_program(path_program)

        ' Restart program
        objShell_WS.Run qt(path_program), C_RunWindowVisibility, false

	If isInputName Then
        msgType = vbInformation
        strMsg = """" & path_program & """ has been restarted !"
        strTitle = "Restart Program"
        MsgBox C_strSpace & strMsg, msgType + vbMsgBoxSetForeground, strTitle
	End If       
   End If
   
   Set objShell_WS = Nothing
   
   
   
' Helper
Function qt(ByVal strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function

' Helper Bundle
' v1.8
' ----------------------------------------------------
Function isRun_program(ByRef path_program)
    isRun_program = isRun_programWithArgs(path_program, "")
End Function

Function isRun_programWithArgs(ByRef path_program, ByRef args)
	Set colProcesses = get_RunningProcesses_withArgs(path_program, args)
    isRun_ProgramWithArgs = colProcesses.Count <> 0
	Set colProcesses = Nothing
End Function

Function get_runningProcesses(ByRef path_program)
    Set get_RunningProcesses = get_RunningProcesses_withArgs(path_program, "")
End Function

Function get_runningProcesses_withArgs(ByRef path_program, ByRef args)
	Set objWMI = GetObject("winmgmts:")
    If Contains(path_program, "\") Then
        nameOrExecutablePath = "ExecutablePath"
        path_program_esc = Replace(path_program, "\", "\\")
    Else
        nameOrExecutablePath = "Name"
        path_program_esc = path_program
    End If
    
    If args = "" Then
        sqlQuery = "SELECT * FROM win32_process WHERE " & nameOrExecutablePath & "='" & path_program_esc & "'"
    Else
        ' you have to escape the backslashes for wmi and the special characters for the like statement
        args_likeEsc = replaceMultiple(args, Array(Array("\", "\\"), _
                                                   Array("[", "[[]"), _
                                                   Array("%", "[%]"), _
                                                   Array("_", "[_]")))
        sqlQuery = "SELECT * FROM win32_process WHERE " & nameOrExecutablePath & "='" & path_program_esc & "'" & _
                   " and CommandLine LIKE '%" & args_likeEsc & "%'"
    End If
    Set get_RunningProcesses_withArgs = objWMI.ExecQuery(sqlQuery)
    Set objWMI = Nothing
End Function

Function isRun_Process(ByRef processID)
	Set objWMI = GetObject("winmgmts:")
    sqlQuery = "SELECT * FROM win32_process WHERE ProcessId='" & processID & "'"
    Set colProcesses = objWMI.ExecQuery(sqlQuery)
    isRun_Process = colProcesses.Count <> 0
    Set objWMI = Nothing
End Function

Sub exit_program(ByRef path_program)
    Call exit_programWithArgs(path_program, "")
End Sub

Sub exit_programWithArgs(ByRef path_program, ByRef args)
    Set colProcesses = get_runningProcesses_withArgs(path_program, args)
    For Each objProcess in colProcesses
        processID = objProcess.processID
        objProcess.Terminate()
        waitForExit_process(processID)
    Next
    Set colProcesses = Nothing
End Sub

Sub waitForExit_program(ByRef path_program)
    Call waitForExit_programWithArgs(path_program, "")
End Sub

Sub waitForExit_programWithArgs(ByRef path_program, ByRef args)
	reCheckInterval_ms = 200
	Do While isRun_programWithArgs(path_program, args)
		WScript.Sleep reCheckInterval_ms
	Loop
End Sub

Sub waitForExit_process(ByRef processID)
	reCheckInterval_ms = 200
	Do While isRun_Process(processID)
		WScript.Sleep reCheckInterval_ms
	Loop
End Sub

Function Contains(ByRef str, ByRef strSearch)	' v1.3
	' converting to lower case is better than vbTextCompare because of dealing with foreign languages
	If InStr(LCase(str), LCase(strSearch)) > 0 Then returnValue = True Else returnValue = False
	Contains = returnValue
End Function
' ----------------------------------------------------