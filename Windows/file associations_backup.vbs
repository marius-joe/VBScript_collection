' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Save Windows file associations to backup XML file
'  v1.0.0
' ******************************************

Const C_Default_Path_Backup = "%USERPROFILE%\Downloads\Backup_FileAssociations.xml"

Dim objShell_A
Dim objShell_WS
Dim objFSO


    Call setGlobalsIfNecessary("objShell_A, objShell_WS, objFSO")
    Set listComplexArgs = getListComplexArguments()
    arg_path_Backup = C_Default_Path_Backup
    isSilentMode = False
    If Not isEmpty_ArrList(listComplexArgs) Then
        For Each arrComplexArg In listComplexArgs
            argument = UCase(arrComplexArg(0))
            value = ""
            If UBound(arrComplexArg) > 0 Then
                value = arrComplexArg(1)
            End If
            
            Select Case argument
                Case "-FILE"
                    arg_path_Backup = value
                    
                Case "-SILENT"    
                    isSilentMode = True     ' not used at the moment
            End Select
        Next
    End If
    
    If arg_path_Backup <> "" Then 
        path_Backup = objShell_WS.ExpandEnvironmentStrings(arg_path_Backup)
        path_Folder_Backup = objFSO.GetParentFolderName(path_Backup)
        Call ensurePath(path_Folder_Backup)
        
        strCmd = "dism /online /Export-DefaultAppAssociations:" & qt(path_Backup)

        strCmd = "cmd /c " & "echo " & strCmd & _
                 " & " & strCmd & _
                 " & echo;" & _
                 " & echo " & "Note: You can ONLY EDIT file associations in the xml, deleting whole lines will destroy the backup !" & _
                 " & echo;" & _
                 " & echo;" & _
                 " & PAUSE" 
        objShell_A.ShellExecute "cmd", strCmd, "", "runas", 1  
    End If

    Call cleanGlobals("All")
	


' Helper
Function qt(ByRef strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function

' Helper
' Not: return value is not for actual use, just for the internal recursion
Function ensurePath(ByRef strPath)  ' v1.2
	'Call setGlobalsIfNecessary("objFSO")
    If Not objFSO.FolderExists(strPath) Then
        Call ensurePath(objFSO.GetParentFolderName(strPath))
        Call objFSO.CreateFolder(strPath)
        isPathExist = false
    Else
        isPathExist = true
    End If
    ensurePath = isPathExist
End Function

' Helper Bundle
' ----------------------------------------------------
Function getListComplexArguments()		' v1.1
	Set objArgs = WScript.Arguments
	countArgs = objArgs.Count
	Set listComplexArguments = CreateObject("System.Collections.ArrayList")
	
	If countArgs > 0 Then
		strArgs = ""
        ' read all arguments (are seperated by " " by default)
		For i = 0 To countArgs-1
			strArgs = strArgs & " " & objArgs(i)
		Next

		If Contains(strArgs, "-") Then
			correctArgs = Split(strArgs, " -")
			UBCorrectArgs = UBound(correctArgs)
			For i = 1 To UBCorrectArgs
				arrArgument = Split(correctArgs(i), " ", 2)
				arrArgument(0) = "-" & arrArgument(0)
				listComplexArguments.Add arrArgument
			Next
		End If
	End If
	Set getListComplexArguments = listComplexArguments
End Function

Function Contains(ByRef str, ByRef strSearch)	' v1.2
	' converting to upper case is better than vbTextCompare because of dealing with foreign languages
	If InStr(UCase(str), UCase(strSearch)) > 0 Then returnValue = True Else returnValue = False
	Contains = returnValue
End Function

' Helper
Function isEmpty_ArrList(ByRef arrOrList)	' v1.6
    functionName = "isEmpty_ArrList"
	returnValue = True
	If IsArray(arrOrList) Then		' is array
		On Error Resume Next
			UBarr = UBound(arrOrList)
			If (Err.Number = 0) And (UBarr >= 0) Then returnValue = False
		On Error GoTo 0
	ElseIf TypeName(arrOrList) = "ArrayList" Then	 ' is list
        If arrOrList.Count > 0 Then
            returnValue = False
        End If
    Else
        Call show_MsgBox("Variable 'arrOrList' is no Array or ArrayList: " & TypeName(arrOrList), vbCritical, "Function: " & functionName)
    End If
	
	isEmpty_ArrList = returnValue
End Function
' ----------------------------------------------------

' Helper Bundle  v1.3
' ----------------------------------------------------
Sub setGlobalsIfNecessary(ByRef strObjectNames)
	arrObjectNames = Split(strObjectNames, ",")
	For Each strName In arrObjectNames
		strObj = UCase(Trim(strName))
		If strObj = UCase("objShell_A") Then
			If IsEmpty(objShell_A) Then Set objShell_A = CreateObject("Shell.Application")
		
		ElseIf strObj = UCase("objShell_WS") Then
			If IsEmpty(objShell_WS) Then Set objShell_WS = CreateObject("WScript.Shell")			
		
		ElseIf strObj = UCase("objFSO") Then
			If IsEmpty(objFSO) Then Set objFSO = CreateObject("Scripting.FileSystemObject")		
		End If
	Next
End Sub

Sub cleanGlobals(ByRef strObjectNames)
    If UCase(strObjectNames) = "ALL" Then
        arrObjectNames = Array("objShell_A", "objShell_WS", "objFSO")
    Else
        arrObjectNames = Split(strObjectNames, ",")
    End If
    
    For Each strName In arrObjectNames
        strObj = UCase(Trim(strName))
        If strObj = UCase("objShell_A") Then
            If Not IsEmpty(objShell_A) Then Set objShell_A = Nothing
        
        ElseIf strObj = UCase("objShell_WS") Then
            If Not IsEmpty(objShell_A) Then Set objShell_WS = Nothing
        
        ElseIf strObj = UCase("objFSO") Then
            If Not IsEmpty(objShell_A) Then Set objFSO = Nothing
        End If
    Next
End Sub
' ----------------------------------------------------