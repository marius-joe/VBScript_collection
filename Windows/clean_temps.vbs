' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Clean temporary folders
'  v1.0.0
' ******************************************


'Papierkorb löschen noch einbauen mit abfrage
'Dateien die gerade in Gebrauch sind ?

Const C_strSpace = "   "    ' space for text in msgBoxes

Dim objShell, objFSO, objSysEnv,objUserEnv
Dim strUserTemp, strSysTemp
Dim userProfile, TempInternetFiles
Dim OSType
Dim isSilentMode


	isSilentMode = False
	Set objArgs = WScript.Arguments
	If objArgs.Count > 0 Then
		If LCase(objArgs(0)) = "silent" Then
			isSilentMode = True
		End If
	End If    

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("WScript.Shell")
	Set objSysEnv = objShell.Environment("System")
	Set objUserEnv = objShell.Environment("User")

	strUserTemp = objShell.ExpandEnvironmentStrings(objUserEnv("TEMP"))
	strSysTemp = objShell.ExpandEnvironmentStrings(objSysEnv("TEMP"))
	userProfile = objShell.ExpandEnvironmentStrings("%userprofile%")
    
	REM ' the Internet Temp files path is diffrent according to OS Type
	REM OSType = FindOSType
	REM If OSType = "Windows 7" OR OSType = "Windows Vista" Then
		REM TempInternetFiles = userProfile & "\AppData\Local\Microsoft\Windows\Temporary Internet Files"
	REM ElseIf  OSType = "Windows 2003" OR OSType = "Windows XP" Then
		REM TempInternetFiles = userProfile & "\Local Settings\Temporary Internet Files"
	REM End If
    
    TempInternetFiles = "'TempInternetFiles' will not be deleted"

	If Not isSilentMode Then
		msgType = vbExclamation
		arrMessage = Array("Temporary Files will be deleted, please wait for the process to finish !", _
							"", _
							strSysTemp, _
							strUserTemp, _
							TempInternetFiles)

		MsgBox C_strSpace & Join(arrMessage, vbCrlf & C_strSpace), vbMsgBoxSetForeground + msgType, "Clean Temps"
	End If     
    
	Call deleteRecursive(strUserTemp, true) ' delete user temp files
	Call deleteRecursive(strSysTemp, true)  ' delete system temp files

	' delete Internet Temp files
	REM Call deleteRecursive(TempInternetFiles, true)

    TempInternetFiles = "'TempInternetFiles' not deleted"
	If Not isSilentMode Then
		msgType = vbInformation
		arrMessage = Array("Temporary Files deleted !", _
							"", _
							strSysTemp, _
							strUserTemp, _
							TempInternetFiles)

		MsgBox C_strSpace & Join(arrMessage, vbCrlf & C_strSpace), vbMsgBoxSetForeground + msgType, "Clean Temps"
	End If



Private Sub deleteRecursive(path, keepRoot)
	On Error Resume Next
	
	Dim folder, file, subFolder

	If objFSO.FolderExists(path) Then
		Set folder = objFSO.GetFolder(path)

		' remove subfolders:
		For Each subFolder In Folder.SubFolders
			Call deleteRecursive(subFolder, false)
		Next

		' remove files
		For Each file In Folder.Files
			objFSO.DeleteFile(file.path)
		Next
		
'Verbesserungen (z.B. minimales Alter der Dateien):

REM #Sets variable Date to current date
REM $Date=get-date
REM #Sets Date variable back 14 days
REM $Date=$Date.AddDays(-14)
REM #Lists all files in the backup folder that were last written to over 14 days ago and deletes them
REM Get-ChildItem 'E:\Backup' -recurse | where-object {$_.LastWriteTime -lt $date} | remove-item -recurse

'ODER


'robocopy e:\empty e:\to_delete /MIR /E
'e:\empty is a empty dir. its fast :-)
'To speed things up you can use the /CREATE option which will create zero-length files.#


REM mkdir c:\delete
REM robocopy c:\Source c:\Delete /e /MOVE /MINAGE:2
REM rmdir c:\delete /s /q

		If Not keepRoot Then
			' remove root folder (can only be deleted if no content is currently in use)
			objFSO.DeleteFolder(path)
		End If

	ElseIf objFSO.FileExists(path) Then
		objFSO.DeleteFile(path)
	End If
	
	On Error Goto 0
End Sub


Function FindOSType
	Dim objWMI, objItem, colItems
	Dim OSVersion, OSName
	Dim ComputerName

	ComputerName = "."

	' Get the WMI object and query results
	Set objWMI = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
	Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem",,48)

	' Get the OS version number (first two) and OS product type (server or desktop) 
	For Each objItem In colItems
		OSVersion = Left(objItem.Version,3)     
	Next

	Select Case OSVersion
		Case "6.1"
			OSName = "Windows 7"
		Case "6.0" 
			OSName = "Windows Vista"
		Case "5.2" 
			OSName = "Windows 2003"
		Case "5.1" 
			OSName = "Windows XP"
		Case "5.0" 
			OSName = "Windows 2000"
	End Select

	' Return the OS name
	FindOSType = OSName

	Set colItems = Nothing
	Set objWMI = Nothing
End Function