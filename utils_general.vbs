' ******************************************
'  Dev:  marius-joe
' ******************************************
'  VBScript Utils
'  v1.3.2
' ******************************************


Const C_strSpace = "   "    ' space for text in msgBoxes



' Helper Bundle - v1.6
' ----------------------------------------------------
Function parseArguments()
    Dim arrComplexArgs()
	Set objArgs = WScript.Arguments
	countArgs = objArgs.Count
	If countArgs > 0 Then
		strArgs = " """ & objArgs(0)
        ' read all arguments (seperated by " " by default)
        ' mark each beginning of an argument part
        ' " can quite safely be used for that, because the " got removed when args were handed to the script
		For i = 1 To countArgs-1
			strArgs = strArgs & " """ & objArgs(i)
		Next
		If Contains(strArgs, " ""-") Then
			arrCorrectArgs = Split(strArgs, " ""-")
			UBarrCorrectArgs = UBound(arrCorrectArgs)
            ReDim arrComplexArgs(UBarrCorrectArgs-1)
			For i = 1 To UBarrCorrectArgs
				arrArgument = Split(Trim(arrCorrectArgs(i)), " """, 2)
				arrArgument(0) = "-" & arrArgument(0)
                If UBound(arrArgument) > 0 Then arrArgument(1) = arrArgument(1)
                arrComplexArgs(i-1) = arrArgument
			Next
		End If
	End If
	parseArguments = arrComplexArgs
End Function

' v1.3
Function Contains(ByRef str, ByRef strSearch)
	' converting to lower case is better than vbTextCompare because of dealing with foreign languages
	If InStr(LCase(str), LCase(strSearch)) > 0 Then returnValue = True Else returnValue = False
	Contains = returnValue
End Function

' v1.9
Function isEmpty_ArrList(ByRef arrOrList)
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
        returnValue = "Variable 'arrOrList' is no Array or ArrayList: " & TypeName(arrOrList)
    End If
	isEmpty_ArrList = returnValue
End Function

' v1.1
Function arr_SafeGet(ByRef arr, ByRef index, ByRef defaultValue)
    If UBound(arr) >= index Then arr_SafeGet = arr(index) Else arr_SafeGet = defaultValue
End Function
' ----------------------------------------------------


' Helper Bundle - v1.5
' req: Dim objShell_A, objShell_WS, objFSO
' use: Call setGlobalsIfNecessary("objFSO, objShell_WS")
'      Call cleanGlobals("All") | Call cleanGlobals("objFSO, objShell_WS")
' ----------------------------------------------------
Sub setGlobalsIfNecessary(ByRef strObjectNames)
	arrObjectNames = Split(strObjectNames, ",")
	For Each strName In arrObjectNames
		strObj = LCase(Trim(strName))
		If strObj = LCase("objShell_A") Then
			If IsEmpty(objShell_A) Then Set objShell_A = CreateObject("Shell.Application")
		
		ElseIf strObj = LCase("objShell_WS") Then
			If IsEmpty(objShell_WS) Then Set objShell_WS = CreateObject("WScript.Shell")			
		
		ElseIf strObj = LCase("objFSO") Then
			If IsEmpty(objFSO) Then Set objFSO = CreateObject("Scripting.FileSystemObject")		
		End If
	Next
End Sub

Sub cleanGlobals(ByRef strObjectNames)
    If LCase(strObjectNames) = "all" Then
        arrObjectNames = Array("objShell_A", "objShell_WS", "objFSO")
    Else
        arrObjectNames = Split(strObjectNames, ",")
    End If
    
    For Each strName In arrObjectNames
        strObj = LCase(Trim(strName))
        If strObj = LCase("objShell_A") Then
            If Not IsEmpty(objShell_A) Then Set objShell_A = Nothing
        
        ElseIf strObj = LCase("objShell_WS") Then
            If Not IsEmpty(objShell_A) Then Set objShell_WS = Nothing
        
        ElseIf strObj = LCase("objFSO") Then
            If Not IsEmpty(objShell_A) Then Set objFSO = Nothing
        End If
    Next
End Sub
' ----------------------------------------------------


' Helper Bundle - v1.5
' req: setGlobalsIfNecessary()
' ----------------------------------------------------
Sub restartAsAdmin()	
	Call setGlobalsIfNecessary("objShell_A, objShell_WS, objFSO")
	Set objArgs = WScript.Arguments
	countArgs = objArgs.Count
	isElevated = False
	If countArgs > 0 Then
		For Each arg In objArgs
			If arg = "-admin" Then
				isElevated = True
				Exit For
			End If
		Next
	End If
	
	If Not isElevated Then
		path_ThisScript = WScript.ScriptFullName
		isScriptReady = False
        isScriptInTemp = False
		
		' when script is running from network path, copy it to local temp folder for better compability
		If isOnNetworkDrive(path_ThisScript) Then
			path_DestFolder = objShell_WS.ExpandEnvironmentStrings("%TEMP%")
			Call roboCopy(path_ThisScript, path_DestFolder, False, "")
			
			path_ScriptFile = objFSO.GetFile(objFSO.BuildPath(path_DestFolder, WScript.ScriptName))

			If path_ScriptFile <> "" Then
				isScriptReady = True
                isScriptInTemp = True
			Else
				msgType = vbCritical
				MsgBox C_strSpace & "Script could not be copied to %TEMP% and be run as admin !", vbMsgBoxSetForeground + msgType, "Restart script as admin"			
			End If
		Else
			path_ScriptFile = path_ThisScript
			isScriptReady = True
		End If
			
		If isScriptReady Then
			objShell_A.ShellExecute "wscript", qt(path_ScriptFile) & " -admin", "", "runas", C_RunWindowVisibility
            If isScriptInTemp Then
                WScript.Sleep 200
                objFSO.DeleteFile path_ScriptFile, true     ' give wscript some time to load the script in memory and then delete it from the temp folder
            End If
        End If
		WScript.Quit
	End If
End Sub

' Helper - v1.1
Function isOnNetworkDrive(ByRef path)
	Set objNetwork = WScript.CreateObject("WScript.Network")
	Set objDrives = objNetwork.EnumNetworkDrives
	
	returnValue = False
	countDrives = objDrives.Count
	For i = 0 To countDrives - 1 Step 2
		driveLetter = objDrives.Item(i)
		'driveAdress = objDrives.Item(i+1)
		If StartsWith(path, driveLetter) Then
			returnValue = True
			Exit For
		End If
	Next
	isOnNetworkDrive = returnValue
End Function

' Helper - v1.2
' req: setGlobalsIfNecessary()
Function ensurePath(ByRef strPath)  ' Return value is not for use, just for the recursion
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

' Helper - v1.0
' req: setGlobalsIfNecessary()
Function IsFolder(ByVal sPath)
	Call setGlobalsIfNecessary("objFSO")
    IsFolder = objFSO.FolderExists(sPath)
End Function

' Helper - v1.0
' req: setGlobalsIfNecessary()
Function IsFile(ByVal sPath)
	Call setGlobalsIfNecessary("objFSO")
    IsFile = objFSO.FileExists(sPath)
End Function

' Helper - v1.3 improved version in RoboCopy.vbs
' Quellpfad Formate:
' C:\Users\folderToCopy
' C:\Users\*
' C:\Users\*.txt
' C:\Users\copyMe.txt
Sub roboCopy(ByRef path_Source, ByRef path_DestFolder, ByRef createDestFolder, ByRef moveData, ByRef strMyArguments)
    If isValidPath(path_Source, true) Then      ' true: allow wildcard in path
        If ensurePath(path_DestFolder) Then
            indexLastBackSlash = InStrRev(path_Source, "\")
            path_SourceFolder = Left(path_Source, indexLastBackSlash-1)
            file = Mid(path_Source, indexLastBackSlash+1)
            
            strRoboCmd = "RoboCopy.Exe " & qt(path_SourceFolder) & " " & qt(path_DestFolder)
            strRetrySettings = "/R:3 /W:5"		'try 3x 5 Sec
            
            If moveData Then strArgs = "/MOVE " Else strArgs = ""
            If strMyArguments <> "" Then strMyArguments = " " & strMyArguments
            
            If file = "*" Then					                ' copy all files and subfolder, no empty folder
                strArgs = strArgs & "/S" & strMyArguments
            Else
                strArgs = strArgs & qt(file) & strMyArguments	' copy files or special file extensions
            End If
               
            Call setGlobalsIfNecessary("objShell_WS")
            objShell_WS.Run strRoboCmd & " " & strArgs & " " & strRetrySettings, 0, true	'wait until cmd has finished
        Else
            MsgBox "No Such destFolder"
        End If
    Else
        MsgBox "Error with Source Path"
    End If
End Sub

' Helper - v1.5
' req: setGlobalsIfNecessary()
Function isValidPath(ByRef path, ByRef allowWildCard)
    isValid = False
    Call setGlobalsIfNecessary("objFSO")
    If objFSO.FileExists(path) OR objFSO.FolderExists(path) Then
        isValid = True
        
    ElseIf allowWildCard Then
        indexLastBackSlash = InStrRev(path, "\")
        path_ParentFolder = Left(path, indexLastBackSlash-1)
   
        If objFSO.FolderExists(path_ParentFolder) Then
            wildCard = Mid(path, indexLastBackSlash+1)
         
            Set objRegEx = New RegExp
            With objRegEx
                .Pattern = "^(\*|(\*(\.\w+)+))$"
            End With
            isValid = objRegEx.Test(wildCard)
        End If
    End If
    isValidPath = isValid
End Function
' ----------------------------------------------------


' Helper - v1.1
' req: setGlobalsIfNecessary()
Function GetOutlookVersionNumber()
	Call setGlobalsIfNecessary("objShell_WS")	
	' Read the Classes Root registry hive (it is a memory-only instance of HKCU\Software\Classes and HKLM\Software\Classes registry keys)
	' as it contains a source of information for the currently active Microsoft Office Outlook application version -
	' it's quicker and easier to read the registry than the file version information after a location lookup).
	' The trailing backslash on the line means that the @ or default registry key value is being queried.
	sTempValue = objShell_WS.RegRead("HKCR\Outlook.Application\CurVer\")
	
	' Check the length of the value found and if greater than 2 digits then read the last two digits for the major Office version value
	If Len(sTempValue) > 2 Then 
		GetOutlookVersionNumber = Replace(Right(sTempValue, 2), ".", "")
	Else
		GetOutlookVersionNumber = ""
	End If
End Function


' Helper Bundle - v1.2
' req: setGlobalsIfNecessary()
' ----------------------------------------------------
Function AccessClipboard(copyText)
	Dim regPath, keyName, objIE
    Call setGlobalsIfNecessary("objShell_WS")
	
	regPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\"
	keyName = "1407"
	objShell_WS.RegWrite RegPath & keyName, 0, "REG_DWORD"	
	 
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Navigate("about:blank")
    
    If Not IsNull(copyText) Then
        Call objIE.document.parentwindow.clipboardData.SetData("text", copyText) ' copy to Clipboard
        result = copyText
    Else
        result = objIE.document.parentwindow.clipboardData.GetData("text")       ' paste from Clipboard
    End If
    objIE.Quit
	'objShell_WS.RegWrite RegPath & keyName, 1, "REG_DWORD"
	Set objIE = Nothing
    AccessClipboard = result
End Function

Function GetClipboardData()
    GetClipboardData = AccessClipboard(Null)
End Function

Sub SetClipboardData(copyText)
    GetClipboardData = AccessClipboard(copyText)
End Sub
' ----------------------------------------------------


' Helper Bundle - v1.8
' req: setGlobalsIfNecessary()
'      Const C_Encoding_Default = -2
'      Const C_Encoding_Unicode = -1
'      Const C_Encoding_ASCII = 0
' ----------------------------------------------------
Function read_File(ByRef strPath, ByRef encoding)
    functionName = "read_File"
	Call setGlobalsIfNecessary("objFSO")
    If objFSO.FileExists(strPath) Then
        On Error Resume Next
            Set objFile = objFSO.GetFile(strPath)
            
            If Err.Number = 0 Then
                Set objTSO = objFile.OpenAsTextStream(1, encoding)
                read_File = objTSO.Read(objFile.Size)
                Set objTSO = Nothing
            Else
                Err.Clear
                read_File = ""
            End If
            
            Set objFile = Nothing
        On Error GoTo 0
    Else
        Call show_MsgBox("File not found: " & strPath, vbCritical, "Function: " & functionName)
    End If
End Function

' Helper - v1.3
Function get_TextLines(ByRef strPath, ByRef encoding)
    fileText = read_File(strPath, encoding)
	If fileText <> "" Then
		get_TextLines = Split(fileText, vbCrlf)
	Else
		get_TextLines =  Array()
	End If
End Function

' Helper - v1.4
' req: setGlobalsIfNecessary()
Sub write_TextToFile(ByRef strPath, ByRef strText, ByRef IOMode, ByRef encoding)
	Call setGlobalsIfNecessary("objFSO")
	If IOMode = 8 AND objFSO.FileExists(strPath) Then strText = vbCrLf & strText  	' for appending
	Set objFile = objFSO.OpenTextFile(strPath, IOMode, True, encoding)	' Create a new File, if not found
	objFile.Write strText
	objFile.Close
	Set objFile = Nothing
End Sub

' Helper - v1.3
' replaceCount = -1 for all occurrences
' compareMethod: vbTextCompare (case insensitive)
'                vbBinaryCompare (case sensitive)
Sub replace_inFile(ByRef strPath, ByRef encoding, ByRef strFind, ByRef strReplace, ByRef replaceCount, ByRef compareMethod)
    text = read_File(strPath, encoding)   
    newText = Replace(text, strFind, strReplace, 1, replaceCount, compareMethod)    
    Call write_TextToFile(strPath, newText, 2, encoding)  
End Sub

' Helper - v1.8
' accepts an array of values or an array of multiple arrays with values
Sub write_ArrayToFile(ByRef strPath, ByRef arr, ByRef IOMode, ByRef encoding)
    If isArray(arr) Then
        isArrOfArrs = False
        On Error Resume Next
            If isArray(arr(0)) Then isArrOfArrs = True
        On Error goto 0
        
        If isArrOfArrs Then
            strOutput = ""
            For Each singleArr In arr
                strOutput = strOutput & Join(singleArr, vbCrLf) & vbCrLf
            Next
        Else
            strOutput = Join(arr, vbCrLf)
        End If
        Call write_TextToFile(strPath, strOutput, IOMOde, encoding)  
    Else
        MsgBox "Variable 'arr' is no Array: " & TypeName(arr)
    End If
End Sub

' Helper - v1.2
Sub show_MsgBox(ByRef msg_ArrOrStr, ByRef msgType, ByRef strTitle)
    strSpace = "   "    ' space for text in msgBoxes
    If IsArray(msg_ArrOrStr) Then
        strMsg = Join(msg_ArrOrStr, vbCrlf & strSpace)
    Else
        strMsg = msg_ArrOrStr
    End If
	MsgBox strSpace & strMsg, msgType + vbMsgBoxSetForeground, strTitle
End Sub
' ----------------------------------------------------


' Helper Bundle - v2.0
' compareMethod: vbTextCompare (case insensitive)
'                vbBinaryCompare (case sensitive)
' cutChars_Begin + cutChars_End can be an Integer, a String or an Array of Strings
' With an array of possibilities the order does matter when the possibilities contain same chars
' ----------------------------------------------------
Function cutBeginEnd(ByRef str, ByRef cutChars_Begin, ByRef cutChars_End, ByRef compareMethod)
    isError = False
    cutBegin_Type = TypeName(cutChars_Begin)
    Select Case cutBegin_Type
        Case "Integer"
            strNew = Right(str, Len(str)-cutChars_Begin)
        Case "String"
            strNew = Replace(str, cutChars_Begin, "", 1, 1, compareMethod)       
        Case "Variant()"
            strNew = str    ' for the case that no possibility matches
            For Each cut_Begin In cutChars_Begin
                If StartsWith(str, cut_Begin) Then
                    strNew = Replace(str, cut_Begin, "", 1, 1, compareMethod)                 
                    Exit For
                End If
            Next        
        Case Else
            isError = True
    End Select
    
    If Not isError Then
        cutEnd_Type = TypeName(cutChars_End)
        Select Case cutEnd_Type
            Case "Integer"
                strNew = Left(strNew, Len(strNew)-cutChars_End)            
            Case "String"
                index_cutChars_End = InStrRev(strNew, cutChars_End)
                If index_cutChars_End > 0 Then
                    strNew = Mid(strNew, 1, index_cutChars_End - 1) & Replace(strNew, cutChars_End, "", index_cutChars_End, 1, compareMethod)
                End If            
            Case "Variant()"
                For Each cut_End In cutChars_End
                    If EndsWith(str, cut_End) Then
                        index_cut_End = InStrRev(strNew, cut_End)
                        If index_cut_End > 0 Then
                            strNew = Mid(strNew, 1, index_cut_End - 1) & Replace(strNew, cut_End, "", index_cut_End, 1, compareMethod)
                        End If                 
                        Exit For
                    End If
                Next            
            Case Else
                isError = True
        End Select
    End If
    If Not isError Then
        cutBeginEnd = strNew
    Else 
        cutBeginEnd = ""
        MsgBox "Error: one cutType not supported - cutBegin_Type=" & cutBegin_Type & " cutEnd_Type=" & cutEnd_Type
    End If
End Function

' v1.3
Function StartsWith(ByRef str, ByRef start)
	StartsWith = (Left(Trim(LCase(str)), Len(start)) = LCase(start))
End Function

' v1.4
Function EndsWith(ByRef str, ByRef ending)
	EndsWith = (Right(Trim(LCase(str)), Len(ending)) = LCase(ending))
End Function
' ----------------------------------------------------


' Helper Bundle - v2.5
' under construction
' ----------------------------------------------------
Function isRun_program(ByRef strCmd)
	Set colProcesses = get_runningProcesses(strCmd)
    isRun_ProgramWithArgs = colProcesses.Count <> 0
	Set colProcesses = Nothing
End Function

' wenn cmd dann ganzen cmd checken
' erster teil bis leerzeichen ohne anfuehrungsstriche muss zu path passen
' kann aber auch leer sein, dann nur nameOrExecutablePath

' Also: 
' wenn nur name: nur name suchen
' wenn path, dann imagepath checken
' wenn cmd dann nur cmd checken

' wenn beginn mit " dann nach " mit leerzeichen dahinter suchen,
' sonst nur nach erstem lerzeichen

' ^"?*^\"?$

' "dgd\gd" = path
' "dggdgd" = name
' dgegd = name
' dgdgd\dgdg = path

' fhf -fhf = cmd
' hffh\fhfh -fghf = cmd
' "fghfh" -fhf = cmd
' "fhf\fhdh" -dfghdg = cmd

Function get_runningProcesses(ByRef strCmd)
	Set objWMI = GetObject("winmgmts:")
' Baustelle / toDo   
    ' remove quotes from program path at the commands beginning
    If StartsWith(strCmd, """") Then
        'indexClosingQt = InStr(2, strCmd, """", vbBinaryCompare)
        
        wenn folgender string (" ) ist und danach noch mehr kommt
        arrCmd = Split(strCmd, """", 3)
        MsgBox arrCmd(0) & vbcrlf & arrCmd(1) & vbcrlf & arrCmd(2)
        

        If UBound(arrCmd) > 0 Then
            targetField = "CommandLine"
        ElseIf Contains(strCmd, "\") Then
            targetField = "ExecutablePath"
            path_program_esc = Replace(strCmd, "\", "\\")
        Else
            targetField = "Name"
            path_program_esc = strCmd    
        End If
        
        path_program = SubString(strCmd, 1, indexClosingQt)
        MsgBox path_program
        If Contains(path_program, "\") Then
            targetField = "ExecutablePath"
            path_program_esc = Replace(path_program, "\", "\\")
        Else
            targetField = "Name"
            path_program_esc = path_program
        End If
        
        'tempCmd = Replace(strCmd, """", "", 1, 2) 
    Else
        arrCmd = Split(strCmd)
        If UBound(arrCmd) > 0 Then
            targetField = "CommandLine"
        ElseIf Contains(strCmd, "\") Then
            targetField = "ExecutablePath"
            path_program_esc = Replace(strCmd, "\", "\\")
        Else
            targetField = "Name"
            path_program_esc = strCmd    
        End If
    End If
    
    If args = "" Then
        sqlQuery = "SELECT * FROM win32_process WHERE " & targetField & "='" & path_program_esc & "'"
    Else
        ' you have to escape the backslashes for wmi and the special characters for the like statement
        args_likeEsc = replaceMultiple(args, Array(Array("\", "\\"), _
                                                   Array("[", "[[]"), _
                                                   Array("%", "[%]"), _
                                                   Array("_", "[_]")))
        sqlQuery = "SELECT * FROM win32_process WHERE " & targetField & "='" & path_program_esc & "'" & _
                   " and CommandLine LIKE '%" & args_likeEsc & "%'"
    End If
    Set get_runningProcesses = objWMI.ExecQuery(sqlQuery)
    Set objWMI = Nothing
End Function

Function isRun_Process(ByRef processID)
	Set objWMI = GetObject("winmgmts:")
    sqlQuery = "SELECT * FROM win32_process WHERE ProcessId='" & processID & "'"
    Set colProcesses = objWMI.ExecQuery(sqlQuery)
    isRun_Process = colProcesses.Count <> 0
    Set objWMI = Nothing
End Function

Sub exit_program(ByRef strCmd)
    Set colProcesses = get_runningProcesses(strCmd)
    For Each objProcess In colProcesses
        processID = objProcess.processID
        objProcess.Terminate()
        Call waitForState_process(processID, False)
    Next
    Set colProcesses = Nothing
End Sub

Function waitForState_program(ByRef strCmd, ByRef desiredState)
	timeOut_s = 0
    reCheckInterval_ms = 200
    isDesiredState = False
    isTimeOut = False
    If timeOut_s > 0 Then startTime = Now()
	Do Until (isDesiredState Or isTimeOut)
        isTimeOut = (timeOut_s > 0 And DateDiff("s", startTime, Now()) > timeOut_s)
        If Not isTimeOut Then
            isDesiredState = (isRun_programWithArgs(path_program, args) = desiredState)
            If Not isDesiredState Then WScript.Sleep reCheckInterval_ms
        End If
	Loop
    waitForState_program = isDesiredState
End Function

Function waitForState_process(ByRef processID, ByRef desiredState)
	timeOut_s = 0
    reCheckInterval_ms = 200
    isDesiredState = False
    isTimeOut = False
    If timeOut_s > 0 Then startTime = Now()
	Do Until (isDesiredState Or isTimeOut)
        isTimeOut = (timeOut_s > 0 And DateDiff("s", startTime, Now()) > timeOut_s)
        If Not isTimeOut Then
            isDesiredState = (isRun_Process(processID) = desiredState)
            If Not isDesiredState Then WScript.Sleep reCheckInterval_ms
        End If
	Loop
    waitForState_process = isDesiredState
End Function

' v1.0
Function replaceMultiple(ByRef str, ByRef arrReplaces)
    For Each arrReplace In arrReplaces
        str = Replace(str, arrReplace(0), arrReplace(1))
    Next
    replaceMultiple = str
End Function

' v1.3
Function Contains(ByRef str, ByRef strSearch)
	' converting to lower case is better than vbTextCompare because of dealing with foreign languages
	If InStr(LCase(str), LCase(strSearch)) > 0 Then returnValue = True Else returnValue = False
	Contains = returnValue
End Function

' v1.3
Function IndexOf(ByRef str, ByRef strSearch)
	IndexOf = InStr(LCase(str), LCase(strSearch))
End Function

' v1.3
Function StartsWith(ByRef str, ByRef start)
	StartsWith = (Left(Trim(LCase(str)), Len(start)) = LCase(start))
End Function

' v1.1
Function SubString(ByRef str, ByRef startIndex, ByRef endIndex)
	SubString = Mid(str, startIndex, endIndex - startIndex)
End Function
' ----------------------------------------------------


' Helper - v1.0
' under Construction
' req: setGlobalsIfNecessary() | isRun_program()
Function waitForCmd_driveMount(ByRef strCmd, ByRef path_verifyMarker, ByRef desiredState, ByRef timeOut_s)
	Call setGlobalsIfNecessary("objFSO")
    reCheckInterval_ms = 200
    isCmdExit = False
    isMount = False
    isTimeOut = False
    If timeOut_s > 0 Then startTime = Now()
	Do Until (isCmdExit Or isMount Or isTimeOut)
        isTimeOut = (timeOut_s > 0 And DateDiff("s", startTime, Now()) > timeOut_s)
        If Not isTimeOut Then
            isMount = (objFSO.FileExists(path_verifyMarker) = desiredState)
            If Not isMount Then
   ' toDo: add isRun_programWithArgs, is more precise, in case similar cmd was already running
                isRun_programWithArgs(ByRef path_program, ByRef args)
                isCmdExit = Not isRun_program(strCmd)
                If Not isCmdExit Then WScript.Sleep reCheckInterval_ms
            End If
        End If
	Loop
    waitForCmd_driveMount = isMount
End Function


' Helper - v1.3
' req: setGlobalsIfNecessary()
Function waitForState_driveMount(ByRef path_verifyMarker, ByRef desiredState, ByRef timeOut_s)
	Call setGlobalsIfNecessary("objFSO")
    reCheckInterval_ms = 200
    isMount = False
    isTimeOut = False
    If timeOut_s > 0 Then startTime = Now()
	Do Until (isMount Or isTimeOut)
        isTimeOut = (timeOut_s > 0 And DateDiff("s", startTime, Now()) > timeOut_s)
        If Not isTimeOut Then
            isMount = (objFSO.FileExists(path_verifyMarker) = desiredState)
            If Not isMount Then WScript.Sleep reCheckInterval_ms
        End If
	Loop
    waitForState_driveMount = isMount
End Function


' Helper Bundle - v1.1
' ----------------------------------------------------
' v1.3
Function insert_DriveLetter(ByRef path)
    arrPathParts = Split(path, "\", 2)
    placeHolder = arrPathParts(0)
    If Len(placeHolder) > 2 Then
        driveName = cutBeginEnd_String(placeHolder, "[", "]", vbTextCompare)
        letter = get_DriveLetterFromLabel(driveName)
        If letter <> "" Then
            returnString = letter & "\" & arrPathParts(1)
        Else
            returnString = path
        End If
    Else
       returnString = path
    End If
    insert_DriveLetter = returnString
End Function

' v1.0
Function get_DriveLetterFromLabel(ByRef driveName)
    ' Ask WMI for the list of volumes with the requested label
    Set volumes = GetObject("winmgmts:") _ 
                  .ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE Label='" & driveName & "'")
    ' If exist an matching volume, get its drive letter
    If volumes.Count > 0 Then 
        For Each volume In volumes 
            result = volume.DriveLetter
            Exit For
        Next
    Else
        result = ""
    End If
    get_DriveLetterFromLabel = result
End Function

' v1.1
' compareMethod: vbTextCompare (case insensitive)
'                vbBinaryCompare (case sensitive)
Function cutBeginEnd_String(ByRef str, ByRef cutChars_Begin, ByRef cutChars_End, ByRef compareMethod)
    strNew = Replace(str, cutChars_Begin, "", 1, 1, compareMethod)       
    index_cutChars_End = InStrRev(strNew, cutChars_End)
    If index_cutChars_End > 0 Then
        strNew = Mid(strNew, 1, index_cutChars_End - 1) & Replace(strNew, cutChars_End, "", index_cutChars_End, 1, compareMethod)
    End If
    cutBeginEnd_String = strNew 
End Function
' ----------------------------------------------------


' Helper - v1.1
Function arr_SafeGet(ByRef arr, ByRef index, ByRef defaultValue)
    If UBound(arr) >= index Then arr_SafeGet = arr(index) Else arr_SafeGet = defaultValue
End Function


' Helper - v1.0	
Function getOSArchitecture()
	strOSArchitecture = GetObject("winmgmts:root\cimv2:Win32_OperatingSystem=@").OSArchitecture
	If strOSArchitecture = "64-Bit" Then getOSArchitecture = "x64" Else getOSArchitecture = "x86"
End Function


' Helper - v1.1
' always case insensitive
' FirstIndex works only for the whole string, not for each submatch
Function indexOf_RegEx(ByRef strText, ByRef find_pattern)
	Set objRegEx = New RegExp
	With objRegEx
		.Pattern = find_pattern
		.Global = False
		.IgnoreCase = True
	End With
    Set matches = objRegEx.Execute(strText)
    If matches.Count > 0 Then
        indexOf_RegEx = matches(0).FirstIndex
    Else
        indexOf_RegEx = -1
    End If
	Set objRegEx = Nothing
End Function


' Helper Bundle - v1.1
' ----------------------------------------------------
' always case insensitive
Function contains_RegEx(ByRef strText, ByRef find_pattern)
	Set objRegEx = New RegExp
	With objRegEx
		.Pattern = find_pattern
		.Global = False
		.IgnoreCase = True
	End With
	contains_RegEx = objRegEx.Test(strText) 
	Set objRegEx = Nothing
End Function

' always case insensitive
Function contains_File_RegEx(ByRef strPath, ByRef encoding, ByRef find_pattern)
    strText = read_File(strPath, encoding)
	contains_File_RegEx = contains_RegEx(strText, find_pattern)
End Function
' ----------------------------------------------------

 
' Helper Bundle - v1.1
' ----------------------------------------------------
' replaces all occurrences
Function replace_RegEx(ByRef strText, ByRef find_pattern, ByRef replace_pattern, ByRef ignoreCase)
	Set objRegEx = New RegExp
	With objRegEx
		.Pattern = find_pattern
		.Global = True
		.IgnoreCase = ignoreCase
	End With
	replace_RegEx = objRegEx.Replace(strText, replace_pattern)
	Set objRegEx = Nothing
End Function

' replaces all occurrences
Sub replace_inFile_RegEx(ByRef strPath, ByRef encoding, ByRef find_pattern, ByRef replace_pattern, ByRef ignoreCase)
    strText = read_File(strPath, encoding)
    newText = replace_RegEx(strText, find_pattern, replace_pattern, ignoreCase)
    Call write_TextToFile(strPath, newText, 2, encoding)
End Sub
' ----------------------------------------------------

' Helper - v1.2
' replaces all occurrences
Sub replace_inFile(ByRef strPath, ByRef encoding, ByRef strFind, ByRef strReplace, ByRef ignoreCase)
    Select Case ignoreCase
        Case True
            compareMethod = vbTextCompare
        Case False
            compareMethod = vbBinaryCompare
    End Select
    text = read_File(strPath, encoding)
    replaceCount = -1   ' replace all occurrences
    newText = Replace(text, strFind, strReplace, 1, replaceCount, compareMethod)    
    Call write_TextToFile(strPath, newText, 2, encoding)  
End Sub
 
 
' Helper - v1.2
' req: setGlobalsIfNecessary()
Function IsEmpty_Folder(ByRef path_folder)
    Call setGlobalsIfNecessary("objFSO")
    Set folder = objFSO.getFolder(path_folder)
    If folder.Files.Count = 0 And folder.SubFolders.Count = 0 Then
        returnValue = True
    Else
        returnValue = False
    End If
    IsEmpty_Folder = returnValue
End Function


' Helper - v1.0
' req: setGlobalsIfNecessary()
Sub sendPW_runAs(ByRef password)
    Call setGlobalsIfNecessary("objShell_WS")
    WScript.Sleep 100
    For k=1 To Len(password)
        For i = 0 to 10
            Activated = objShell_WS.AppActivate ("runas")
            If Activated = 0 Then
                objShell_WS.SendKeys Mid(password,k,1)
                WScript.Sleep 10
                Exit For
            End If
        Next
    Next
    
    For i = 0 to 10
        Activated = objShell_WS.AppActivate ("runas")
        If Activated = 0 Then
            objShell_WS.SendKeys "{ENTER}"
            Exit For
        End If
    Next
End Sub


' Helper - v1.0
Function open_OutlookContact(ByRef contact_entryID)
    Set Outlook = CreateObject("Outlook.Application")
    Set myNameSpace = Outlook.GetNamespace("MAPI")
    On Error Resume Next
        myNameSpace.GetItemFromID(contact_entryID).Display
        If Err.Number <> 0 Then
            Err.Clear
            msgType = vbExclamation
            MsgBox C_strSpace & "Outlook Contact not found !", vbMsgBoxSetForeground + msgType, "Outlook_Calls"
        End If
    On Error GoTo 0
End Function


' v1.2
Sub show_MsgBox(ByRef msg_ArrOrStr, ByRef msgType, ByRef strTitle)
    strSpace = "   "    ' space for text in msgBoxes
    If IsArray(msg_ArrOrStr) Then
        strMsg = Join(msg_ArrOrStr, vbCrlf & strSpace)
    Else
        strMsg = msg_ArrOrStr
    End If
	MsgBox strSpace & strMsg, msgType + vbMsgBoxSetForeground, strTitle
End Sub


' Helper
' req: setGlobalsIfNecessary()
	Call setGlobalsIfNecessary("objFSO")
    IsFile = objFSO.FileExists(sPath)


' Helper - v1.1
' Powershell:   -ArgumentList of start-process expects the format:
'               'arg1', 'arg2', 'arg3'
Function create_ArgumentStr_PS(ByRef listArgs_1, ByRef listArgs_2) 
    If listArgs_1.Count + listArgs_2.Count > 0 Then
        returnString = "'"
        If listArgs_1.Count > 0 Then returnString = returnString & Join(listArgs_1.toArray, "', '")
        If listArgs_2.Count > 0 Then returnString = returnString & "', '" & Join(listArgs_2.toArray, "', '")
        returnString = returnString & "'"
    Else
        returnString = ""
    End If
    create_ArgumentStr_PS = returnString
End Function



' Helper - v1.0
' Powershell:   Escaping every whitespace offers more compability
'               than escaping the whole string with `'my_str`'
Function esc_spaces_PS(ByRef strValue)
    esc_spaces_PS = Replace(strValue, " ", "` ")
End Function

' Helper - v1.0
Function create_ArgumentStr_PS(ByRef listArgs)
    create_ArgumentStr_PS = "'" & Join(listArgs.toArray, "', '") & "'"
End Function

' Helper - v1.2
' Powershell:	Expressions in single-quoted strings are not evaluated
'				In double-quoted strings any variable names such as "$myVar" will be replaced with the variable's value
Function qt_s(ByRef strValue)
    qt_s = Chr(39) & strValue & Chr(39)
End Function

	
' Helper - v1.2
Function qt(ByRef strValue)
    qt = Chr(34) & strValue & Chr(34)
End Function

' Helper - v1.4
' runas needs escaped doublequotes
Function qt_esc_runAs(ByRef strValue)
    escapedQuote = "\" & Chr(34)
    qt_esc_runAs = escapedQuote & strValue & escapedQuote
End Function

' Helper - v1.1
' PowerShell escaped singlequotes
Function qt_esc_PS(ByRef strValue)
    escapedQuote = "`" & Chr(39)
    qt_esc_PS = escapedQuote & strValue & escapedQuote
End Function

' Helper - v1.2
' Powershell:   In double-quoted strings any variable names such as "$myVar"
'               will be replaced with the variable's value
Function qt_d(ByRef strValue)
    doubleQuotes = Chr(34) & Chr(34)
    qt_d = doubleQuotes & strValue & doubleQuotes
End Function


' Helper - v1.2
' req: setGlobalsIfNecessary()
' saves only unique items !
Sub addCSV_toDictOrList(ByRef dictOrList, ByRef strFolder, ByRef strFile, ByRef encoding)
	Call setGlobalsIfNecessary("objFSO")
    dictOrList_Type = TypeName(dictOrList)
    path_File = objFSO.BuildPath(strFolder, strFile)
    arrLines = get_TextLines(path_File, encoding)
	If Not isEmpty_ArrList(arrLines) Then
        If dictOrList_Type = "Dictionary" Then
            For Each strLine In arrLines
                strLine = Trim(strLine)
                If strLine <> "" Then
                    If Not dictOrList.Exists(strLine) Then
                        dictOrList.Add strLine, ""
                    End If
                End If
            Next
        ElseIf dictOrList_Type = "ArrayList" Then
            For Each strLine In arrLines
                strLine = Trim(strLine)
                If strLine <> "" Then dictOrList.Add strLine
            Next
        Else
            MsgBox "Input is no dictionary or list"
        End If
	End If
End Sub

' Helper - v1.0
' req: setGlobalsIfNecessary()
' saves only unique items !
Function get_DictFromCSV(ByRef strFolder, ByRef strFile)
	Call setGlobalsIfNecessary("objFSO")
    Set dictFromCSV = CreateObject("Scripting.Dictionary")
    filePath = objFSO.BuildPath(strFolder, strFile)
    arrLines = get_TextLines(filePath, C_Encoding_ASCII)			
	For Each strLine In arrLines
    	strLine = Trim(strLine)
		If strLine <> "" Then
            strUser = cutBeginEnd(strLine, """", """,""""", vbBinaryCompare)
            If Not dictFromCSV.Exists(strUser) Then            
                dictFromCSV.Add strUser, ""
            End If
		End If
	Next
    Set get_DictFromCSV = dictFromCSV
End Sub


' Helper Bundle - v1.0
' String operations
' ----------------------------------------------------
' Helper - v1.0
' compareMethod: vbTextCompare (case insensitive)
'                vbBinaryCompare (case sensitive)
Function cutBeginEnd_String(ByRef str, ByRef cutChars_Begin, ByRef cutChars_End, ByRef compareMethod)
    strNew = Replace(str, cutChars_Begin, "", 1, 1, compareMethod)       
    index_cutChars_End = InStrRev(strNew, cutChars_End)
    If index_cutChars_End > 0 Then
        strNew = Mid(strNew, 1, index_cutChars_End - 1) & Replace(strNew, cutChars_End, "", index_cutChars_End, 1, compareMethod)
    End If
End Function

' Helper - v1.0
Function replaceMultiple(ByRef str, ByRef arrReplaces)
    For Each arrReplace In arrReplaces
        str = Replace(str, arrReplace(0), arrReplace(1))
    Next
    replaceMultiple = str
End Function

' Helper - v1.1
Function SubString(ByRef str, ByRef startIndex, ByRef endIndex)
	SubString = Mid(str, startIndex, endIndex - startIndex)
End Function

' Helper - v1.2
' Trim for a broader range of whitespaces
Function ExtTrim(ByRef str)
	Set objRegEx = New RegExp
	With objRegEx
		.Pattern = "^[\s\xA0]+|[\s\xA0]+$"
		.Global = True
		.IgnoreCase = True
	End With
	ExtTrim = objRegEx.Replace(str, "")
	Set objRegEx = Nothing
End Function

' Helper - v1.3
Function Contains(ByRef str, ByRef strSearch)
	' converting to lower case is better than vbTextCompare because of dealing with foreign languages
	If InStr(LCase(str), LCase(strSearch)) > 0 Then returnValue = True Else returnValue = False
	Contains = returnValue
End Function

' Helper - v1.3
Function IndexOf(ByRef str, ByRef strSearch)
	IndexOf = InStr(LCase(str), LCase(strSearch))
End Function

' Helper - v1.3
Function StartsWith(ByRef str, ByRef start)
	StartsWith = (Left(Trim(LCase(str)), Len(start)) = LCase(start))
End Function

' Helper - v1.4
Function EndsWith(ByRef str, ByRef ending)
	EndsWith = (Right(Trim(LCase(str)), Len(ending)) = LCase(ending))
End Function

' Helper - v1.3
Function enclosedWith(ByRef str, ByRef strStart, ByRef strEnd)
    enclosedWith = ((Left(Trim(LCase(str)), Len(strStart)) = LCase(strStart)) _ 
                    And (Right(Trim(LCase(str)), Len(strEnd)) = LCase(strEnd)))
End Function
' ----------------------------------------------------


' Helper - v1.1
Function validateURL(ByRef strURL)
	Dim reValid 
	Set reValid = New RegExp 
	reValid.Pattern = "^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$" 
	reValid.MultiLine = False 
	reValid.Global = True 
	validateURL = reValid.Test(strURL) 
End Function


' Helper - v1.0
Function createSpace(ByRef length)
	strSpace = ""
	For i = 1 To length
		strSpace = strSpace & " "
	Next
	createSpace= strSpace
End Function


' Helper - v1.1
Function getShuffledArray(ByVal arr)
    Dim i, j, temp
    Randomize
    UBarr = UBound(arr)
    For i = 0 To UBarr
        j = CLng(((UBarr - i) * Rnd) + i)
        If i <> j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
        End If
    Next
    getShuffledArray = arr
End Function