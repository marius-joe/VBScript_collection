' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Arg Runner Util
'  v1.6.3
' ******************************************
'  Arguments:  -file "path to file"
'              (opt) -args ["arg1" "arg2"]
'              (opt) -admin
'              (opt) -nowindow 
'              (opt) -noexit          
'              (opt. arguments order doesn't matter)

'  Path options:  this\local file path      relative to this vbs file
'                 parent\local file path    relative to this vbs file
'                 absolute path
' ******************************************

'toDo
' Pfad richtig übergeben, nicht relativ zum Arg Runner
' alle formate testen
' wait Funktionalität testen 
' noch zwischen Powershell und File Args unterscheiden und testen 

' arraylists durch dynmisch anwachsene arrays ersetzen wegen .NET Unabhängigkeit

Const C_isDebugMode = False     ' turns on debug mode

Const C_strSpace = "   "    ' space for text in msgBoxes

Dim objFSO							 					  
Dim objShell_A
Dim objShell_WS


    If C_isDebugMode Then runWindowVisibility = 1 Else runWindowVisibility = 0
    isArgError = False
	arrComplexArgs = getComplexArguments()
	If Not isEmpty_ArrList(arrComplexArgs) Then
 ' namen überarbeiten, muss auf alle fileTypes passen
        Set listArgs_Program = CreateObject("System.collections.arraylist")	    ' Program parameters
        Set listArgs_Script = CreateObject("System.collections.arraylist")		' Script specific parameters
        path_File = ""
        strUser = ""
        password = ""
        asAdmin = False
		waitForProgram = False  
		For Each complexArg In arrComplexArgs
			argument = LCase(complexArg(0))
            value = arr_SafeGet(complexArg, 1, "")
            
			Select Case argument
				Case "-file"
					If value <> "" Then path_File = value Else isArgError = True
                    
				Case "-nowindow"
					' noch einbauen, geht das überhaupt, wenn auch mehrere Dateitypen unterstützt werden ??
                    
				Case "-wait"
					waitForProgram = True
                    
				Case "-admin"
					asAdmin = True
                     
				Case "-noexit"
					listArgs_Program.Add("-noexit")
                     
				Case "-args"
                    If enclosedWith(value, "[", "]") Then listArgs_Script.Add(cutBeginEnd(value, 1, 1, vbTextCompare)) Else isArgError = True
                   
				Case "-user"
					If value <> "" Then strUser = value Else isArgError = True
                    
				Case "-pw"
					If value <> "" Then password = value Else isArgError = True
			End Select
		Next
    Else
        isArgError = True
    End If

    If Not isArgError Then
        If EndsWith(path_File, ".vbs") OR EndsWith(path_File, ".js") Then
            fileType = "script"
        ElseIf EndsWith(path_File, ".ps1") Then
            fileType = "ps1"
        ElseIf EndsWith(path_File, ".msc") Then
            fileType = "msc"
        ElseIf EndsWith(path_File, ".exe") Then
            fileType = "exe"
        Else
            isNoValidType = True
        End If
        
   ' man könnte auch nur -cmd "" angeben dann wird direkt mit "cmd.exe" der Befehl ausgeführt, vbs, js und powershell befehle genau so
    
        If Not isNoValidType Then
            Call setGlobalsIfNecessary("objFSO")
            isKnownProgram = False
            If StartsWith(path_File, "this\") Then
                strFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
                path_File_Abs = objFSO.BuildPath(strFolder, cutBeginEnd(path_File, "this\", 0, vbTextCompare))
                
            ElseIf StartsWith(path_File, "parent\") Then
                strFolder = objFSO.GetParentFolderName(objFSO.GetParentFolderName(WScript.ScriptFullName))        
                path_File_Abs = objFSO.BuildPath(strFolder, cutBeginEnd(path_File, "parent\", 0, vbTextCompare))               
            Else
                path_File_Abs = path_File
                If (fileType = "exe" OR fileType = "msc") AND (Not Contains(path_File, "\")) Then isKnownProgram = True
            End If

            If objFSO.FileExists(path_File_Abs) OR isKnownProgram Then    ' so muss für Systemprogramme, z.B. cmd.exe nicht der komplette Pfad angegeben werden, ShellExecute sucht dann selber
                Select Case fileType
                    Case "script"
                        program = "wscript.exe"
                        listArgs_Program.Add(esc_spaces_PS(path_File_Abs))
                        
' wird Escape von Leerzeichen nur bei Powershell gebraucht ???
                    Case "ps1"
                        program = "powershell.exe"
                        listArgs_Program.Add("-ExecutionPolicy Bypass")
                        listArgs_Program.Add(esc_spaces_PS(path_File_Abs))
                        
                    Case "msc"
                        program = "mmc.exe"
                        listArgs_Program.Add(esc_spaces_PS(path_File_Abs))
                        
                    Case "exe"
                        program = path_File_Abs
                        
                End Select
                
                
                
                If listArgs_Program.Count > 0 Then
                    strArgs_Program = "'" & Join(listArgs_Program.toArray, "', '") & "', " & strArgs_Program					
                End If
                If listArgs_Script.Count > 0 Then						
                    strArgs_Script = ", " & "'" & Join(listArgs_Script.toArray, "', '") & "'"				
                End If
                    
        ' Code hier, wenn kein User für admin Zugriff übergeben wurde
        ' was ist mit Whitespaces in Args/Pfaden, die für das aufgerufene Script übergeben werden
                If strUser = "" Then
' noch checken ob so richtig      

                    strArgs_cmdlet = create_ArgumentStr_PS(listArgs_Program, listArgs_Script)

                    If asAdmin Then
                        userRights = " -Verb runAs"
                    Else 
                        userRights = ""
                        optCmdArguments2 = " -NoNewWindow"       ' -NoNewWindow kann nicht zusammen mit  " -Verb runAs"  benutzt werden. 
                    End If                              
                  
  ' hier muss wohl noch zwischen Powershell und File Args unterschieden werden'           
                    If strArgs_cmdlet <> "" Then
                        optCmdArguments1 = "$arr_arguments = " & strArgs_cmdlet & " ; "
                        optCmdArguments2 = optCmdArguments2 & " -ArgumentList $arr_arguments"
                    End If                    
                    

                    If Not C_isDebugMode Then
                        ' the script calls Powershell, Powershell then calls the program needed to execute the file as admin - has to be done this way, to support "wait" functionality + the correct UAC Message which .ShellExecute does not
                        strCmd = "powershell " & "-Command " & qt(optCmdArguments1 & "start-process " & qt_s(program) & userRights & " -Wait" & optCmdArguments2)
                
                        Call setGlobalsIfNecessary("objShell_WS")
                        objShell_WS.Run strCmd, runWindowVisibility, false
                    End If
                    
                    
                    
        ' Fall mit User + Passwort erst später fertig machen            
                Else
                    If asAdmin Then userRights = " -Verb runAs" Else userRights = ""
                    strArguments = strArgs_Program & strArgs_Script
                    If strArguments <> "" Then
                        optCmdArguments1 = "$arr_arguments = " & strArguments & " ; "
                        optCmdArguments2 = " -ArgumentList $arr_arguments"
                    End If
                    ' the script calls Powershell as a user with the needed rights, Powershell then calls the program needed to execute the file as admin -> no easier way with "objShell_A.Run "runas /user""
                    ' we have to replace whitespaces in the path because you cannot use doublequotes in here cause of the first parsing with runas
                    strCmd = "powershell " & "-Command " & qt_esc_runAs(optCmdArguments1 & "start-process " & qt_s(program) & userRights & " -Wait -NoNewWindow" & optCmdArguments2)  'wait + nonewwindows sind neu, noch testen

                    If Not C_isDebugMode Then
                        Call setGlobalsIfNecessary("objShell_WS")
                        objShell_WS.Run "runas /user:" & qt(strUser) & " " & qt(strCmd), runWindowVisibility, false		' don't wait until script has finished, Window has to be shown for the password entry
                        
                        If password <> "" Then
                            Call sendPW_runAs(password)
                        End If
                        Call cleanGlobals("objShell_WS")
                    End If
                End If
                    
                If C_isDebugMode Then
                    If asAdmin Then strRights = "Admin" Else strRights = "No Admin"
                    optLoginInfo = strUser & vbCrlf & password
                    arrMsg = Array("Program :  " & program, _
                                   "Command :  " & strCmd, _
                                   "User :  " & strRights, _
                                   "Login: " & vbCrlf & optLoginInfo, _
                                   "listArgs_Program :" & vbCrlf & Replace(strArgs_Program, " -", vbCrlf & "-"), _
                                   "listArgs_Script :" & vbCrlf & Replace(strArgs_Script, " -", vbCrlf & "-"))
                    Call show_MsgBox(arrMsg, vbInformation, "Arg Runner")
                End If
                
            Else
                arrMsg = Array("Pfad nicht gefunden und Programm nicht bekannt !", _
                               "", _
                               path_File_Abs)
                msgType = vbCritical
                strTitle = "Run: " & qt(path_File_Abs)
                Call show_MsgBox(arrMsg, msgType, strTitle)            
            End If
            
            Call cleanGlobals("objFSO")
        Else
            msgType = vbCritical
            msgTitle = "Run: " & qt(strCmd)
            strMessage = "Dateityp wird nicht unterstuetzt (nur .exe .vbs .js .msc) !"

            MsgBox C_strSpace & strMessage, vbMsgBoxSetForeground + msgType, msgTitle
        End If
    Else
		msgType = vbCritical
		msgTitle = "Arg Runner"
		strMessage = "Argumente fehlerhaft !"

		MsgBox C_strSpace & strMessage, vbMsgBoxSetForeground + msgType, msgTitle
	End If
    
    
    
' Helper
Sub sendPW_runAs(ByVal password)  ' v1.0
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

' Helper
' Powershell:   Escaping every whitespace offers more compability
'               than escaping the whole string with `'my_str`'
Function esc_spaces_PS(ByVal strValue)   ' v1.0
    esc_spaces_PS = Replace(strValue, " ", "` ")
End Function


' Helper
' Powershell:   -ArgumentList of start-process expects the format:
'               'arg1', 'arg2', 'arg3'
Function create_ArgumentStr_PS(ByRef listArgs_1, ByRef listArgs_2)   ' v1.1 
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


' Helper
' Powershell:	Expressions in single-quoted strings are not evaluated
'				In double-quoted strings any variable names such as "$myVar" will be replaced with the variable's value
Function qt_s(ByVal strValue)   ' v1.2
    qt_s = Chr(39) & strValue & Chr(39)
End Function

	
' Helper
Function qt(ByRef strValue)   ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function

' Helper
' runas needs escaped doublequotes
Function qt_esc_runAs(ByVal strValue)   ' v1.4
    escapedQuote = "\" & Chr(34)
    qt_esc_runAs = escapedQuote & strValue & escapedQuote
End Function

' Helper
' PowerShell escaped singlequotes
Function qt_esc_PS(ByVal strValue)   ' v1.1
    escapedQuote = "`" & Chr(39)
    qt_esc_PS = escapedQuote & strValue & escapedQuote
End Function

' Helper
' Powershell:   In double-quoted strings any variable names such as "$myVar"
'               will be replaced with the variable's value
Function qt_d(ByVal strValue)  ' v1.2
    doubleQuotes = Chr(34) & Chr(34)
    qt_d = doubleQuotes & strValue & doubleQuotes
End Function

' Helper
Function enclosedWith(ByRef str, ByRef strStart, ByRef strEnd)  ' v1.3
    enclosedWith = ((Left(Trim(LCase(str)), Len(strStart)) = LCase(strStart)) _ 
                And (Right(Trim(LCase(str)), Len(strEnd)) = LCase(strEnd)))
End Function

' Helper Bundle
' compareMethod: vbTextCompare (case insensitive)
'                vbBinaryCompare (case sensitive)
' cutChars_Begin + cutChars_End can be an Integer, a String or an Array of Strings
' With an array of possibilities the order does matter when the possibilities contain same chars
' ----------------------------------------------------
Function cutBeginEnd(ByRef str, ByRef cutChars_Begin, ByRef cutChars_End, ByRef compareMethod)  ' v2.0
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

Function StartsWith(ByRef str, ByRef start)  ' v1.3
	StartsWith = (Left(Trim(LCase(str)), Len(start)) = LCase(start))
End Function

Function EndsWith(ByRef str, ByRef ending)  ' v1.4
	EndsWith = (Right(Trim(LCase(str)), Len(ending)) = LCase(ending))
End Function
' ----------------------------------------------------

' Helper Bundle  v1.5
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

' Helper Bundle
' v1.5
' ----------------------------------------------------
Function getComplexArguments()		' v1.5
    Dim arrComplexArgs()
	Set objArgs = WScript.Arguments
	countArgs = objArgs.Count
	If countArgs > 0 Then
		strArgs = " """ & objArgs(0)
        ' read all arguments (are seperated by " " by default)
        ' mark each beginning of an argument part
        ' (" can quite safely be used for that, because the got removed when handed to the script)
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
	getComplexArguments = arrComplexArgs
End Function

Function Contains(ByRef str, ByRef strSearch)	' v1.3
	' converting to lower case is better than vbTextCompare because of dealing with foreign languages
	If InStr(LCase(str), LCase(strSearch)) > 0 Then returnValue = True Else returnValue = False
	Contains = returnValue
End Function

Function isEmpty_ArrList(ByRef arrOrList)	' v1.8
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

Function arr_SafeGet(ByRef arr, ByRef index, ByRef defaultValue)		' v1.1
    If UBound(arr) >= index Then arr_SafeGet = arr(index) Else arr_SafeGet = defaultValue
End Function

Sub show_MsgBox(ByRef msg_ArrOrStr, ByRef msgType, ByRef strTitle)  ' v1.2
    strSpace = "   "    ' space for text in msgBoxes
    If IsArray(msg_ArrOrStr) Then
        strMsg = Join(msg_ArrOrStr, vbCrlf & strSpace)
    Else
        strMsg = msg_ArrOrStr
    End If
	MsgBox strSpace & strMsg, msgType + vbMsgBoxSetForeground, strTitle
End Sub
' ----------------------------------------------------