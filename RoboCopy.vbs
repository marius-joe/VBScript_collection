' ******************************************
'  Dev:  marius-joe
' ******************************************
'  v1.2
' ******************************************
'  Arguments: "-source ..." "-destFolder ..." [opt] "-createDest" [opt] "-move" [opt] "-silent"
'  (Case Or Order don't matter)

Const C_isDebugMode = True      ' turns on debug mode

Const C_strSpace = "   "    ' space for text in msgBoxes

Dim objFSO
Dim objShell_WS
Dim objShell_A


    If C_isDebugMode Then runWindowVisibility = 1 Else runWindowVisibility = 0
	Set listComplexArgs = getListComplexArguments()
    isArgError = False
	If Not isEmpty_ArrList(listComplexArgs) Then 
        source = ""
        destFolder = ""
        createDest = False
        move = False
        isSilentMode = False   
		For Each arrComplexArg In listComplexArgs
			argument = LCase(arrComplexArg(0))
            value = arr_SafeGet(arrComplexArg, 1)

			Select Case argument
				Case "-source"
					If value <> "" Then path_Source = value Else isArgError = True
                    

				Case "-destfolder"
					If value <> "" Then destFolder = value Else isArgError = True
                    
					
				Case "-createdest"
					createDest = True
                    
                    
				Case "-move"
					move = True
                    
                    
				Case "-silent"
					isSilentMode = True
			End Select
		Next
    Else
        isArgError = True
    End If
    
    If Not isArgError Then
        Call roboCopy(path_Source, destFolder, createDest, move, isSilentMode)
    Else
        MsgBox "args sind kacke"
    End If
    
    Call cleanGlobals("All")

    
        
' Helper Bundle
' ----------------------------------------------------
' Quellpfad Formate:
' C:\Users\zuKopierenderOrdner
' C:\Users\*
' C:\Users\*.txt
' C:\Users\copyMe.txt
Sub roboCopy(path_Source, path_DestFolder, createDestFolder, moveData, isSilentMode)  ' v1.4
    If isValidPath(path_Source, true) Then      ' true: allow wildcard in path      
        If createDestFolder Then Call ensurePath(path_DestFolder)
        Call setGlobalsIfNecessary("objShell_WS, objFSO, objShell_A")
        Set listArgs = CreateObject("System.Collections.ArrayList")
            REM /E 	Kopiert Unterverzeichnisse, auch die leeren
            REM /R:1 	Es wird bei einem Fehler 1x versucht, die Datei erneut zu kopieren.
            REM /W:1 	Es wird 1 Sekunde gewartet, bevor ein erneuter Kopierversuch gestartet wird.
            REM /MIR 	Spiegelt einen gesamten Verzeichnisbaum und löscht am Ziel Daten, welche auf der Quelle nicht vorhanden sind. (Entspricht: /e und /PURGE).
            REM /COPYALL 	Kopiert alle Dateiinformationen (Entspricht: D Data; A Attributes; T Time stamps; S NTFS access control list (ACL); O Owner information; U Auditing information).
            REM /MT:16 	Multi Tasking; Die Zahl gibt an, wieviele Kopiervorgänge gleichzeitig ausgeführt werden. Bei der Angabe von nur MT wird das Default 8 verwendet. Steigert die Performance uU erheblich
            REM /SECFIX 	immer zusammen mit /sec benutzen (sec ist hier schon in DATS einegbaut), Repariert die Dateiberechtigungen auch an Dateien, die sich nicht geändert haben2)
            REM /B 	Dadurch kann ein Administrator auch Dateien kopieren, auf die er normalerweise keinen Zugriff hat, sofern er das für Administratoren voreingestellte Recht für Backups besitzt.
            REM /XJ 	überpringt sogenannte Junctions; dabei handelt es sich um spezielle Ordnerverknüpfungen im NTFS-Dateisystem. Lässt man den Schalter weg, kopiert Robocopy den Inhalt des Ordners, auf den die Junction verweist. Auch Hard Links, eine weitere NTFS-Spezialität, kann Robocopy als solche nicht kopieren und behandelt sie wie herkömmliche Dateien
            REM /DCOPY:T 	Es werden auch die Zeitstempel für Verzeichnisse kopiert.
            REM /XD 	Exclude Directory. Die angebenen Verzeichnisse werden nicht synchronisiert
            REM /LOG:C:\cpfiles.log 	Schreibt ein Log der Kopiervorgänge in die Datei C:\cpfiles.log. Es werden keine Logs in die Konsole geschrieben. 
            REM /XO     Bereits existierende gleiche oder neuere Dateien werden nicht überschrieben
            REM /FFT    Diese Option toleriert Zeitstempel mit einer Differenz von 2s um sie als neu zu erkennen (Nützlich beim kopieren von NTFS nach FAT oder auf NAS Geräte)
            REM /LOG:"%USERPROFILE%\Documents\Scripts\Lookeen Index Downloader\log.txt" /NJH /FP /NDL /NFL /NP
            REM kleinerer Log
            REM /L /TEE     Vorgang nur simulieren und Ausgabe in Console + Datei
            
        Set listExclude = CreateObject("System.Collections.ArrayList")
        listExclude.Add "/XF"
        
        sourceType = ""
        If objFSO.FolderExists(path_Source) Then
            sourceType = "folder"
            path_SourceFolder = path_Source
            listArgs.Add "/E /XJ"                   ' alle Dateien und Unterordner eines Ordners kopieren, auch leere Ordner
        Else
            indexLastBackSlash = InStrRev(path_Source, "\")
            path_SourceFolder = Left(path_Source, indexLastBackSlash-1)
            file = Mid(path_Source, indexLastBackSlash+1)
            
            If file = "*" Then					    ' alle Dateien und Unterordner eines Ordners kopieren, auch leere Ordner
                sourceType = "folderContent"
                listArgs.Add "/E /XJ"
                listExclude.Add qt_s("desktop.ini")
            Else
                sourceType = "files"
                listArgs.Add qt_s(esc_spaces_PS(file))       	    ' Dateien oder auch nur bestimmte Datei-Endungen kopieren
            End If
        End If
                                          
        If moveData Then    ' Bei Verwendung von /MOVE kann /MT nicht verwende werden, führt zu Problemen
            listArgs.Add "/MOVE"
            If sourceType = "folderContent" Then
                path_LockFile = objFSO.BuildPath(path_SourceFolder, ".keepRoot")
                objFSO.CreateTextFile path_LockFile, True
                listExclude.Add qt_s(esc_spaces_PS(path_LockFile))     ' to prevent the root folder from deleting when only movin the content (*) is set 
            End If
        Else
            listArgs.Add "/MT:32"
        End If

        listArgs.Add "/R:2 /W:3"		' RetrySettings: try 2x 3 Sec
        listArgs.Add "/XO /COPY:DATSO /SECFIX /DCOPY:T /B /FFT"
        
' /fft ist eigentlich nur bei Move und overwrite wichtig    
        
        If C_isDebugMode Then listArgs.Add "/L /TEE"   ' args_simulation
        
        strArgs = Join(listArgs.ToArray, " ")
        strArgs_Exclude = ""
        If listExclude.Count > 1 Then strArgs_Exclude = Join(listExclude.ToArray, " ")
        
        program = "RoboCopy.exe"
        strArgs_Program = qt_s(esc_spaces_PS(path_SourceFolder)) & " " & qt_s(esc_spaces_PS(path_DestFolder)) & " " _
                            & qt_s(strArgs & " " & strArgs_Exclude)
        

        If not silent Then
            asAdmin = True
            If asAdmin Then
                userRights = " -Verb runAs"
            Else 
                userRights = ""
                strArgs_cmdlet = " -NoNewWindow"       ' -NoNewWindow kann nicht zusammen mit  " -Verb runAs"  benutzt werden. 
            End If       

MsgBox "Blub jetzt kommt powershell"
            ' the script calls Powershell, Powershell then calls the program needed to execute the file as admin - has to be done this way, to support "wait" functionality + the correct UAC Message which .ShellExecute does not                
            define_strArgs_cmdlet = "$arr_arguments = " & strArgs_Program & " ;"
            strArgs_cmdlet = strArgs_cmdlet & " -ArgumentList $arr_arguments"
            strCmd = "powershell -noexit " &  "-Command " & qt(define_strArgs_cmdlet & " start-process " & qt_s(program) & userRights & " -Wait" & strArgs_cmdlet) '-PassThru ? -NoNewWindow
MsgBox strCmd                  
            ' powershell wohl nötig, weil shellexecute kein wait kann, doch geht wohl über start cmd /wait oder Call (ist wie start mit /wait aber in geicher shell )blubb aber besser testen
   'z.B. -ArgumentList "/c dir `"%systemdrive%\program files`""         
'-ArgumentList
'Specifies parameters or parameter values to use when this cmdlet starts the process.
'If parameters or parameter values contain a space, they need surrounded with escaped double quotes.                
            
            objShell_WS.Run strCmd, runWindowVisibility, false                
            
            exitCode = -1
        Else
            exitCode = objShell_A.ShellExecute("cmd", qt("RoboCopy.Exe " & strRoboArgs), "", "runas", 0)
        End If
        
        If moveData AND sourceType = "folderContent" Then objFSO.DeleteFile path_LockFile, True
   
        If not silent Then
            strRoboInfo = vbCrlf & vbCrlf & path_Source & "   ->" & vbCrlf & path_DestFolder 
            If exitCode < 8 Then
                If moveData Then
                    allMoved = False
                    If sourceType = "folderContent" Then
                        If IsFolderEmpty(path_SourceFolder) Then allMoved = True    ' with the wildcard * at the end of the source path only the content is moved
                    Else
                        If Not objFSO.FolderExists(path_SourceFolder) Then allMoved = True     ' if the source path has no wildcard at the end the whole folder is moved
                    End If
                    
                    If allMoved Then
                        MsgBox "Verschiebe-Vorgang erfolgreich." & strRoboInfo                      
                    Else
                        MsgBox "Verschiebe-Vorgang erfolgreich. Moeglicherweise noch vorhandene Dateien im Quellordner waren in neuerer oder gleicher Version schon im Ziel enthalten" & strRoboInfo                      
                    End If
                Else
                   MsgBox "Kopier-Vorgang erfolgreich, neuere Dateien im Ziel wurden nicht ueberschrieben" & strRoboInfo
                End If
            Else
                MsgBox "Beim Kopier/Verschiebe-Vorgang sind Fehler aufgetreten" & strRoboInfo               
            End If
        End If
    Else
        MsgBox "Error with Source Path"
    End If
End Sub

' Helper
Function IsFolderEmpty(path_folder)  ' v1.0
    Call setGlobalsIfNecessary("objFSO")
    Set folder = objFSO.getFolder(path_folder)
    If folder.Files.Count = 0 And folder.SubFolders.Count = 0 Then
        returnValue = True
    Else
        returnValue = False
    End If
    IsFolderEmpty = returnValue
End Function


' Helper
Function isValidPath(path, allowWildCard)  ' v1.3
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


' Helper
Function ensurePath(ByVal strPath)  ' v1.0
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
' v1.0
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

Function Contains(ByRef str, ByRef strSearch)	' v1.3
	' converting to lower case is better than vbTextCompare because of dealing with foreign languages
	If InStr(LCase(str), LCase(strSearch)) > 0 Then returnValue = True Else returnValue = False
	Contains = returnValue
End Function

Function isEmpty_ArrList(ByRef arrOrList)	' v1.7
    functionName = "isEmpty_ArrList"
	isEmpty = True
	If IsArray(arrOrList) Then		' is array
		On Error Resume Next
			UBarr = UBound(arrOrList)
			If (Err.Number = 0) And (UBarr >= 0) Then isEmpty = False
		On Error GoTo 0
	ElseIf TypeName(arrOrList) = "ArrayList" Then	 ' is list
        If arrOrList.Count > 0 Then
            isEmpty = False
        End If
    Else
        Call show_MsgBox("Variable 'arrOrList' is no Array or ArrayList: " & TypeName(arrOrList), vbCritical, "Function: " & functionName)
    End If
	
	isEmpty_ArrList = isEmpty
End Function

Function arr_SafeGet(ByRef arr, ByRef index)		' v1.0
    If UBound(arr) >= index Then arr_SafeGet = arr(index) Else arr_SafeGet = ""
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

' Helper
' Powershell:	Expressions in single-quoted strings are not evaluated
'				In double-quoted strings any variable names such as "$myVar" will be replaced with the variable's value
Function qt_s(ByVal strValue)   ' v1.2
    qt_s = Chr(39) & strValue & Chr(39)
End Function

	
' Helper
Function qt(ByVal strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function

' Helper
' Powershell:   Escaping every whitespace offers more compability
'               than escaping the whole string with `'my_str`'
Function esc_spaces_PS(ByVal strValue)   ' v1.0
    esc_spaces_PS = Replace(strValue, " ", "` ")
End Function