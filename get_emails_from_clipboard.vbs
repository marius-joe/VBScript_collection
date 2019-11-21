' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Extract E-mail adresses from the current text in clipboard
'  v1.1.4
' ******************************************

Dim objShell_A
Dim objShell_WS
Dim objFSO


clipboardData = GetClipboardData
Set objRegEx = New RegExp
With objRegEx
    '.Pattern = "/.+@.+\..+/i"   ' Einfacher Check mit Punkt in Domain
    .Pattern = "\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.(?!\bjpg\b)[A-Z]{2,}\b"	' Check mit erlaubten Zeichen und Trennzeichen vor und nach Email, bekannte Dateiendungen werden auch ausgeschlossen
    .Global = True
    .IgnoreCase = True
End With

Set colResults = objRegEx.Execute(clipboardData)
Set dictResults = CreateObject("Scripting.Dictionary")
For Each result in colResults
    strResult = CStr(result)
    If Not dictResults.Exists(strResult) Then   
        dictResults.Add strResult, ""
    End If
Next

Call SetClipboardData(Join(dictResults.Keys, vbCrlf))


REM Alternative:
REM 'CLEAR
REM ClipBoard("")

REM 'SET
REM ClipBoard("Hello World!")

REM 'GET
REM Result = ClipBoard(Null)

REM Function ClipBoard(input)
  REM If IsNull(input) Then
    REM ClipBoard = CreateObject("HTMLFile").parentWindow.clipboardData.getData("Text")
    REM If IsNull(ClipBoard) Then ClipBoard = ""
  REM Else
    REM CreateObject("WScript.Shell").Run _
      REM "mshta.exe javascript:eval(""document.parentWindow.clipboardData.setData('text','" _
      REM & Replace(Replace(Replace(input, "'", "\\u0027"), """","\\u0022"),Chr(13),"\\r\\n") & "');window.close()"")", _
      REM 0,True
  REM End If
REM End Function


' Helper Bundle
' ----------------------------------------------------
Function AccessClipboard(copyText)     ' v1.0
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

' Helper
Function GetClipboardData()     ' v1.2
    GetClipboardData = AccessClipboard(Null)
End Function

' Helper
Sub SetClipboardData(copyText)     ' v1.1
    GetClipboardData = AccessClipboard(copyText)
End Sub
' ----------------------------------------------------

' Helper Bundle
' ----------------------------------------------------
Sub setGlobalsIfNecessary(strObjectNames)  ' v1.1
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

' Helper
Sub cleanGlobals(strObjectNames)  ' v1.3
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