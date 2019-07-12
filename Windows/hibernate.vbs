' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Enter hibernation mode after confirmation
'  v1.0.0
' ******************************************

Const C_strSpace = "   "    ' space for text in msgBoxes

Dim isSilentMode


	isSilentMode = False
	Set objArgs = WScript.Arguments
	If objArgs.Count > 0 Then
		If LCase(objArgs(0)) = "silent" Then
			isSilentMode = True
		End If
	End If

    If Not isSilentMode Then
        msgType = vbExclamation
        strMsg = "Ruhezustand jetzt starten ?"
        strTitle = "Ruhezustand"
        yesNo = MsgBox(C_strSpace & strMsg, vbYesNo + vbDefaultButton2 + msgType + vbMsgBoxSetForeground, strTitle)
    Else
        yesNo = vbYes
    End If
    
    If yesNo = vbYes Then
        Dim objShell_WS
        Set objShell_WS = CreateObject("WScript.Shell")

        ' hibernate
        objShell_WS.Run("rundll32.exe powrprof.dll, SetSuspendState")

        Set objShell_WS = Nothing
    End If