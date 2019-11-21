    Set objShell_WS = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    path_ParentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
    
    relPath_Program = "\blub\blub.exe"
    args_Program = ""
    
    waitForExit = False
    runWindowVisibility = 1
    path_Program = objFSO.BuildPath(path_ParentFolder, relPath_Program)    
    strCmd = qt(path_Program) & " " & args_Program
    objShell_WS.Run strCmd, runWindowVisibility, waitForExit

    
' Helper
Function qt(ByVal strValue)
    strReturn = strValue
    If Left(strValue, 1) <> Chr(34) Then strReturn = Chr(34) & strReturn
    If Right(strValue, 1) <> Chr(34) Then strReturn = strReturn & Chr(34)
    qt = strReturn
End Function