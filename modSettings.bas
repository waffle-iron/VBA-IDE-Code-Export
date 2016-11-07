Attribute VB_Name = "modSettings"
Option Explicit

Public Const STRTHISPROJECTNAME     As String = "VBAExport"
Public Const STRCONFIGFILENAME      As String = "CodeExportFileList.conf"
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8

Public g_blnConfigAvailable         As Boolean
Public g_blnMakeConfFile            As Boolean
Public g_blnBuildFileListOnExport   As Boolean

Public g_strExportTo                As String
Public g_strImportFrom              As String
Public g_strConfigFilePath          As String
Public g_strActiveVBProjectName     As String


Sub CollectSettings()
        
    '// so this will populate the global
    '// vars with the configured file locations if
    '// the .conf file exists
    
     g_strActiveVBProjectName = Application.VBE.ActiveVBProject.Filename
    
    '// first check for the config file
    If fConfFileExists Then
        '// populate global vars
        Dim tsFile      As Scripting.TextStream
        Dim strFileName As String
        Dim strTextLine As String
                
        Dim FSO As New Scripting.FileSystemObject
        
        Set tsFile = FSO.OpenTextFile(g_strConfigFilePath, ForReading)
        
        Do Until tsFile.AtEndOfStream
            strTextLine = tsFile.ReadLine

            If Left(strTextLine, InStr(strTextLine, ":") - 1) = "ImportFrom" Then
                g_strImportFrom = Right(strTextLine, Len(strTextLine) - Len(Left(strTextLine, InStr(strTextLine, ":"))))
                shtConfig.Range("rImportFrom") = g_strImportFrom
            ElseIf Left(strTextLine, InStr(strTextLine, ":") - 1) = "ExportTo" Then
                g_strExportTo = Right(strTextLine, Len(strTextLine) - Len(Left(strTextLine, InStr(strTextLine, ":"))))
                shtConfig.Range("rExportTo") = g_strExportTo
            End If
            
        Loop
        tsFile.Close
        
    Else '// default settings
        g_strExportTo = fAddPathSeparator(FSO.GetParentFolderName(g_strActiveVBProjectName))
        g_strImportFrom = fAddPathSeparator(FSO.GetParentFolderName(g_strActiveVBProjectName))
    End If
    
    g_blnMakeConfFile = shtConfig.Range("rComponentTXTList")
    
End Sub


Sub UpdateFile(fileToUpdate As String, targetText As String, replaceText As String)

    Dim tempName    As String
    Dim tempFile    As Scripting.TextStream
    Dim file        As Scripting.TextStream
    Dim currentLine As String
    Dim newLine     As String
    Dim FSO         As New Scripting.FileSystemObject

    '// creates a temp file and outputs the original files contents but with the replacements
    tempName = fileToUpdate & ".tmp"
    Set tempFile = FSO.CreateTextFile(tempName, True)

    '// open the original file and for each line replace any matching text
    Set file = FSO.OpenTextFile(fileToUpdate)
    Do Until file.AtEndOfStream
        currentLine = file.ReadLine
        newLine = Replace(currentLine, targetText, replaceText)
        '// write to the new line containing replacements to the temp file
        tempFile.WriteLine newLine
    Loop
    
    file.Close
    tempFile.Close

    '// delete the original file and replace with the temporary file
    FSO.DeleteFile fileToUpdate, True
    FSO.MoveFile tempName, fileToUpdate
    
End Sub


