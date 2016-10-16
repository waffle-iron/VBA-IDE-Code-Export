Attribute VB_Name = "modSettings"
Option Explicit

Public Const strThisProjectName As String = "VBAExport"
Public Const strConfigFileName  As String = "CodeExportFileList.conf"
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8

Public blnConfigAvailable       As Boolean
Public blnMakeConfFile          As Boolean

Public strExportTo              As String
Public strImportFrom            As String
Public strConfigFilePath        As String

Sub CollectSettings()
        
    '// so this will populate the global
    '// vars with the configured file locations if
    '// the .conf file exists
    
    '// first check for the config file
    If ConfFileExists Then
        '// populate global vars
        Dim tsFile      As Scripting.TextStream
        Dim strFileName As String
        Dim strTextLine As String
        
        Dim FSO As New Scripting.FileSystemObject
        
        strFileName = strConfigFilePath
        
        Set tsFile = FSO.OpenTextFile(strFileName, ForReading)
        
        Do Until tsFile.AtEndOfStream
            strTextLine = tsFile.ReadLine

            If Left(strTextLine, InStr(strTextLine, ":") - 1) = "ImportFrom" Then
                strImportFrom = Right(strTextLine, Len(strTextLine) - Len(Left(strTextLine, InStr(strTextLine, ":"))))
                shtConfig.Range("rImportFrom") = strImportFrom
            ElseIf Left(strTextLine, InStr(strTextLine, ":") - 1) = "ExportTo" Then
                strExportTo = Right(strTextLine, Len(strTextLine) - Len(Left(strTextLine, InStr(strTextLine, ":"))))
                shtConfig.Range("rExportTo") = strExportTo
            End If
            
        Loop
        tsFile.Close
        
        blnMakeConfFile = shtConfig.Range("rComponentTXTList")
        
    Else '// default settings
        strExportTo = FSO.GetParentFolderName(Application.VBE.ActiveVBProject.FileName)
        strImportFrom = FSO.GetParentFolderName(Application.VBE.ActiveVBProject.FileName)
    End If

End Sub




