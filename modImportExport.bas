Attribute VB_Name = "modImportExport"
Option Explicit

Public Const STRTHISPROJECTNAME     As String = "VBAExport"
Public Const STRCONFIGFILENAME      As String = "CodeExportFileList.conf"
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8

'// if config file is available and ListConf is checked
'// then make file list, import and export from file
'// else make file list, import and export from module

Public Sub MakeFileList()

    Dim prjActVBProject     As VBProject
    Dim strConfigFilePath   As String
    Dim comComponent        As VBComponent
    Dim fsoFile             As Scripting.TextStream
    Dim FSO                 As New Scripting.FileSystemObject
    Dim strDocumentName     As String

    On Error GoTo catchError

    Set prjActVBProject = Application.VBE.ActiveVBProject
    If prjActVBProject Is Nothing Then Exit Sub

    strConfigFilePath = ConfigFilePath(prjActVBProject)

    '// delete the config file if it exists
    With FSO
        If .FileExists(strConfigFilePath) Then
            .DeleteFile strConfigFilePath
        End If
    End With
    '// create the file
    Set fsoFile = FSO.CreateTextFile(strConfigFilePath)

    '// For each module form etc, add the name to the config file
    For Each comComponent In prjActVBProject.VBComponents
        Select Case comComponent.Type
            Case Is = vbext_ct_StdModule
                 fsoFile.WriteLine fComponentTypeToString(vbext_ct_StdModule) & ": " & comComponent.Name
            Case Is = vbext_ct_ClassModule
                fsoFile.WriteLine fComponentTypeToString(vbext_ct_ClassModule) & ": " & comComponent.Name
            Case Is = vbext_ct_MSForm
                fsoFile.WriteLine fComponentTypeToString(vbext_ct_MSForm) & ": " & comComponent.Name
            Case Is = vbext_ct_ActiveXDesigner
                fsoFile.WriteLine fComponentTypeToString(vbext_ct_ActiveXDesigner) & ": " & comComponent.Name
            Case Is = vbext_ct_Document
                '// determine id ThisWorkbook or not
                If comComponent.Properties(30).Name = "IsAddin" Then
                    fsoFile.WriteLine fComponentTypeToString(vbext_ct_Document) & ": " & comComponent.Name
                Else
                    strDocumentName = CleanIllegalCharacters(comComponent.Properties(7).Value)
                    fsoFile.WriteLine fComponentTypeToString(vbext_ct_Document) & ": " & comComponent.Name & "[" & strDocumentName & "]" '<ActualSheet name
                End If
        End Select
    Next

exitSub:
    Exit Sub

catchError:
    MsgBox "Error building file list" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & Err.Description _
         , vbExclamation, "modImportExport.MakeFileList"

    '// reset the ide menu
    Call auto_close
    Call auto_open

    GoTo exitSub

End Sub


Sub ExportFiles()

    Dim prjActVBProject     As VBProject
    Dim strConfigFilePath   As String
    Dim strVBASourceDirPath As String
    Dim strModuleName       As String
    Dim intModRowCounter    As Integer
    Dim FSO                 As New Scripting.FileSystemObject
    Dim fsoFile             As Scripting.TextStream
    Dim strLine             As String
    Dim strDocType          As String

    Dim modTemp             As VBIDE.CodeModule

    On Error GoTo ErrHandler

    Set prjActVBProject = Application.VBE.ActiveVBProject
    If prjActVBProject Is Nothing Then Exit Sub

    strConfigFilePath = ConfigFilePath(prjActVBProject)
    strVBASourceDirPath = VBASourceDirPath(prjActVBProject)

    '// check that .conf file exists
    With FSO
        If Not .FileExists(strConfigFilePath) Then
            MsgBox "You need to create file list config file before you can export files!"
            Exit Sub
        End If
    End With

    '// open the .conf file
    Set fsoFile = FSO.OpenTextFile(strConfigFilePath, ForReading)

    '// loop through each object listed in the .conf file and export with file extension
    Do Until fsoFile.AtEndOfStream
        strLine = fsoFile.ReadLine

        Select Case Left(strLine, InStr(strLine, ": "))
            Case Is = "Document Module:"
                strModuleName = Right(strLine, Len(strLine) - 17) '// Remove Document Module:
                If InStr(1, strLine, "[") <> 0 Then
                    strModuleName = Mid(strModuleName, 1, InStr(1, strModuleName, "[") - 1) '// Remove >Name
                End If
                '// this is taken from workbook and worksheet
                Select Case prjActVBProject.VBComponents(strModuleName).Properties(4).Name
                    Case Is = "AcceptLabelsInFormulas" '// Workbook
                        strDocType = ".wbk"
                    Case Is = "CodeName" '// Worksheet
                        strDocType = ".sht"
                End Select

                Set modTemp = prjActVBProject.VBComponents(strModuleName).CodeModule
                modTemp.Parent.Name = strModuleName & "_temp"
                prjActVBProject.VBComponents(modTemp.Parent.Name).Export (strVBASourceDirPath & strModuleName & strDocType)
                modTemp.Parent.Name = strModuleName

                modTemp.DeleteLines 1, modTemp.CountOfLines '// remove code from module

            Case Is = "Code Module:"
                strModuleName = Right(strLine, Len(strLine) - 13)
                prjActVBProject.VBComponents(strModuleName).Export (strVBASourceDirPath & strModuleName & ".bas")
                prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
            Case Is = "Class Module:"
                strModuleName = Right(strLine, Len(strLine) - 14)
                prjActVBProject.VBComponents(strModuleName).Export (strVBASourceDirPath & strModuleName & ".cls")
                prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
            Case Is = "UserForm:"
                strModuleName = Right(strLine, Len(strLine) - 10)
                prjActVBProject.VBComponents(strModuleName).Export (strVBASourceDirPath & strModuleName & ".frm")
                prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
        End Select

    Loop

    MsgBox "Finished exporting " & prjActVBProject.Name, vbInformation

exitSub:
    Exit Sub

ErrHandler:
    MsgBox "Error in Exporting Files" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & Err.Description _
         , vbExclamation, "modImportExport.ExportFiles"

    Call auto_close
    Call auto_open

    GoTo exitSub

End Sub


Sub ImportFiles()

    Dim prjActVBProject     As VBProject
    Dim strConfigFilePath   As String
    Dim strVBASourceDirPath As String
    Dim modFileList         As VBComponent
    Dim strModuleName       As String
    Dim strDocumentName     As String
    Dim intModRowCounter    As Integer
    Dim FSO                 As New Scripting.FileSystemObject
    Dim fsoFile             As Scripting.TextStream
    Dim strLine             As String

    Dim modCodeCopy         As VBIDE.CodeModule
    Dim modCodePaste        As VBIDE.CodeModule
    Dim modTemp             As VBComponent

    On Error GoTo catchError

    Set prjActVBProject = Application.VBE.ActiveVBProject
    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub

    strConfigFilePath = ConfigFilePath(prjActVBProject)
    strVBASourceDirPath = VBASourceDirPath(prjActVBProject)

    '// check that .conf file exists
    With FSO
        If Not .FileExists(strConfigFilePath) Then
            MsgBox "You need to create file list config file before you can import files!"
            Exit Sub
        End If
    End With

    '// open the .conf file
    Set fsoFile = FSO.OpenTextFile(strConfigFilePath, ForReading)

    '// loop through each object listed in the .conf file and export with file extension
    Do Until fsoFile.AtEndOfStream
        strLine = fsoFile.ReadLine

        Select Case Left(strLine, InStr(strLine, ": "))
            Case Is = "Document Module:"
                strModuleName = Right(strLine, Len(strLine) - 17)
                '// this is taken from workbook and worksheet
                If InStr(1, strModuleName, "[") > 0 Then
                    strModuleName = Left(strModuleName, InStr(1, strModuleName, "[") - 1)
                    strDocumentName = CleanIllegalCharacters(Mid(strLine, InStr(1, strLine, "["), InStrRev(strLine, "[")))
                End If

                Select Case prjActVBProject.VBComponents(strModuleName).Properties(4).Name
                    Case Is = "AcceptLabelsInFormulas" '// AcceptLabelsInFormulas=Workbook
                        prjActVBProject.VBComponents.Import (strVBASourceDirPath & strModuleName & ".wbk")
                    Case Is = "CodeName" '// CodeName=Worksheet
                        prjActVBProject.VBComponents.Import (strVBASourceDirPath & strModuleName & ".sht")
                End Select

                On Error Resume Next
                Set modTemp = prjActVBProject.VBComponents(strModuleName & "_temp")
                On Error GoTo catchError

                Set modCodeCopy = prjActVBProject.VBComponents(modTemp.Name).CodeModule
                Set modCodePaste = prjActVBProject.VBComponents(strModuleName).CodeModule

                modCodePaste.DeleteLines 1, modCodePaste.CountOfLines

                If modCodeCopy.CountOfLines > 0 Then
                    modCodePaste.AddFromString modCodeCopy.Lines(1, modCodeCopy.CountOfLines)
                End If

                '// module already exists, so first remove it
                prjActVBProject.VBComponents.Remove modTemp

            Case Is = "Code Module:"
                strModuleName = Right(strLine, Len(strLine) - 13)
                prjActVBProject.VBComponents.Import (strVBASourceDirPath & strModuleName & ".bas")
            Case Is = "Class Module:"
                strModuleName = Right(strLine, Len(strLine) - 14)
                prjActVBProject.VBComponents.Import (strVBASourceDirPath & strModuleName & ".cls")
            Case Is = "UserForm:"
                strModuleName = Right(strLine, Len(strLine) - 10)
                prjActVBProject.VBComponents.Import (strVBASourceDirPath & strModuleName & ".frm")
        End Select

    Loop

    MsgBox "Finished building " & prjActVBProject.Name, vbInformation

exitSub:
    Exit Sub

catchError:
    MsgBox "Error in Importing Files" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & Err.Description _
         , vbExclamation, "modImportExport.ImportFiles"

    Call auto_close
    Call auto_open

    GoTo exitSub

End Sub

' Config file path for a given VBProject
Private Function ConfigFilePath(Project As VBProject) As String
    ConfigFilePath = ProjParentDirPath(Project) & STRCONFIGFILENAME
End Function

' Path of the VBA source directory for a given VBProject
Private Function VBASourceDirPath(Project As VBProject) As String
    VBASourceDirPath = ProjParentDirPath(Project)
End Function

' The parent directory path for a given VBProject
Private Function ProjParentDirPath(Project As VBProject) As String
    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    ProjParentDirPath = FSO.GetParentFolderName(Project.Filename) & Application.PathSeparator
End Function

Private Function fComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
            fComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            fComponentTypeToString = "Class Module"
        Case vbext_ct_Document
            fComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm
            fComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule
            fComponentTypeToString = "Code Module"
        Case Else
            fComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function

Private Function CleanIllegalCharacters(strClean As String) As String

    On Error GoTo catchError

    strClean = Replace(strClean, "[", "")
    strClean = Replace(strClean, "]", "")
    strClean = Replace(strClean, "-", "_")
    strClean = Replace(strClean, " ", "_")

    CleanIllegalCharacters = strClean

exitFunction:
    Exit Function
catchError:

    MsgBox "Error cleaning string." & vbCrLf & "Error Number: " & Err.Number & vbCrLf & Err.Description _
        , vbExclamation, "modFunctions.CleanIllegalCharacters"

    GoTo exitFunction
End Function
