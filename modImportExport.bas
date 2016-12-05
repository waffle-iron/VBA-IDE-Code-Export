Attribute VB_Name = "modImportExport"
Option Explicit

'// Worksheet named ranges
'// rImportFrom
'// rExportTo
'// rComponentTXTList
'// rConfFileName

'// if config file is available and ListConf is checked
'// then make file list, import and export from file
'// else make file list, import and export from module

Public Sub MakeFileList()

    Dim prjActVBProject     As VBProject
    Dim modFileList         As VBComponent
    Dim comComponent        As VBComponent
    Dim fsoFile             As Scripting.TextStream
    Dim FSO                 As New Scripting.FileSystemObject
    Dim strDocumentName     As String
    
    On Error GoTo catchError
    
    '// name this project if it has not been already
    If ThisWorkbook.VBProject.Name <> STRTHISPROJECTNAME Then ThisWorkbook.VBProject.Name = STRTHISPROJECTNAME

    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    Set prjActVBProject = Application.VBE.ActiveVBProject
    
    '// Add logic sso that this project is not listed
    If prjActVBProject.Name = STRTHISPROJECTNAME Then Exit Sub
    
    '// determine if  the list needs to be in a module or txt file
    If g_blnMakeConfFile Then
        '// write out to conf file
        
        '// delete the file if it exists
        With FSO
            If .FileExists(g_strConfigFilePath) Then
                .DeleteFile g_strConfigFilePath
            End If
        End With
        '// create the file
        Set fsoFile = FSO.CreateTextFile(FSO.GetParentFolderName(g_strActiveVBProjectName) & Application.PathSeparator & STRCONFIGFILENAME)
        
        '// Add import and export locations
        fsoFile.WriteLine "ImportFrom:" & shtConfig.Range("rImportFrom")
        fsoFile.WriteLine "ExportTo:" & shtConfig.Range("rExportTo")
        
        '// For each module form etc, add the name to the modFileList Module
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
        
    Else '// add details to module modFileList
        On Error Resume Next
        Set modFileList = prjActVBProject.VBComponents("modFileList")
        On Error GoTo catchError
        
        If modFileList Is Nothing Then
            '// module does not already exist
        Else
            '// module already exists, so first remove it
            prjActVBProject.VBComponents.Remove modFileList
        End If
    
        '// Add module
        Set modFileList = prjActVBProject.VBComponents.Add(vbext_ct_StdModule)
        modFileList.Name = "modFileList"
    
        With modFileList.CodeModule
            .AddFromString ("'DO NOT DELETE THIS MODULE")
    
            '// For each module form etc, add the name to the modFileList Module
            For Each comComponent In prjActVBProject.VBComponents
                Select Case comComponent.Type
                Case Is = vbext_ct_StdModule
                    If UCase(comComponent.Name) <> UCase("modFileList") Then
                        .AddFromString ("'Module: " & comComponent.Name)
                    End If
                Case Is = vbext_ct_ClassModule
                    .AddFromString ("'Class: " & comComponent.Name)
                Case Is = vbext_ct_MSForm
                    .AddFromString ("'Form: " & comComponent.Name)
                Case Is = vbext_ct_ActiveXDesigner
                    .AddFromString ("'Designer: " & comComponent.Name)
                End Select
            Next
        End With
    
    End If
    
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
    Dim modFileList         As VBComponent
    Dim strModuleName       As String
    Dim intModRowCounter    As Integer
    Dim FSO                 As New Scripting.FileSystemObject
    Dim fsoFile             As Scripting.TextStream
    Dim strLine             As String
    Dim strDocType          As String
    
    Dim modTemp             As VBIDE.CodeModule
    
    On Error GoTo ErrHandler
    
    '// if checked make file list
    
        
    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    Set prjActVBProject = Application.VBE.ActiveVBProject
    
    '// determine if  the list needs to be in a module or txt file
    If g_blnMakeConfFile Then
        '// check that .conf file exists
        With FSO
            If Not .FileExists(g_strConfigFilePath) Then
                MsgBox "You need to create modFileList before you can export files!"
                Exit Sub
            End If
        End With
        
        '// open the .conf file
        Set fsoFile = FSO.OpenTextFile(g_strConfigFilePath, ForReading)
        
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
                    prjActVBProject.VBComponents(modTemp.Parent.Name).Export (g_strExportTo & strModuleName & strDocType)
                    modTemp.Parent.Name = strModuleName
                    
                    modTemp.DeleteLines 1, modTemp.CountOfLines '// remove code from module
                
                Case Is = "Code Module:"
                    strModuleName = Right(strLine, Len(strLine) - 13)
                    prjActVBProject.VBComponents(strModuleName).Export (g_strExportTo & strModuleName & ".bas")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                Case Is = "Class Module:"
                    strModuleName = Right(strLine, Len(strLine) - 14)
                    prjActVBProject.VBComponents(strModuleName).Export (g_strExportTo & strModuleName & ".cls")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                Case Is = "UserForm:"
                    strModuleName = Right(strLine, Len(strLine) - 10)
                    prjActVBProject.VBComponents(strModuleName).Export (g_strExportTo & strModuleName & ".frm")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
            End Select
            
        Loop
        
    Else
    
        '// Check modFileList module exists
        On Error Resume Next
        Set modFileList = prjActVBProject.VBComponents("modFileList")
        On Error GoTo ErrHandler
    
        '// If modFileList module doesnt exist, you need to warn user then exit sub
        If modFileList Is Nothing Then
            MsgBox "You need to create modFileList before you can export files!"
            Exit Sub
        End If
    
        With modFileList.CodeModule
            '// loop through each module name listed in modFileList and import the associated module
            For intModRowCounter = 1 To .CountOfDeclarationLines
                Select Case Left(.Lines(intModRowCounter, 1), InStr(.Lines(intModRowCounter, 1), ": "))
                Case Is = "'Module:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 9)
                    prjActVBProject.VBComponents(strModuleName).Export (g_strExportTo & strModuleName & ".bas")
                    If UCase(strModuleName) <> UCase("modFileList") Then
                        prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                    End If
                Case Is = "'Class:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 8)
                    prjActVBProject.VBComponents(strModuleName).Export (g_strExportTo & strModuleName & ".cls")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                Case Is = "'Form:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 7)
                    prjActVBProject.VBComponents(strModuleName).Export (g_strExportTo & strModuleName & ".frm")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                End Select
            Next intModRowCounter
        End With
    
    End If
    
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

    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    Set prjActVBProject = Application.VBE.ActiveVBProject
    
    '// determine if  the list needs to be in a module or txt file
    If g_blnMakeConfFile Then
        '// check that .conf file exists
        With FSO
            If Not .FileExists(g_strConfigFilePath) Then
                MsgBox "You need to create modFileList before you can import files!"
                Exit Sub
            End If
        End With
        
        '// open the .conf file
        Set fsoFile = FSO.OpenTextFile(g_strConfigFilePath, ForReading)
        
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
                            prjActVBProject.VBComponents.Import (g_strImportFrom & strModuleName & ".wbk")
                        Case Is = "CodeName" '// CodeName=Worksheet
                            prjActVBProject.VBComponents.Import (g_strImportFrom & strModuleName & ".sht")
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
                    prjActVBProject.VBComponents.Import (g_strImportFrom & strModuleName & ".bas")
                Case Is = "Class Module:"
                    strModuleName = Right(strLine, Len(strLine) - 14)
                    prjActVBProject.VBComponents.Import (g_strImportFrom & strModuleName & ".cls")
                Case Is = "UserForm:"
                    strModuleName = Right(strLine, Len(strLine) - 10)
                    prjActVBProject.VBComponents.Import (g_strImportFrom & strModuleName & ".frm")
            End Select
            
        Loop
    Else
    
        '// Check modFileList module exists
        On Error Resume Next
        Set modFileList = prjActVBProject.VBComponents("modFileList")
        On Error GoTo catchError
    
        '// If modFileList module doesnt exist, you need to warn user then exit sub
        If modFileList Is Nothing Then
            MsgBox "You need to create modFileList before you can import files!"
            Exit Sub
        End If
    
        With modFileList.CodeModule
            '// loop through each module name listed in modFileList and import the associated module
            For intModRowCounter = 1 To .CountOfDeclarationLines
                Select Case Left(.Lines(intModRowCounter, 1), InStr(.Lines(intModRowCounter, 1), ": "))
                Case Is = "'Module:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 9)
                    prjActVBProject.VBComponents.Import (g_strImportFrom & strModuleName & ".bas")
                Case Is = "'Class:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 8)
                    prjActVBProject.VBComponents.Import (g_strImportFrom & strModuleName & ".cls")
                Case Is = "'Form:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 7)
                    prjActVBProject.VBComponents.Import (g_strImportFrom & strModuleName & ".frm")
                End Select
            Next intModRowCounter
        End With
    End If
    
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


Sub ConfigureExport()
    frmConfigure.Show
End Sub


