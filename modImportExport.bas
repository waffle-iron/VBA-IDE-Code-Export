Attribute VB_Name = "modImportExport"
Option Explicit

'// if config file is available and ListConf is checked
'// then make file list, import and export from file
'// else make file list, import and export from module

Public Sub MakeFileList()

    Dim prjActVBProject     As VBProject
    Dim modFileList         As VBComponent
    Dim comComponent        As VBComponent
    Dim fsoFile             As Scripting.TextStream
    Dim FSO                 As New Scripting.FileSystemObject
    
    On Error GoTo CatchError
    
    Call CollectSettings
    
    '// name this project if it has not been already
    If ThisWorkbook.VBProject.Name <> strThisProjectName Then ThisWorkbook.VBProject.Name = strThisProjectName

    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    Set prjActVBProject = Application.VBE.ActiveVBProject
    
    '// Add logic sso that this project is not listed
    If prjActVBProject.Name = strThisProjectName Then Exit Sub
    
    '// determine if  the list needs to be in a module or txt file
    If blnMakeConfFile Then
        '// write out to conf file
        
        '// delete the file if it exists
        With FSO
            If .FileExists(strConfigFilePath) Then
                .DeleteFile strConfigFilePath
            End If
        End With
        '// create the file
        Set fsoFile = FSO.CreateTextFile(FSO.GetParentFolderName(Application.VBE.ActiveVBProject.Filename) & Application.PathSeparator & strConfigFileName)
        
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
                    fsoFile.WriteLine fComponentTypeToString(vbext_ct_Document) & ": " & comComponent.Name
            End Select
        Next
        
    Else '// add details to module modFileList
        On Error Resume Next
        Set modFileList = prjActVBProject.VBComponents("modFileList")
        On Error GoTo CatchError
        
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
    
ExitSub:
    Exit Sub

CatchError:
    MsgBox "Error building file list" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & Err.Description _
         , vbExclamation, "modImportExport.MakeFileList"
    GoTo ExitSub

End Sub


Sub ImportFiles()
    Dim prjActVBProject     As VBProject
    Dim modFileList         As VBComponent
    Dim strModuleName       As String
    Dim strActVBProjectDir  As String
    Dim intModRowCounter    As Integer
    Dim FSO                 As New Scripting.FileSystemObject
    Dim fsoFile             As Scripting.TextStream
    Dim strLine             As String

    On Error GoTo ErrHandler
    
    Call CollectSettings

    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    Set prjActVBProject = Application.VBE.ActiveVBProject
    
    '// TODO set to the global var for the project
    strActVBProjectDir = Left(prjActVBProject.Filename, Len(prjActVBProject.Filename) - _
                                                        Len(Dir(prjActVBProject.Filename, vbNormal)))
    
        '// determine if  the list needs to be in a module or txt file
    If blnMakeConfFile Then
        '// check that .conf file exists
        With FSO
            If Not .FileExists(strConfigFilePath) Then
                MsgBox "You need to create modFileList before you can import files!"
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
                    '// TODO identify ThisWorkbook and Worksheet before exporting
                Case Is = "Code Module:"
                    strModuleName = Right(strLine, Len(strLine) - 13)
                    prjActVBProject.VBComponents.Import (strActVBProjectDir & strModuleName & ".bas")
                Case Is = "Class Module:"
                    strModuleName = Right(strLine, Len(strLine) - 14)
                    prjActVBProject.VBComponents.Import (strActVBProjectDir & strModuleName & ".cls")
                Case Is = "UserForm:"
                    strModuleName = Right(strLine, Len(strLine) - 10)
                    prjActVBProject.VBComponents.Import (strActVBProjectDir & strModuleName & ".frm")
            End Select
            
        Loop
    Else
    
        '// Check modFileList module exists
        On Error Resume Next
        Set modFileList = prjActVBProject.VBComponents("modFileList")
        On Error GoTo ErrHandler
    
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
                    prjActVBProject.VBComponents.Import (strActVBProjectDir & strModuleName & ".bas")
                Case Is = "'Class:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 8)
                    prjActVBProject.VBComponents.Import (strActVBProjectDir & strModuleName & ".cls")
                Case Is = "'Form:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 7)
                    prjActVBProject.VBComponents.Import (strActVBProjectDir & strModuleName & ".frm")
                End Select
            Next intModRowCounter
        End With
    End If
    
    MsgBox "Finished building " & prjActVBProject.Name

    Exit Sub

ErrHandler:
    MsgBox "Error in Importing Files" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & Err.Description _
         , vbExclamation, "modImportExport.ImportFiles"
End Sub


Sub ExportFiles()

    Dim prjActVBProject     As VBProject
    Dim modFileList         As VBComponent
    Dim strModuleName       As String
    Dim strActVBProjectDir  As String
    Dim intModRowCounter    As Integer
    Dim FSO                 As New Scripting.FileSystemObject
    Dim fsoFile             As Scripting.TextStream
    Dim strLine             As String
    
    On Error GoTo ErrHandler
    
    Call CollectSettings
        
    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub
    Set prjActVBProject = Application.VBE.ActiveVBProject

    '// TODO set to the global var for the project
    strActVBProjectDir = Left(prjActVBProject.Filename, Len(prjActVBProject.Filename) - _
                                                        Len(Dir(prjActVBProject.Filename, vbNormal)))
    
    '// determine if  the list needs to be in a module or txt file
    If blnMakeConfFile Then
        '// check that .conf file exists
        With FSO
            If Not .FileExists(strConfigFilePath) Then
                MsgBox "You need to create modFileList before you can export files!"
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
                    '// TODO identify ThisWorkbook and Worksheet before exporting
                Case Is = "Code Module:"
                    strModuleName = Right(strLine, Len(strLine) - 13)
                    prjActVBProject.VBComponents(strModuleName).Export (strActVBProjectDir & strModuleName & ".bas")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                Case Is = "Class Module:"
                    strModuleName = Right(strLine, Len(strLine) - 14)
                    prjActVBProject.VBComponents(strModuleName).Export (strActVBProjectDir & strModuleName & ".cls")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                Case Is = "UserForm:"
                    strModuleName = Right(strLine, Len(strLine) - 10)
                    prjActVBProject.VBComponents(strModuleName).Export (strActVBProjectDir & strModuleName & ".frm")
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
                    prjActVBProject.VBComponents(strModuleName).Export (strActVBProjectDir & strModuleName & ".bas")
                    If UCase(strModuleName) <> UCase("modFileList") Then
                        prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                    End If
                Case Is = "'Class:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 8)
                    prjActVBProject.VBComponents(strModuleName).Export (strActVBProjectDir & strModuleName & ".cls")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                Case Is = "'Form:"
                    strModuleName = Right(.Lines(intModRowCounter, 1), Len(.Lines(intModRowCounter, 1)) - 7)
                    prjActVBProject.VBComponents(strModuleName).Export (strActVBProjectDir & strModuleName & ".frm")
                    prjActVBProject.VBComponents.Remove prjActVBProject.VBComponents(strModuleName)
                End Select
            Next intModRowCounter
        End With
    
    End If
    
    MsgBox "Finished exporting " & prjActVBProject.Name
    
    Exit Sub

ErrHandler:
    MsgBox "Error in Exporting Files" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & Err.Description _
         , vbExclamation, "modImportExport.ExportFiles"
End Sub

Sub ConfigureExport()
    frmConfigure.Show
End Sub


