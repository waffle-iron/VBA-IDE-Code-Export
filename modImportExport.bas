Attribute VB_Name = "modImportExport"
Option Explicit

'// Add references for :
'//     Microsoft Visual Basic For Applications Extensibility 5.3
'//     Microsoft Scripting Runtime
'// Also check the 'Trust access to the VBA project model check box', located...
'// Trust Centre, Trust Centre Settings, Macro Settings, Trust access to the VBA project model

Private Const STRCONFIGFILENAME         As String = "CodeExportFileList.conf"
Private Const STR_CONFIGKEY_MODULEPATHS As String = "Module Paths"
Private Const ForReading                As Integer = 1


'// if config file is available and ListConf is checked
'// then make file list, import and export from file
'// else make file list, import and export from module
Public Sub MakeFileList()

    Dim prjActVBProject     As VBProject
    Dim strConfigFilePath   As String
    Dim comComponent        As VBComponent
    Dim tsConfigStream      As Scripting.TextStream
    Dim FSO                 As Scripting.FileSystemObject
    Dim dictConfig          As Dictionary
    Dim dictModulePaths     As Dictionary
    Dim strConfigJson       As String
    Dim strExtension        As String

    On Error GoTo catchError

    Set prjActVBProject = Application.VBE.ActiveVBProject
    If prjActVBProject Is Nothing Then Exit Sub

    '// Collect the name of each module, form, etc.
    Set dictModulePaths = New Dictionary
    For Each comComponent In prjActVBProject.VBComponents

        strExtension = vbNullString
        Select Case comComponent.Type
            Case vbext_ct_StdModule
                strExtension = "bas"
            Case vbext_ct_ClassModule, vbext_ct_Document
                strExtension = "cls"
            Case vbext_ct_MSForm
                strExtension = "frm"
        End Select

        If Not strExtension = vbNullString Then
            dictModulePaths.Add comComponent.Name, comComponent.Name & "." & strExtension
        End If

    Next comComponent

    Set dictConfig = New Dictionary
    dictConfig.Add STR_CONFIGKEY_MODULEPATHS, dictModulePaths
    strConfigJson = JsonConverter.ConvertToJson(dictConfig, vbTab)

    strConfigFilePath = ConfigFilePath(prjActVBProject)
    Set FSO = New Scripting.FileSystemObject
    Set tsConfigStream = FSO.CreateTextFile(strConfigFilePath, True)
    tsConfigStream.Write strConfigJson
    tsConfigStream.Close

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


Public Sub ExportFiles()

    Dim prjActVBProject     As VBProject
    Dim strConfigFilePath   As String
    Dim strVBASourceDirPath As String
    Dim varModuleName       As Variant
    Dim strModuleName       As String
    Dim FSO                 As Scripting.FileSystemObject
    Dim tsConfigStream      As Scripting.TextStream
    Dim strConfigJson       As String
    Dim dictConfig          As Dictionary
    Dim dictModulePaths     As Dictionary
    Dim strModulePath       As String
    Dim comModuleComponent  As VBComponent

    On Error GoTo ErrHandler

    Set prjActVBProject = Application.VBE.ActiveVBProject
    If prjActVBProject Is Nothing Then Exit Sub

    strConfigFilePath = ConfigFilePath(prjActVBProject)
    Set FSO = New Scripting.FileSystemObject
    If Not FSO.FileExists(strConfigFilePath) Then
        MsgBox "You need to create file list config file before you can export files!"
        Exit Sub
    End If
    Set tsConfigStream = FSO.OpenTextFile(strConfigFilePath, ForReading)
    strConfigJson = tsConfigStream.ReadAll()
    tsConfigStream.Close
    Set dictConfig = JsonConverter.ParseJson(strConfigJson)

    strVBASourceDirPath = VBASourceDirPath(prjActVBProject)
    Set dictModulePaths = dictConfig(STR_CONFIGKEY_MODULEPATHS)
    For Each varModuleName In dictModulePaths.Keys

        strModuleName = varModuleName
        strModulePath = dictModulePaths(strModuleName)
        strModulePath = NormalisePath(strModulePath, strVBASourceDirPath)
        Set comModuleComponent = prjActVBProject.VBComponents(strModuleName)

        comModuleComponent.Export strModulePath

        If comModuleComponent.Type = vbext_ct_Document Then
            comModuleComponent.CodeModule.DeleteLines 1, comModuleComponent.CodeModule.CountOfLines
        Else
            prjActVBProject.VBComponents.Remove comModuleComponent
        End If

    Next varModuleName

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


Public Sub ImportFiles()

    Dim prjActVBProject     As VBProject
    Dim strConfigFilePath   As String
    Dim strVBASourceDirPath As String
    Dim varModuleName       As Variant
    Dim strModuleName       As String
    Dim FSO                 As Scripting.FileSystemObject
    Dim tsConfigStream      As Scripting.TextStream
    Dim strConfigJson       As String
    Dim dictConfig          As Dictionary
    Dim dictModulePaths     As Dictionary
    Dim strModulePath       As String
    Dim comNewImport        As VBComponent
    Dim comExistingComp     As VBComponent

    Dim modCodeCopy         As VBIDE.CodeModule
    Dim modCodePaste        As VBIDE.CodeModule

    On Error GoTo catchError

    Set prjActVBProject = Application.VBE.ActiveVBProject
    If Application.VBE.ActiveVBProject Is Nothing Then Exit Sub

    strConfigFilePath = ConfigFilePath(prjActVBProject)
    Set FSO = New Scripting.FileSystemObject
    If Not FSO.FileExists(strConfigFilePath) Then
        MsgBox "You need to create file list config file before you can import files!"
        Exit Sub
    End If
    Set tsConfigStream = FSO.OpenTextFile(strConfigFilePath, ForReading)
    strConfigJson = tsConfigStream.ReadAll()
    tsConfigStream.Close
    Set dictConfig = JsonConverter.ParseJson(strConfigJson)

    strVBASourceDirPath = VBASourceDirPath(prjActVBProject)
    Set dictModulePaths = dictConfig(STR_CONFIGKEY_MODULEPATHS)
    For Each varModuleName In dictModulePaths.Keys

        strModuleName = varModuleName
        strModulePath = dictModulePaths(strModuleName)
        strModulePath = NormalisePath(strModulePath, strVBASourceDirPath)

        Set comNewImport = prjActVBProject.VBComponents.Import(strModulePath)
        If comNewImport.Name <> strModuleName Then
            If CollectionKeyExists(prjActVBProject.VBComponents, strModuleName) Then

                Set comExistingComp = prjActVBProject.VBComponents(strModuleName)
                If comExistingComp.Type = vbext_ct_Document Then

                    Set modCodeCopy = comNewImport.CodeModule
                    Set modCodePaste = comExistingComp.CodeModule
                    modCodePaste.DeleteLines 1, modCodePaste.CountOfLines
                    If modCodeCopy.CountOfLines > 0 Then
                        modCodePaste.AddFromString modCodeCopy.Lines(1, modCodeCopy.CountOfLines)
                    End If
                    prjActVBProject.VBComponents.Remove comNewImport

                Else

                    prjActVBProject.VBComponents.Remove comExistingComp
                    comNewImport.Name = strModuleName

                End If
            Else

                comNewImport.Name = strModuleName

            End If

        End If

    Next varModuleName

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


'// Config file path for a given VBProject
Private Function ConfigFilePath(ByVal Project As VBProject) As String

    ConfigFilePath = ProjParentDirPath(Project) & STRCONFIGFILENAME

End Function


'// Parse path name
Private Function NormalisePath(ByVal Path As String, ByVal BaseDir As String) As String

    Dim FSO As Scripting.FileSystemObject

    Set FSO = New Scripting.FileSystemObject
    If FSO.GetDriveName(Path) = vbNullString Then
        '// Assume path is relative
        NormalisePath = FSO.BuildPath(BaseDir, Path)
    Else
        '// Assume path is absolute
        NormalisePath = Path
    End If
    NormalisePath = FSO.GetAbsolutePathName(NormalisePath)
    
End Function


'// Path of the VBA source directory for a given VBProject
Private Function VBASourceDirPath(ByVal Project As VBProject) As String

    VBASourceDirPath = ProjParentDirPath(Project)

End Function


'// The parent directory path for a given VBProject
Private Function ProjParentDirPath(ByVal Project As VBProject) As String

    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject

    ProjParentDirPath = FSO.GetParentFolderName(Project.Filename) & Application.PathSeparator

End Function


Private Function CollectionKeyExists(ByVal coll As Object, ByVal key As String) As Boolean

    On Error Resume Next
    coll (key)
    CollectionKeyExists = (Err.Number = 0)
    On Error GoTo 0

End Function
