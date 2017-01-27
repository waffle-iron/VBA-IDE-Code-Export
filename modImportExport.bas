Attribute VB_Name = "modImportExport"
Option Explicit

'// Add references for :
'//     Microsoft Visual Basic For Applications Extensibility 5.3
'//     Microsoft Scripting Runtime
'// Also check the 'Trust access to the VBA project model check box', located...
'// Trust Centre, Trust Centre Settings, Macro Settings, Trust access to the VBA project model

Private Const STRCONFIGFILENAME         As String = "CodeExport.config.json"

Private Const STR_CONFIGKEY_MODULEPATHS             As String = "Module Paths"
Private Const STR_CONFIGKEY_REFERENCES              As String = "References"
Private Const STR_CONFIGKEY_REFERENCE_NAME          As String = "Name"
Private Const STR_CONFIGKEY_REFERENCE_DESCRIPTION   As String = "Description"
Private Const STR_CONFIGKEY_REFERENCE_GUID          As String = "GUID"
Private Const STR_CONFIGKEY_REFERENCE_MAJOR         As String = "Major"
Private Const STR_CONFIGKEY_REFERENCE_MINOR         As String = "Minor"
Private Const STR_CONFIGKEY_REFERENCE_PATH          As String = "Path"

Private Const ForReading                As Integer = 1


Public Sub MakeConfigFile()

    Dim prjActProj          As VBProject

    Dim dictConfig          As Dictionary
    Dim dictModulePaths     As Dictionary
    Dim collReferences      As Collection
    Dim dictReferenceConfig As Dictionary

    Dim comModule           As VBComponent
    Dim strFileExt          As String
    Dim refReference        As Reference

    On Error GoTo catchError

    Set prjActProj = Application.VBE.ActiveVBProject
    If prjActProj Is Nothing Then GoTo exitSub

    Set dictConfig = New Dictionary

    '// Collect the name of each module, form, etc.
    Set dictModulePaths = New Dictionary
    For Each comModule In prjActProj.VBComponents

        strFileExt = vbNullString
        Select Case comModule.Type
            Case vbext_ct_StdModule
                strFileExt = "bas"
            Case vbext_ct_ClassModule, vbext_ct_Document
                strFileExt = "cls"
            Case vbext_ct_MSForm
                strFileExt = "frm"
        End Select

        If Not strFileExt = vbNullString Then
            If Not ModuleEmpty(comModule) Then
                dictModulePaths.Add comModule.Name, comModule.Name & "." & strFileExt
            End If
        End If

    Next comModule
    dictConfig.Add STR_CONFIGKEY_MODULEPATHS, dictModulePaths

    '// Collect the references
    Set collReferences = New Collection
    For Each refReference In prjActProj.References

        If Not refReference.BuiltIn Then
            Set dictReferenceConfig = New Dictionary
            dictReferenceConfig.Add STR_CONFIGKEY_REFERENCE_NAME, refReference.Name
            dictReferenceConfig.Add STR_CONFIGKEY_REFERENCE_DESCRIPTION, refReference.Description

            If refReference.Type = vbext_rk_TypeLib Then
                dictReferenceConfig.Add STR_CONFIGKEY_REFERENCE_GUID, refReference.GUID
                dictReferenceConfig.Add STR_CONFIGKEY_REFERENCE_MAJOR, refReference.Major
                dictReferenceConfig.Add STR_CONFIGKEY_REFERENCE_MINOR, refReference.Minor
            Else
                dictReferenceConfig.Add STR_CONFIGKEY_REFERENCE_PATH, refReference.FullPath
            End If

            collReferences.Add dictReferenceConfig
        End If

    Next refReference
    dictConfig.Add STR_CONFIGKEY_REFERENCES, collReferences

    '// Write config to JSON config file
    WriteConfigFile prjActProj, dictConfig

exitSub:
    Exit Sub

catchError:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub


Public Sub Export()

    Dim prjActProj          As VBProject

    Dim dictConfig          As Dictionary
    Dim dictModulePaths     As Dictionary
    Dim collConfigRefs      As Collection
    Dim dictDeclaredRef     As Dictionary

    Dim varModuleName       As Variant
    Dim strModuleName       As String
    Dim strModulePath       As String
    Dim comModule           As VBComponent

    On Error GoTo ErrHandler

    Set prjActProj = Application.VBE.ActiveVBProject
    If prjActProj Is Nothing Then GoTo exitSub

    '// Read config file and parse it to construct the Config object.
    Set dictConfig = ReadConfigFile(prjActProj)

    If dictConfig.Exists(STR_CONFIGKEY_MODULEPATHS) Then
        '// Export each module listed in the module paths to it's designated location
        Set dictModulePaths = dictConfig(STR_CONFIGKEY_MODULEPATHS)
        For Each varModuleName In dictModulePaths.Keys
    
            strModuleName = varModuleName
            strModulePath = dictModulePaths(strModuleName)
            strModulePath = EvaluatePath(prjActProj, strModulePath)
            Set comModule = prjActProj.VBComponents(strModuleName)
    
            comModule.Export strModulePath
    
            If comModule.Type = vbext_ct_Document Then
                comModule.CodeModule.DeleteLines 1, comModule.CodeModule.CountOfLines
            Else
                prjActProj.VBComponents.Remove comModule
            End If
    
        Next varModuleName
    End If

    If dictConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        '// For each reference listed in the config file, delete the references in the project
        Set collConfigRefs = dictConfig(STR_CONFIGKEY_REFERENCES)
        For Each dictDeclaredRef In collConfigRefs
    
            If CollectionKeyExists(prjActProj.References, dictDeclaredRef(STR_CONFIGKEY_REFERENCE_NAME)) Then
                prjActProj.References.Remove prjActProj.References(dictDeclaredRef(STR_CONFIGKEY_REFERENCE_NAME))
            End If
    
        Next dictDeclaredRef
    End If

exitSub:
    Exit Sub

ErrHandler:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub


Public Sub Import()

    Dim prjActProj          As VBProject

    Dim dictConfig          As Dictionary
    Dim dictModulePaths     As Dictionary
    Dim collConfigRefs      As Collection
    Dim dictDeclaredRef     As Dictionary

    Dim varModuleName       As Variant
    Dim strModuleName       As String
    Dim strModulePath       As String

    On Error GoTo catchError

    Set prjActProj = Application.VBE.ActiveVBProject
    If Application.VBE.ActiveVBProject Is Nothing Then GoTo exitSub

    Set dictConfig = ReadConfigFile(prjActProj)

    If dictConfig.Exists(STR_CONFIGKEY_MODULEPATHS) Then
        '// For each module path declared in the config file, import that module
        '// overwritting any existing modules.
        Set dictModulePaths = dictConfig(STR_CONFIGKEY_MODULEPATHS)
        For Each varModuleName In dictModulePaths.Keys

            strModuleName = varModuleName
            strModulePath = EvaluatePath(prjActProj, dictModulePaths(strModuleName))
            ImportModule prjActProj, strModuleName, strModulePath

        Next varModuleName
    End If

    If dictConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        '// Add each reference declared in the config file
        Set collConfigRefs = dictConfig(STR_CONFIGKEY_REFERENCES)
        For Each dictDeclaredRef In collConfigRefs

            If CollectionKeyExists(prjActProj.References, dictDeclaredRef(STR_CONFIGKEY_REFERENCE_NAME)) Then
                prjActProj.References.Remove prjActProj.References(dictDeclaredRef(STR_CONFIGKEY_REFERENCE_NAME))
            End If

            If dictDeclaredRef.Exists(STR_CONFIGKEY_REFERENCE_GUID) Then
                prjActProj.References.AddFromGuid _
                    dictDeclaredRef(STR_CONFIGKEY_REFERENCE_GUID), _
                    dictDeclaredRef(STR_CONFIGKEY_REFERENCE_MAJOR), _
                    dictDeclaredRef(STR_CONFIGKEY_REFERENCE_MINOR)
            Else
                prjActProj.References.AddFromFile dictDeclaredRef(STR_CONFIGKEY_REFERENCE_PATH)
            End If

        Next dictDeclaredRef
    End If

exitSub:
    Exit Sub

catchError:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub


'// Import a VBA code module... how hard could it be right?
Private Sub ImportModule(ByVal Project As VBProject, ByVal ModuleName As String, ByVal ModulePath As String)

    Dim comNewImport        As VBComponent
    Dim comExistingComp     As VBComponent
    Dim modCodeCopy         As CodeModule
    Dim modCodePaste        As CodeModule

    Set comNewImport = Project.VBComponents.Import(ModulePath)
    If comNewImport.Name <> ModuleName Then
        If CollectionKeyExists(Project.VBComponents, ModuleName) Then

            Set comExistingComp = Project.VBComponents(ModuleName)
            If comExistingComp.Type = vbext_ct_Document Then

                Set modCodeCopy = comNewImport.CodeModule
                Set modCodePaste = comExistingComp.CodeModule
                modCodePaste.DeleteLines 1, modCodePaste.CountOfLines
                If modCodeCopy.CountOfLines > 0 Then
                    modCodePaste.AddFromString modCodeCopy.Lines(1, modCodeCopy.CountOfLines)
                End If
                Project.VBComponents.Remove comNewImport

            Else

                Project.VBComponents.Remove comExistingComp
                comNewImport.Name = ModuleName

            End If
        Else

            comNewImport.Name = ModuleName

        End If
    End If

End Sub


Private Function ModuleEmpty(ByVal comModule As VBComponent) As Boolean

    Dim lngNumLines As Long
    Dim lngCurLine As Long
    Dim strCurLine As String

    ModuleEmpty = True

    lngNumLines = comModule.CodeModule.CountOfLines
    For lngCurLine = 1 To lngNumLines
        strCurLine = comModule.CodeModule.Lines(lngCurLine, 1)
        If Not (strCurLine = "Option Explicit" Or strCurLine = "") Then
            ModuleEmpty = False
            Exit Function
        End If
    Next lngCurLine

End Function


'// Read an parse the config file for a project
Private Function ReadConfigFile(ByVal Project As VBProject) As Dictionary

    Dim strConfigFilePath   As String
    Dim tsConfigStream      As Scripting.TextStream
    Dim strConfigJson       As String
    Dim FSO                 As Scripting.FileSystemObject

    Set FSO = New Scripting.FileSystemObject

    strConfigFilePath = ConfigFilePath(Project)
    If Not FSO.FileExists(strConfigFilePath) Then
        MsgBox "You need to create file list config file before you can import or export files!"
        Exit Function
    End If
    Set tsConfigStream = FSO.OpenTextFile(strConfigFilePath, ForReading)
    strConfigJson = tsConfigStream.ReadAll()
    tsConfigStream.Close
    Set ReadConfigFile = JsonConverter.ParseJson(strConfigJson)

End Function


'// Write a configuration to the config file for a project
Private Sub WriteConfigFile(ByVal Project As VBProject, ByVal Config As Dictionary)

    Dim FSO                 As Scripting.FileSystemObject
    Dim tsConfigStream      As Scripting.TextStream
    Dim strConfigFilePath   As String
    Dim strConfigJson       As String

    Set FSO = New Scripting.FileSystemObject

    strConfigJson = JsonConverter.ConvertToJson(Config, vbTab)
    strConfigFilePath = ConfigFilePath(Project)
    Set tsConfigStream = FSO.CreateTextFile(strConfigFilePath, True)
    tsConfigStream.Write strConfigJson
    tsConfigStream.Close

End Sub


'// Config file path for a given VBProject
Private Function ConfigFilePath(ByVal Project As VBProject) As String

    ConfigFilePath = ProjParentDirPath(Project) & STRCONFIGFILENAME

End Function


'// Parse a path name
Private Function EvaluatePath(ByVal Project As VBProject, ByVal Path As String) As String

    Dim FSO         As Scripting.FileSystemObject
    Dim BaseDir     As String

    Set FSO = New Scripting.FileSystemObject

    '// Tack on the BaseDir if the Path is relative
    BaseDir = SourceDirPath(Project)
    If FSO.GetDriveName(Path) = vbNullString Then
        '// Assume path is relative
        EvaluatePath = FSO.BuildPath(BaseDir, Path)
    Else
        '// Assume path is absolute
        EvaluatePath = Path
    End If

    '// Resolve any parts of the path such as '..' and '.'
    EvaluatePath = FSO.GetAbsolutePathName(EvaluatePath)

End Function


'// Path of the VBA source directory for a given VBProject
Private Function SourceDirPath(ByVal Project As VBProject) As String

    SourceDirPath = ProjParentDirPath(Project)

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

Private Function HandleCrash(ByVal ErrNumber As Long, ByVal ErrDesc As String, ByVal ErrSource As String) As Boolean

    Dim UserAction As Integer

    UserAction = MsgBox( _
        Prompt:="An unexpected problem occured, would you like to debug?", _
        Buttons:=vbYesNo, _
        Title:="Unexpected problem")

    HandleCrash = UserAction = vbYes

End Function
