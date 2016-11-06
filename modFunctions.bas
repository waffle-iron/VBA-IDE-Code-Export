Attribute VB_Name = "modFunctions"
Option Explicit


Function fConfFileExists() As Boolean

    '// Determine if current location has a
    '// project file saved in it looking for the
    '// .conf file in the ThisWorkbook location
    '// pre populate rExportTo and rImportFrom if found

    Dim FSO     As New Scripting.FileSystemObject
    Dim file    As Scripting.file
    Dim strPath As String
    Dim strFile As String
    
    strFile = STRCONFIGFILENAME
    
    '// check to see if the config file is at the root of the project
    strPath = fAddPathSeparator(FSO.GetParentFolderName(g_ActiveVBProjectName))
    
    If FSO.FileExists(strPath & strFile) Then
'        g_blnConfigAvailable = True
'        shtConfig.Range("rComponentTXTList") = g_blnConfigAvailable
        g_strConfigFilePath = strPath & strFile
        fConfFileExists = True
        GoTo ExitFunction
    End If
        
    '// if not config
    g_blnConfigAvailable = False
    fConfFileExists = False
    
ExitFunction:
    Exit Function

CatchError:
    GoTo ExitFunction

End Function


Function fFilePicker(strPickType As String, Optional strFileSpec As String, Optional strTitle As String, _
    Optional strFilterString As String, Optional bolAllowMultiSelect As Boolean) As String

    Dim fdiBox                      As FileDialog
    Dim lngIdx                      As Long
    Dim lngCount                    As Long
    Dim varArrFilters()             As Variant
    Dim varArrFilterElements()      As Variant
    Dim strSiteName                 As String

    On Error GoTo CatchError
   
    Select Case LCase(strPickType)
        Case "file"
            Set fdiBox = Application.FileDialog(msoFileDialogFilePicker)
        Case "folder"
            Set fdiBox = Application.FileDialog(msoFileDialogFolderPicker)
    End Select
    
    With fdiBox
        .InitialFileName = strFileSpec
        .AllowMultiSelect = bolAllowMultiSelect
        
        If strTitle <> "" Then
            .Title = strTitle
        End If
        
        .Filters.Clear
        
        If strFilterString <> "" Then
            varArrFilters = Split(strFilterString, "|")
            
            For lngIdx = LBound(varArrFilters) To UBound(varArrFilters)
                varArrFilterElements = Split(varArrFilters(lngIdx), ",")
                
                .Filters.Add varArrFilterElements(0), "*." & varArrFilterElements(1)
            Next
        End If

        If .Show = -1 Then

            For lngIdx = 1 To .SelectedItems.Count
                If lngIdx > 1 Then
                    fFilePicker = fFilePicker & "|"
                End If
                    
                fFilePicker = fFilePicker & fConvToUNC(CStr(.SelectedItems(lngIdx)))
            Next
        End If
    End With
    
    'Set the object variable to Nothing.
    Set fdiBox = Nothing

ExitFunction:
    Exit Function

CatchError:
    GoTo ExitFunction
    
End Function


Function fConvToUNC(strPath As String) As String
        
    '// converts a URL to a UNC path adding the @SSL where required for SharePoint
    
    If LCase(Left(strPath, 4)) = "http" Then
    
        If InStr(1, strPath, "https://") Then
            strPath = Replace(strPath, "https://", "")
            strPath = Replace(strPath, "/", "@SSL\", , 1)
        ElseIf InStr(1, strPath, "https:\\") Then
            strPath = Replace(strPath, "https:\\", "")
            strPath = Replace(strPath, "\", "@SSL\", , 1)
        ElseIf InStr(1, strPath, "http://") Then
            strPath = Replace(strPath, "http://", "")
        ElseIf InStr(1, strPath, "http:\\") Then
            strPath = Replace(strPath, "http:\\", "")
        End If
        
        strPath = "\\" & Replace(strPath, "/", "\")
        '// added to cater for spaces
        strPath = Replace(strPath, "%20", " ")
    End If
    
    fConvToUNC = strPath

End Function


Function fComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
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


Sub fSearchCodeModule(strComponentName As String, strFindWhat As String)
    
    '// used to search for Workbook_ or Worksheet_
    '// to help determine what to do with the code
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim FindWhat As String
    Dim SL As Long ' start line
    Dim EL As Long ' end line
    Dim SC As Long ' start column
    Dim EC As Long ' end column
    Dim Found As Boolean
    
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(strComponentName)
    Set CodeMod = VBComp.CodeModule
    
    With CodeMod
        SL = 1
        EL = .CountOfLines
        SC = 1
        EC = 255
        Found = .Find(Target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
            EndLine:=EL, EndColumn:=EC, _
            wholeword:=True, MatchCase:=False, patternsearch:=False)
        Do Until Found = False
            Debug.Print "Found at: Line: " & CStr(SL) & " Column: " & CStr(SC)
            EL = .CountOfLines
            SC = EC + 1
            EC = 255
            Found = .Find(Target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
                EndLine:=EL, EndColumn:=EC, _
                wholeword:=True, MatchCase:=False, patternsearch:=False)
        Loop
    End With
End Sub


Function fFSOTextStream(FSO As Scripting.FileSystemObject) As Scripting.TextStream

    '// create the file
    Set fFSOTextStream = FSO.CreateTextFile(FSO.GetParentFolderName(g_ActiveVBProjectName) & Application.PathSeparator & STRCONFIGFILENAME)
    
    '// open the .conf file
    Set fFSOTextStream = FSO.OpenTextFile(g_strConfigFilePath, ForReading)

End Function


Function fAddPathSeparator(strPath As String)
    
    If Not Right(strPath, 1) = Application.PathSeparator Then
        strPath = strPath & Application.PathSeparator
        fAddPathSeparator = strPath
    End If
    
End Function




    
