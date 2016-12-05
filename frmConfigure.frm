VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfigure 
   Caption         =   "Configure "
   ClientHeight    =   3346
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   6495
   OleObjectBlob   =   "frmConfigure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkInternalComponents_Click()
    shtConfig.Range("rComponentTXTList") = chkInternalComponents.Value
    g_blnMakeConfFile = chkInternalComponents.Value
End Sub

Private Sub cmdExportLocation_Click()
    txtExportTo = fAddPathSeparator(fFilePicker("folder", , "please select export location."))
    shtConfig.Range("rExportTo") = txtExportTo
End Sub

Private Sub cmdImportLocation_Click()
    txtImportFrom = fAddPathSeparator(fFilePicker("folder", , "Please select import location."))
    shtConfig.Range("rImportFrom") = txtImportFrom
End Sub

'// on startup do an initial scan
Private Sub UserForm_Initialize()
    
    Dim FSO As New Scripting.FileSystemObject
    
    chkInternalComponents.Value = shtConfig.Range("rComponentTXTList")
  
    If g_blnConfigAvailable Then
        txtExportTo = g_strExportTo
        txtImportFrom = g_strImportFrom
    Else
        txtExportTo = fAddPathSeparator(FSO.GetParentFolderName(g_strActiveVBProjectName))
        txtImportFrom = fAddPathSeparator(FSO.GetParentFolderName(g_strActiveVBProjectName))
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim FSO     As New Scripting.FileSystemObject
    Dim fsoFile As Scripting.TextStream
    
    '// if .conf available then write g_strExportTo, g_strImportFrom
    If Not fConfFileExists And g_blnMakeConfFile Then
        
        '// create the file
        Set fsoFile = FSO.CreateTextFile(fAddPathSeparator(FSO.GetParentFolderName(g_strActiveVBProjectName)) & STRCONFIGFILENAME)
        
        '// Add import and export locations
        fsoFile.WriteLine "ImportFrom:" & shtConfig.Range("rImportFrom")
        fsoFile.WriteLine "ExportTo:" & shtConfig.Range("rExportTo")
        
    ElseIf fConfFileExists And g_blnMakeConfFile Then
        
        Call UpdateFile(g_strConfigFilePath, g_strImportFrom, txtImportFrom)
        Call UpdateFile(g_strConfigFilePath, g_strExportTo, txtExportTo)
    
    Else
        '// if g_strExportTo and g_strImportFrom are not same as ThisWorkbook then
        
    End If
    
    g_blnMakeConfFile = shtConfig.Range("rComponentTXTList")
    
End Sub
