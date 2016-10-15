VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfigure 
   Caption         =   "Configure "
   ClientHeight    =   3346
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6489
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
End Sub

'// on startup do an initial scan
Private Sub UserForm_Initialize()
    
End Sub
