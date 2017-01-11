Attribute VB_Name = "modMenu"
Option Explicit

'// Add references for :
'//     Microsoft Visual Basic For Applications Extensibility 5.3
'//     Microsoft Scripting Runtime
'// Also check the 'Trust access to the VBA project model check box', located...
'// Trust Centre, Trust Centre Settings, Macro Settings, Trust access to the VBA project model

Private MnuEvt      As clsVBECmdHandler
Private EvtHandlers As New Collection


Public Sub auto_open()

    Call CreateVBEMenu
    '// Call CreateXLMenu

End Sub


Public Sub auto_close()

    Call RemoveVBEMenu

End Sub


Private Sub CreateVBEMenu()

    Dim objMenu     As CommandBarPopup
    Dim objMenuItem As Object

    Set objMenu = Application.VBE.CommandBars(1).Controls.Add(Type:=msoControlPopup)
    With objMenu
        objMenu.Caption = "E&xport for VCS"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "MakeFileList"
        Call MenuEvents(objMenuItem)
        objMenuItem.Caption = "&Make File List"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "ImportFiles"
        Call MenuEvents(objMenuItem)
        objMenuItem.Caption = "&Import Files"

        Set objMenuItem = .Controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "ExportFiles"
        Call MenuEvents(objMenuItem)
        objMenuItem.Caption = "&Export Files"

    End With

    Set objMenuItem = Nothing
    Set objMenu = Nothing

End Sub


Private Sub MenuEvents(ByVal objMenuItem As Object)

    Set MnuEvt = New clsVBECmdHandler
    Set MnuEvt.EvtHandler = Application.VBE.Events.CommandBarEvents(objMenuItem)
    EvtHandlers.Add MnuEvt

End Sub


Private Sub CreateXLMenu()

    MenuBars(xlWorksheet).Menus.Add Caption:="E&xport for VCS"
    With MenuBars(xlWorksheet).Menus("Export for VCS").MenuItems
        .Add Caption:="&Make File List", _
             OnAction:="MakeFileList"
        .Add Caption:="&Import Files", _
             OnAction:="ImportFiles"
        .Add Caption:="&Export Files", _
             OnAction:="ExportFiles"
    End With

End Sub


Private Sub RemoveVBEMenu()

    On Error Resume Next

    Application.VBE.CommandBars(1).Controls("Export for VCS").Delete

    '// Clear the EvtHandlers collection if there is anything in it
    While EvtHandlers.Count > 0
        EvtHandlers.Remove 1
    Wend

    Set EvtHandlers = Nothing
    Set MnuEvt = Nothing

    Application.CommandBars("Worksheet Menu Bar").Controls("E&xport for VCS").Delete
    On Error GoTo 0

End Sub
