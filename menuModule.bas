Attribute VB_Name = "menuModule"
Option Explicit

Dim MnuEvt As VBECmdHandler
Dim EvtHandlers As New Collection

Sub auto_open()
    Call CreateVBEMenu
    'Call CreateXLMenu
End Sub

Sub auto_close()
    Call RemoveVBEMenu
End Sub

Sub CreateVBEMenu()
    Dim objMenu As CommandBarPopup
    Dim objMenuItem As Object

    Set objMenu = Application.VBE.CommandBars(1).Controls.Add(Type:=msoControlPopup)
    With objMenu
        objMenu.Caption = "E&xport for TFS"

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

Sub MenuEvents(objMenuItem As Object)
    Set MnuEvt = New VBECmdHandler
    Set MnuEvt.EvtHandler = Application.VBE.Events.CommandBarEvents(objMenuItem)
    EvtHandlers.Add MnuEvt
End Sub

Sub CreateXLMenu()
    MenuBars(xlWorksheet).Menus.Add Caption:="E&xport for TFS"
    With MenuBars(xlWorksheet).Menus("Export for TFS").MenuItems
        .Add Caption:="&Make File List", _
             OnAction:="MakeFileList"
        .Add Caption:="&Import Files", _
             OnAction:="ImportFiles"
        .Add Caption:="&Export Files", _
             OnAction:="ExportFiles"
    End With
End Sub

Sub RemoveVBEMenu()
    On Error Resume Next

    Application.VBE.CommandBars(1).Controls("Export for TFS").Delete

    'Clear the EvtHandlers collection if there is anything in it
    While EvtHandlers.Count > 0
        EvtHandlers.Remove 1
    Wend

    Set EvtHandlers = Nothing
    Set MnuEvt = Nothing
    
    Application.CommandBars("Worksheet Menu Bar").Controls("E&xport for TFS").Delete
    On Error GoTo 0

End Sub
