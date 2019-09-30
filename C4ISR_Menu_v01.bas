Attribute VB_Name = "C4ISR_Menu"
Option Explicit
Global Const sToolbarC4ISR As String = "C4ISRRibbon"

Public Sub AddRibbonsC4ISR()
    'Add user ribbons, call it from Workbook_Open
    Call AddRibbonLineC4ISR
End Sub
Public Sub DeleteRibbonsC4ISR()
    'Delete ribbons, call it from Workbook_BeforeClose
    On Error Resume Next
    Application.CommandBars(sToolbarC4ISR).Delete
End Sub

Sub AddRibbonLineC4ISR()
    'C4ISRRibbon
    Dim cbToolBar
    
    Dim ctButton1
    Dim ctButton2
    Dim ctButton3
    Dim ctButton4
    Dim ctButton5
    Dim ctButton6
    
    On Error Resume Next
    Set cbToolBar = Application.CommandBars.Add(sToolbarC4ISR, msoBarTop, False, True)
    With cbToolBar
        Set ctButton1 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton2 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton3 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton4 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton5 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton6 = .Controls.Add(Type:=msoControlButton, ID:=2950)
    End With
    
    With ctButton1
        .Caption = "Del Scan"
        .FaceId = 2087
        .OnAction = "DeleteScannedData"
        .TooltipText = "Delete scanned data"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton2
        .Caption = "Add2INV"
        .FaceId = 2046
        .OnAction = "AddToInventory"
        .TooltipText = "Add to Inventory sheet"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton3
        .Caption = "Add2FULLINV"
        .FaceId = 2045
        .OnAction = "AddToFullInventory"
        .TooltipText = "Add to Full Inventory sheet"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton4
        .Caption = "Read File"
        .FaceId = 1947
        .OnAction = "ReadFromFile"
        .TooltipText = "Read from Scanner file"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton5
        .Caption = "HUN"
        .FaceId = 205
        .OnAction = "SwitchToHUN"
        .TooltipText = "Switch to HUN keyboard"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton6
        .Caption = "ENG"
        .FaceId = 205
        .OnAction = "SwitchToENG"
        .TooltipText = "Switch to ENGkeyboard"
        .Style = msoButtonIconAndCaption
    End With
    
    
    
    
    With cbToolBar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub
Sub MenuNULL()
    'Empty dummy subroutine
End Sub

