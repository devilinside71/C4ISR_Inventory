Attribute VB_Name = "C4ISR_Menu"
Option Explicit
Global Const sToolbarC4ISR As String = "C4ISRRibbon"
Global Const sToolbarC4ISRFile As String = "C4ISRFileRibbon"
Global Const sToolbarLang As String = "LanguageRibbon"

Public Sub AddRibbonsC4ISR()
    'Add user ribbons, call it from Workbook_Open
    Call AddRibbonLineC4ISR
    Call AddRibbonLineC4ISRFile
    Call AddRibbonLineLang
End Sub
Public Sub DeleteRibbonsC4ISR()
    'Delete ribbons, call it from Workbook_BeforeClose
    On Error Resume Next
    Application.CommandBars(sToolbarC4ISR).Delete
    Application.CommandBars(sToolbarC4ISRFile).Delete
    Application.CommandBars(sToolbarLang).Delete
End Sub

Sub AddRibbonLineC4ISR()
    'C4ISRRibbon
    Dim cbToolBar
    
    Dim ctButton1
    Dim ctButton2
    Dim ctButton3
    Dim ctButton4
    
    On Error Resume Next
    Set cbToolBar = Application.CommandBars.Add(sToolbarC4ISR, msoBarTop, False, True)
    With cbToolBar
        Set ctButton1 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton2 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton3 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton4 = .Controls.Add(Type:=msoControlButton, ID:=2950)
    End With
    
    With ctButton1
        .Caption = "Del Scan"
        .FaceId = 2087
        .OnAction = "DeleteScannedData"
        .TooltipText = "Delete scanned data"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton2
        .Caption = "ADD"
        .FaceId = 535
        .OnAction = "Add2FullInventoryAndInventory"
        .TooltipText = "Add to both FullInventory and Inventory sheets"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton3
        .Caption = "Add2INV"
        .FaceId = 2046
        .OnAction = "AddToInventory"
        .TooltipText = "Add to Inventory sheet"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton4
        .Caption = "Add2FULLINV"
        .FaceId = 2045
        .OnAction = "AddToFullInventory"
        .TooltipText = "Add to Full Inventory sheet"
        .Style = msoButtonIconAndCaption
    End With
    
    
    With cbToolBar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub
Sub AddRibbonLineLang()
    'LangRibbon
    Dim cbToolBar
    
    Dim ctButton1
    Dim ctButton2
    Dim ctButton3
    
    On Error Resume Next
    Set cbToolBar = Application.CommandBars.Add(sToolbarLang, msoBarTop, False, True)
    With cbToolBar
        Set ctButton1 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton2 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton3 = .Controls.Add(Type:=msoControlButton, ID:=2950)
    End With
    
    
    With ctButton1
        .Caption = "HUN"
        .FaceId = 205
        .OnAction = "SwitchToHUN"
        .TooltipText = "Switch to HUN keyboard"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton2
        .Caption = "ENG"
        .FaceId = 205
        .OnAction = "SwitchToENG"
        .TooltipText = "Switch to ENG keyboard"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton3
        .Caption = "FRA"
        .FaceId = 205
        .OnAction = "SwitchToFRA"
        .TooltipText = "Switch to FRA keyboard"
        .Style = msoButtonIconAndCaption
    End With
    
    
    
    
    With cbToolBar
        .Visible = True
        .Protection = msoBarNoChangeVisible
    End With
End Sub
Sub AddRibbonLineC4ISRFile()
    'C4ISRFileRibbon
    Dim cbToolBar
    
    Dim ctButton1
    Dim ctButton2
    
    On Error Resume Next
    Set cbToolBar = Application.CommandBars.Add(sToolbarC4ISRFile, msoBarTop, False, True)
    With cbToolBar
        Set ctButton1 = .Controls.Add(Type:=msoControlButton, ID:=2950)
        Set ctButton2 = .Controls.Add(Type:=msoControlButton, ID:=2950)
    End With
    
    
    With ctButton1
        .Caption = "Read Motorola File"
        .FaceId = 2603
        .OnAction = "ReadFromFile"
        .TooltipText = "Read from Motorola Scanner file"
        .Style = msoButtonIconAndCaption
    End With
    
    With ctButton2
        .Caption = "Read M3 File"
        .FaceId = 960
        .OnAction = "ReadFromM3File"
        .TooltipText = "Read from M3 mobile compia handheld PC file"
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

