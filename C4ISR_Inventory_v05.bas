Attribute VB_Name = "C4ISR_Inventory"
Option Explicit
'Password to unlock page: pwd123

Private Const cmshInventory As String = "Inventory"
Private Const cmshFullInventory As String = "Full Inventory"
Private Const cmshScan As String = "Scan"
Private Const cmshCoverPage As String = "Cover Page"
Private Const cmlScanListStartRow As Long = 2
Private Const cmlScanListEndRow As Long = 5000
Private Const cmlInventoryStartRow As Long = 2
Private Const cmlInventoryEndRow As Long = 5000
Private isMacOS As Boolean
Private isWindows As Boolean
Private Is64BitOffice As Boolean
Private clLogger As New LoggerClass
Private clMotorola As New MotorolaCS3070Class
Private Const cmbDebugMode As Boolean = True

Sub C4ISR_Inventory_Start()
  'Description
  'Parameters:
  'Created by: Laszlo Tamas


  On Error GoTo PROC_ERR

  clLogger.logINFO "Program start", "C4ISR_Inventory.C4ISR_Inventory"

  '---------------
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure C4ISR_Inventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.C4ISR_Inventory"
  End If
  Resume PROC_EXIT
End Sub

Private Sub C4ISR_InventoryTest()
  'Test procedure for C4ISR_Inventory
  C4ISR_Inventory
End Sub
Private Function WantToDeleteScanned() As Boolean
  'Want ot delete scanned data
  'Returns:{Boolean}
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim bRes As Boolean
  Dim mgConfirm As Integer
  
  On Error GoTo FUNC_ERR
  
  bRes = False
  
  mgConfirm = MsgBox("Are you sure?", 36, "Confirm delete scanned data")
  Select Case mgConfirm
    Case vbYes
      bRes = True
    Case vbNo
      bRes = False
  End Select
  WantToDeleteScanned = bRes
  ' - - - - - - - - - - - - - - -
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function WantToDeleteScanned"
  If Err.Number Then
    clLogger.logERROR Err.Description, "WantToDeleteScanned"
  End If
  Resume FUNC_EXIT
End Function
Private Sub WantToDeleteScannedTest()
  'Test procedure For WantToDeleteScanned
  'Want ot delete scanned data
  clLogger.logDEBUG WantToDeleteScanned(), "WantToDeleteScannedTest"
End Sub
Private Function WanToDeleteInventory() As Boolean
  'Want ot delete inventory data
  'Returns:{Boolean}
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim bRes As Boolean
  Dim mgConfirm As Integer
  
  On Error GoTo FUNC_ERR
  
  bRes = False
  
  mgConfirm = MsgBox("Are you sure?", 36, "Confirm delete inventory data")
  Select Case mgConfirm
    Case vbYes
      bRes = True
    Case vbNo
      bRes = False
  End Select
  WanToDeleteInventory = bRes
  ' - - - - - - - - - - - - - - -
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function C4ISR_Inventory.WanToDeleteInventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.WanToDeleteInventory"
  End If
  Resume FUNC_EXIT
End Function

Private Sub WanToDeleteInventoryTest()
  'Test procedure For WanToDeleteInventory
  'Want ot delete inventory data
  clLogger.logDEBUG WanToDeleteInventory(), "WanToDeleteInventoryTest"
End Sub
  

Private Function DeleteInventory() As Boolean
  'Delete the content of Inventory sheet
  'Created by: Laszlo Tamas
  
  On Error GoTo FUNC_ERROR

FUNC_EXIT:
  On Error GoTo 0
  DeleteInventory = bDelete
  Exit Function
FUNC_ERROR:
  Debug.Print "Error in Function DeleteInventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.DeleteInventory"
  End If
  Resume FUNC_EXIT
End Function

Sub AddToFullInventory()
  On Error GoTo PROC_ERR

PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure AddToFullInventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.AddToFullInventory"
  End If
  Resume PROC_EXIT
End Sub

Sub AddToInventory()
  On Error GoTo PROC_ERR

PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure AddToInventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.AddToInventory"
  End If
  Resume PROC_EXIT
End Sub

Private Sub DeleteScannedData()
  'Delete scanned datat from sheet
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim sRange As String
  
  On Error GoTo PROC_ERR
  Sheets(cmshScan).Activate
  If WantToDeleteScanned() Then
    Sheets(cmshScan).Select
    sRange = "A2:A" & Trim(CStr(cmlScanListEndRow))
    Range(sRange).Select
    Selection.ClearContents
    Range("A2").Select
  End If
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure DeleteScannedData"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.DeleteScannedData"
  End If
  Resume PROC_EXIT
End Sub
  


Sub ReadFromFile()
  'Read scanned datat from TXT file stored On scanner
  'Created by: Laszlo Tamas
  'Licence: MIT
  Dim textdata() As String
  Dim i As Long
  
  On Error GoTo PROC_ERR
  
  If WantToDeleteScanned() Then
    textdata = clMotorola.GetTextData()
    If textdata(1) <> "" Then
      For i = 1 To UBound(textdata)
        Sheets(cmshScan).Cells(i + 1, 1) = clMotorola.GetBarcodeDataFromBarcodeLine(textdata(i))
      Next i
    Else
      clLogger.logDEBUG "File is EMPTY", "C4ISR_Inventory.ReadFromFile"
    End If
  End If
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure ReadFromFile"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.ReadFromFile"
  End If
  Resume PROC_EXIT
End Sub
  

Sub SwitchToENG()
  Call clMotorola.SwitchToENG
  
End Sub
Sub SwitchToHUN()
  Call clMotorola.SwitchToHUN
  
End Sub
