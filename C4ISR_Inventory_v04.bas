Attribute VB_Name = "C4ISR_Inventory"
Option Explicit
'Password to unlock page: pwd123
#If Win64 Then
  Private Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As LongPtr, ByVal Flag As Long) As LongPtr
#Else
  #If Win32 Then
    Private Declare Function ActivateKeyboardLayout Lib "user32.dll" (ByVal HKL As Long, Flag As Long) As Long
  #Else
    'Mac
  #End If
#End If

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
Private Const cmbDebugMode As Boolean = True

Sub C4ISR_Inventory()
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


  On Error GoTo PROC_ERR

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
  
  On Error GoTo PROC_ERR
  
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
  'Switch to ENG keyboard
  #If Mac Then
  #Else
    ActivateKeyboardLayout 1033, 0
  #End If
End Sub

Sub SwitchToHUN()
  'Switch to HUN keyboard
  #If Mac Then
  #Else
    ActivateKeyboardLayout 1038, 0
  #End If
End Sub

