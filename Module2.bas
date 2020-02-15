Attribute VB_Name = "Module2"
Sub Makr�1()
Attribute Makr�1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makr�1 Makr�
'

'
    ActiveSheet.Shapes.Range(Array("Flag")).Select
End Sub

Private Sub Add2FullInventoryAndInventory(Parameter As String)
  'Add both to FullInventory and Inventory
  'Parameters:
  '           {String} Parameter
  'Created by: Laszlo Tamas
  'Licence: MIT

  On Error GoTo PROC_ERR

  'Code here

  '---------------
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Sub Add2FullInventoryAndInventory"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "Add2FullInventoryAndInventory")
  End If
  Resume PROC_EXIT
End Sub
Private Sub Add2FullInventoryAndInventoryTest
  'Test procedure for Add2FullInventoryAndInventory
  'Add both to FullInventory and Inventory
  Dim testVal As String
  Dim dtmStartTime As Date
  testVal = Nothing
  dtmStartTime = Now()
  Call Add2FullInventoryAndInventory(testVal)
End Sub
  
  

