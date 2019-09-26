Attribute VB_Name = "C4ISR_Inventory"
Option Explicit
#If Win64 Then
    Private Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As LongPtr, ByVal Flag As Long) As LongPtr
#Else
    Private Declare Function ActivateKeyboardLayout Lib "user32.dll" (ByVal HKL As Long, Flag As Long) As Long
#End If
Sub C4ISR_Inventory()
    'Description
    'Parameters:
    'Created by: Laszlo Tamas


    On Error GoTo PROC_ERR

    'Code here

    '---------------
PROC_EXIT:
    On Error GoTo 0
    Exit Sub
PROC_ERR:
    Debug.Print "Error in Procedure C4ISR_Inventory"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume PROC_EXIT
End Sub
Private Sub C4ISR_InventoryTest()
    'Test procedure for C4ISR_Inventory
    Dim dtmStartTime As Date



    dtmStartTime = Now()
    Call C4ISR_Inventory
End Sub
Private Sub Add2FullInventory(Path2ADSInventoryFile As String)
  'Add data to Full Inventory
  'Parameters:
  ' {String} Path2ADSInventoryFile
  'Created by: Laszlo Tamas
  Dim sPath2Inventory As String
  Dim FileNum As Integer
  Dim DataLine As String
  Dim sPath As String
  Dim iChoice As Integer
  Dim wbkADSInventory As Workbook
  Dim wbkCurr As Workbook
  Dim lLastRowFullInventory As Long
  Dim ws As Worksheet
  Dim pos As Integer
  Dim shInventorySheet As String
  Dim sPackageName As String
  Dim i As Long
  Dim k As Long
  Dim sCell As String
  Dim mgConfirm As Integer
  Dim bNotInTheList As Boolean
  
  On Error GoTo PROC_ERR
  Set wbkCurr = Application.ActiveWorkbook
  lLastRowFullInventory = GetLastRow("Full Inventory", 1, 2, 50000, False)
  Debug.Print "Last row of Full Inventory sheet: " & lLastRowFullInventory
  
  
  
  shInventorySheet = vbNullString
  
  Set wbkADSInventory = Workbooks.Open(Filename:=Path2ADSInventoryFile)
  For Each ws In wbkADSInventory.Worksheets
    ' Debug.Print ws.Name
    pos = InStr(ws.Name, "Inventory")
    If pos <> 0 Then
      shInventorySheet = ws.Name
    End If
  Next
  Debug.Print "Inventory sheet name: " & shInventorySheet
  
  sPackageName = ScanForPackageName(wbkADSInventory)
  Debug.Print "package name: " & sPackageName
  
  ' Check if package is in the Full Inventory already
  bNotInTheList = True
  For i = 2 To lLastRowFullInventory
    sCell = Trim(CStr(wbkCurr.Sheets("Full Inventory").Cells(i, 1)))
    If sCell = sPackageName Then
      mgConfirm = MsgBox("Are you sure to continue?", 36, sPackageName & " is already in the list")
      Select Case mgConfirm
          Case vbYes
              bNotInTheList = True
          Case vbNo
              bNotInTheList = False
      End Select
      Exit For
    End If
  Next i
  If Not bNotInTheList Then
    Debug.Print "Interrupted because " & sPackageName & " is in the list."
    GoTo PROC_EXIT
  End If
  
  'Copy data
  For i = 2 To 3000
    lLastRowFullInventory = lLastRowFullInventory + 1
    sCell = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, 1)))
    If sCell = "" Then
      Exit For
    End If
    wbkCurr.Sheets("Full Inventory").Cells(lLastRowFullInventory, 1) = sPackageName
    For k = 1 To 8
      wbkCurr.Sheets("Full Inventory").Cells(lLastRowFullInventory, k + 1) = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, k)))
    Next k
  Next i
  
  '---------------
PROC_EXIT:
  On Error GoTo 0
  wbkADSInventory.Close
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure Add2FullInventory"
  If Err.Number Then
    Debug.Print Err.Description
  End If
  Resume PROC_EXIT
End Sub


Private Sub Add2FullInventoryTest()
    'Test procedure for Add2FullInventory
    Dim dtmStartTime As Date
    Dim sPath2ADSInventoryFile As String

    sPath2ADSInventoryFile = GetPath2Inventory()


    dtmStartTime = Now()
    Call Add2FullInventory(sPath2ADSInventoryFile)
End Sub
Sub AddToFullInventory()
    'Add2FullInventory
    Dim dtmStartTime As Date
    Dim sPath2ADSInventoryFile As String

    sPath2ADSInventoryFile = GetPath2Inventory()


    dtmStartTime = Now()
    Call Add2FullInventory(sPath2ADSInventoryFile)
    MsgBox "Finished!"
End Sub
Private Function GetPath2Inventory() As String
    'Open file dialog to determine ADS inventory filename
    'Parameters:
    'Returns:{String}
    'Created by: Laszlo Tamas

    Dim sRes As String
    Dim FileNum As Integer
    Dim DataLine As String
    Dim sPath As String
    Dim iChoice As Integer
    
    On Error GoTo FUNC_ERR

    sPath = vbNullString

    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    'make the file dialog visible to the user
    iChoice = Application.FileDialog(msoFileDialogOpen).Show
    'determine what choice the user made
    If iChoice <> 0 Then
        'get the file path selected by the user
        sPath = Application.FileDialog( _
        msoFileDialogOpen).SelectedItems(1)
'        FileNum = FreeFile()
'        Open sPath For Input As #FileNum
'
'        While Not EOF(FileNum)
'            Line Input #FileNum, DataLine ' read in data 1 line at a time
'            ' decide what to do with dataline,
'            ' depending on what processing you need to do for each case
'        Wend
    End If
    sRes = sPath
    

    GetPath2Inventory = sRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function GetPath2Inventory"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function
Private Sub GetPath2InventoryTest()
    'Test procedure for GetPath2Inventory
    'Open file dialog to determine ADS inventory filename
    Dim dtmStartTime As Date



    dtmStartTime = Now()
    Debug.Print "Function GetPath2Inventory test: >> " & GetPath2Inventory()
End Sub
Private Function ScanForPackageName(InvetoryBook As Workbook) As String
  'Scan for Package name
  'Parameters:
  ' {Workbook} InvetoryBook
  'Returns:{String}
  'Created by: Laszlo Tamas
  Dim pos As Integer
  Dim sRes As String
  Dim i As Long
  Dim sCell As String
  
  On Error GoTo FUNC_ERR
  
  sRes = "text"
  'Code here
  For i = 1 To 30
    sCell = Trim(CStr(InvetoryBook.Sheets("Cover Page").Cells(i, 1)))
    pos = InStr(sCell, "INVENTORY")
    If pos <> 0 Then
      sRes = Trim(CStr(Replace(sCell, "INVENTORY", "")))
      Exit For
    End If
  Next i
  
  
  ScanForPackageName = sRes
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function ScanForPackageName"
  If Err.Number Then
    Debug.Print Err.Description
  End If
  Resume FUNC_EXIT
End Function


Private Sub ScanForPackageNameTest()
  'Test procedure for ScanForPackageName
  'Scan for Package name
  Dim dtmStartTime As Date
  Dim wbkInvetoryBook As Workbook
  Dim sPath2ADSInventoryFile As String
  
  ' wbkInvetoryBook = ActiveWorkbook
  sPath2ADSInventoryFile = GetPath2Inventory()
  Set wbkInvetoryBook = Workbooks.Open(Filename:=sPath2ADSInventoryFile)
  
  dtmStartTime = Now()
  Debug.Print "Function ScanForPackageName test: >> " & ScanForPackageName(wbkInvetoryBook)
  wbkInvetoryBook.Close

End Sub


'----------------
'Columns and Rows
'----------------
Private Function Col_Letter(lngCol As Long) As String
  'Get letter from column number
  Dim vArr
  
  ' On Error Resume Next
  vArr = Split(Cells(1, lngCol).Address(True, False), "$")
  Col_Letter = vArr(0)
End Function
Private Function Col_LetterHeader(sheetName As String, headText As String, Optional headRow = 1) As String
  'Get column letter from header text
  Dim lngColNumber As Long
  
  lngColNumber = Col_NumberHeader(sheetName, headText, headRow)
  Col_LetterHeader = Col_Letter(lngColNumber)
End Function
Private Function Col_Number(colLetter) As Long
  'Get column number from column letter
  Col_Number = Range(colLetter & "1").Column
End Function
Private Function Col_NumberHeader(sheetName As String, headText As String, Optional headRow = 1) As Long
  'Get column number from header text
  Dim i As Long
  Dim strCellString As String
  
  Col_NumberHeader = 0
  For i = 1 To 400
    strCellString = Trim(CStr(Sheets(sheetName).Cells(headRow, i)))
    If strCellString = headText Then
      Col_NumberHeader = i
      Exit Function
    End If
  Next i
End Function
Private Sub ColLetterTests()
  'Test for Col_Letter, Col_LetterHeader, Col_Number and Col_NumberHeader
  Debug.Print Col_Letter(12)
  Debug.Print Col_LetterHeader("Hogyallunk", "Any.csop.")
  Debug.Print Col_Number("H")
  Debug.Print Col_NumberHeader("Hogyallunk", "Any.csop.")
End Sub
Private Function GetLastRow(sheetName As String, checkColumn As Long, _
  Optional firstrow = 2, Optional lastrow = 600000, _
  Optional backwardCheck = True) As Long
  'Adott fül utolsó sora
  Dim i As Long
  Dim curSheet As Worksheet
  Dim strCell As String
  
  Set curSheet = ActiveWorkbook.ActiveSheet
  Sheets(sheetName).Activate
  GetLastRow = 0
  If backwardCheck Then
    For i = lastrow To firstrow Step -1
      strCell = Trim(CStr(Cells(i, checkColumn)))
      If strCell <> "" Then
        GetLastRow = i
        Exit For
      End If
    Next i
  Else
    For i = firstrow To lastrow
      strCell = Trim(CStr(Cells(i, checkColumn)))
      If strCell = "" Then
        GetLastRow = i - 1
        Exit For
      End If
    Next i
  End If
  curSheet.Activate
  Set curSheet = Nothing
  Debug.Print "LastRow of " & sheetName & ": " & GetLastRow & " ChkCol:" & checkColumn
End Function


'--------------
'Refresh ON OFF
'--------------
Private Sub RefreshOFF()
  'Screen update OFF
  With Application
    .ScreenUpdating = False
    .EnableEvents = False
    '.Calculation = xlCalculationManual
  End With
End Sub
Private Sub RefreshON()
  'Screen update ON
  With Application
    .ScreenUpdating = True
    .EnableEvents = True
    '.Calculation = xlCalculationAutomatic
  End With
End Sub

'----------------------
'Change keyboard layout
'----------------------
Sub SwitchToENG()
  'Váltás angolra
  Call ActivateKeyboardLayout(1033, 0)
End Sub
Sub SwitchToHUN()
  'Váltás magyarra
  Call ActivateKeyboardLayout(1038, 0)
End Sub






