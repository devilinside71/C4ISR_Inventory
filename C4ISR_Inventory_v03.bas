Attribute VB_Name = "C4ISR_Inventory"
Option Explicit
'Password to unlock page: pwd123
#If Win64 Then
    Private Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As LongPtr, ByVal Flag As Long) As LongPtr
#Else
    Private Declare Function ActivateKeyboardLayout Lib "user32.dll" (ByVal HKL As Long, Flag As Long) As Long
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

Sub SetPlattform()
  Call clLogger.logINFO("Program start", "C4ISR_Inventory.SetPlattform")
  #If Mac Then
    isMacOS = True
    isWindows = False
    Call clLogger.logINFO("OS: MacOS", "C4ISR_Inventory.SetPlattform")
  #Else
    isMacOS = False
    isWindows = True
    Call clLogger.logINFO("OS: Windows", "C4ISR_Inventory.SetPlattform")
  #End If
  #If Win64 Then
    Is64BitOffice = True
    Call clLogger.logINFO("Office: 64bit", "C4ISR_Inventory.SetPlattform")
  #Else
    Is64BitOffice = False
    #If Mac Then
      Call clLogger.logINFO("Office: MacOS", "C4ISR_Inventory.SetPlattform")
    #Else
      Call clLogger.logINFO("Office: 32bit", "C4ISR_Inventory.SetPlattform")
    #End If
  #End If
End Sub
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
        Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.C4ISR_Inventory")
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
  lLastRowFullInventory = GetLastRow(cmshFullInventory, 1, 2, 50000, False)
  If cmbDebugMode Then
    Call clLogger.logDEBUG("Last row of " & cmshFullInventory & " sheet: " & lLastRowFullInventory, "C4ISR_Inventory.Add2FullInventory")
  End If
  
  
  shInventorySheet = vbNullString
  
  Set wbkADSInventory = Workbooks.Open(Filename:=Path2ADSInventoryFile)
  For Each ws In wbkADSInventory.Worksheets
    If cmbDebugMode Then
      Call clLogger.logDEBUG(ws.Name, "Add2FullInventory")
    End If
    pos = InStr(ws.Name, "Inventory")
    If pos <> 0 Then
      shInventorySheet = ws.Name
    End If
  Next
  Call clLogger.logINFO("Inventory sheet name: " & shInventorySheet, "C4ISR_Inventory.Add2FullInventory")
  
  sPackageName = ScanForPackageName(wbkADSInventory)
  If cmbDebugMode Then
    Call clLogger.logDEBUG("Package name: " & sPackageName, "C4ISR_Inventory.Add2FullInventory")
  End If
  ' Check if package is in the Full Inventory already
  bNotInTheList = True
  For i = 2 To lLastRowFullInventory
    sCell = Trim(CStr(wbkCurr.Sheets(cmshFullInventory).Cells(i, 1)))
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
    If cmbDebugMode Then
      Call clLogger.logDEBUG("Interrupted because " & sPackageName & " is in the list.", "C4ISR_Inventory.Add2FullInventory")
    End If
    GoTo PROC_EXIT
  End If
  
  'Copy data
  For i = 2 To 3000
    lLastRowFullInventory = lLastRowFullInventory + 1
    sCell = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, 1)))
    If sCell = "" Then
      Exit For
    End If
    wbkCurr.Sheets(cmshFullInventory).Cells(lLastRowFullInventory, 1) = sPackageName
    For k = 1 To 8
      wbkCurr.Sheets(cmshFullInventory).Cells(lLastRowFullInventory, k + 1) = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, k)))
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
    Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.Add2FullInventory")
  End If
  Resume PROC_EXIT
End Sub

Private Sub Add2Inventory(Path2ADSInventoryFile As String)
  'Add data to Inventory
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
  Dim lLastRowInventory As Long
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
  lLastRowInventory = 1
  
  
  
  shInventorySheet = vbNullString
  
  Set wbkADSInventory = Workbooks.Open(Filename:=Path2ADSInventoryFile)
  For Each ws In wbkADSInventory.Worksheets
    If cmbDebugMode Then
      Call clLogger.logDEBUG(ws.Name, "Add2Inventory")
    End If
    pos = InStr(ws.Name, "Inventory")
    If pos <> 0 Then
      shInventorySheet = ws.Name
    End If
  Next
  If cmbDebugMode Then
    Call clLogger.logDEBUG("Inventory sheet name: " & shInventorySheet, "C4ISR_Inventory.Add2Inventory")
  End If
  
 
  
  'Copy data
  For i = cmlInventoryStartRow To cmlInventoryEndRow
    lLastRowInventory = lLastRowInventory + 1
    sCell = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, 1)))
    If sCell = "" Then
      Exit For
    End If
    For k = 1 To 8
      wbkCurr.Sheets(cmshInventory).Cells(lLastRowInventory, k) = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, k)))
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
    Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.Add2Inventory")
  End If
  Resume PROC_EXIT
End Sub
Private Function DeleteInventory() As Boolean
  'Delete the content of Inventory sheet
  'Created by: Laszlo Tamas
  Dim lLastRowInventory As Long
  Dim sRowString As String
  Dim mgConfirm As Integer
  Dim bDelete As Boolean
  
  On Error GoTo FUNC_ERROR
  bDelete = False
  mgConfirm = MsgBox("Are you sure to delete Inventory items?", 36, "Confirmation")
  Select Case mgConfirm
    Case vbYes
      bDelete = True
    Case vbNo
      bDelete = False
  End Select
  If bDelete Then
    lLastRowInventory = GetLastRow(cmshInventory, 1, 2, 50000, False)
    sRowString = Trim(CStr(cmlInventoryStartRow)) & ":" & Trim(CStr(lLastRowInventory))
    Sheets(cmshInventory).Select
    Rows(sRowString).Select
    Selection.ClearContents
    Range("A2").Select
  End If
  
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  DeleteInventory = bDelete
  Exit Function
FUNC_ERROR:
  Debug.Print "Error in Function DeleteInventory"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.DeleteInventory")
  End If
  Resume FUNC_EXIT
End Function

Private Sub DeleteInventoryTest()
    'Test procedure for DeleteInventory
    Dim dtmStartTime As Date
    Dim res As Boolean


    dtmStartTime = Now()
    res = DeleteInventory()
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
    'Created by: Laszlo Tamas
    Dim dtmStartTime As Date
    Dim sPath2ADSInventoryFile As String

    sPath2ADSInventoryFile = GetPath2Inventory()


    dtmStartTime = Now()
    Call Add2FullInventory(sPath2ADSInventoryFile)
    MsgBox "Finished!"
End Sub
Sub AddToInventory()
  'Add2Inventory
  'Created by: Laszlo Tamas
  Dim dtmStartTime As Date
  Dim sPath2ADSInventoryFile As String
  Dim res As Boolean
  
  
  res = DeleteInventory()
  
  If res Then
    sPath2ADSInventoryFile = GetPath2Inventory()
    Call Add2Inventory(sPath2ADSInventoryFile)
  End If
  dtmStartTime = Now()
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
        Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.GetPath2Inventory")
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
    sCell = Trim(CStr(InvetoryBook.Sheets(cmshCoverPage).Cells(i, 1)))
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
    Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.ScanForPackageName")
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
Private Sub DeleteScannedData()
  'Delete previously scanned data
  'Parameters:
  'Created by: Laszlo Tamas
  Dim bDelete As Boolean


  On Error GoTo PROC_ERR

  bDelete = DeleteScan()
  If Not bDelete Then
    If cmbDebugMode Then
      Call clLogger.logDEBUG("Delete scanned data has been interrupted because scanned data was not deleted.", "C4ISR_Inventory.DeleteScannedData")
    End If
    GoTo PROC_EXIT
  End If

  '---------------
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure DeleteScannedData"
  If Err.Number Then
      Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.DeleteScannedData")
  End If
  Resume PROC_EXIT
End Sub
Private Sub DeleteScannedDataTest()
    'Test procedure for DeleteScannedData
    Dim dtmStartTime As Date



    dtmStartTime = Now()
    Call DeleteScannedData
End Sub

Function DeleteScan() As Boolean
  'Korábban beolvasott adatok törlése
  'Created by: Laszlo Tamas
  Dim i As Long
  Dim sRange As String
  Dim bDelete As Boolean
  Dim mbResult As Integer

  On Error GoTo FUNC_ERR
  
  'Code here
  Sheets(cmshScan).Activate
  'Confirmation
  bDelete = False
  mbResult = MsgBox("Are you sure to delete scanned data?", vbYesNo + vbQuestion)
  Select Case mbResult
    Case vbYes
      Sheets(cmshScan).Select
      sRange = "A2:A" & Trim(CStr(cmlScanListEndRow))
      Range(sRange).Select
      Selection.ClearContents
      Range("A2").Select
      bDelete = True
    Case vbNo
      bDelete = False
  End Select
  
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  DeleteScan = bDelete
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function DeleteScan"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.DeleteScan")
  End If
  Resume FUNC_EXIT
End Function
Sub ReadFromFile()
  'Read from Motorola CS3070 Barcode.txt
  'Parameters:
  'Created by: Laszlo Tamas
  
  Dim FileNum As Integer
  Dim DataLine As String
  Dim sPath As String
  Dim iChoice As Integer
  Dim sLineArr() As String
  Dim i As Long
  Dim bDelete As Boolean
  
  On Error GoTo PROC_ERR
  
  bDelete = DeleteScan()
  If Not bDelete Then
    If cmbDebugMode Then
      Call clLogger.logDEBUG("Read from file has been interrupted because scanned data was not deleted.", "C4ISR_Inventory.ReadFromFile")
    End If
    GoTo PROC_EXIT
  End If
  i = 1
  'only allow the user to select one file
  Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
  'make the file dialog visible to the user
  iChoice = Application.FileDialog(msoFileDialogOpen).Show
  'determine what choice the user made
  If iChoice <> 0 Then
    'get the file path selected by the user
    sPath = Application.FileDialog( _
    msoFileDialogOpen).SelectedItems(1)
    'print the file path to sheet 1
    'Cells(2, 1) = sPath
    FileNum = FreeFile()
    Open sPath For Input As #FileNum
    
    While Not EOF(FileNum)
      Line Input #FileNum, DataLine ' read in data 1 line at a time
      ' decide what to do with dataline,
      ' depending on what processing you need to do for each case
      sLineArr = Split(DataLine, ",")
      'Debug.Print sLineArr(3)
      i = i + 1
      Sheets(cmshScan).Cells(i, 1) = sLineArr(3)
    Wend
    Close #FileNum
  End If
  
  
  
  
  '---------------
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure ReadFromFile"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "C4ISR_Inventory.ReadFromFile")
  End If
  Resume PROC_EXIT
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
  'Last row of a sheet
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
  'Switch to ENG keyboard
  #If Mac Then
  #Else
    Call ActivateKeyboardLayout(1033, 0)
  #End If
End Sub
Sub SwitchToHUN()
  'Switch to HUN keyboard
  #If Mac Then
  #Else
    Call ActivateKeyboardLayout(1038, 0)
  #End If
End Sub






