VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "M3Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    'Class For M3 mobile compia Handheld PC
    'Created by: Laszlo Tamas
    'Licence: MIT
Private clLogger As New LoggerClass
Private m_PathTextData As String
Public Property Get PathTextData() As String
  On Error GoTo PROP_ERR
  PathTextData = m_PathTextData
PROP_EXIT:
  Exit Property
PROP_ERR:
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "Property.Get.PathTextData")
  End If
  Resume PROP_EXIT
End Property
Private Sub Class_Initialize()
  clLogger.logDEBUG "Class M3Class initialized", _
        "M3Class.Class_Initialize"
End Sub
  
Private Sub Class_Terminate()
  clLogger.logDEBUG "Class M3Class terminated", _
        "M3Class.Class_Terminate"
End Sub
  
Function GetTextData() As Variant
  'Read text data stored On Scanner in USB mode
  'Returns:{Variant()} Data lines
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim sPath() As String
  Dim sLines() As String
  Dim DataLine As String
  Dim sTempLine As String
  Dim lLineCount As Long
  Dim FileNum As Long
  
  
  On Error GoTo FUNC_ERR
  
  
  #If Mac Then
    sPath = SelectFileMac("{""Public.plain-text""}")
  #Else
    sPath = SelectFile(False, "Select Barcode File", _
       "Scanner Barcode files", "*.txt,*.pwd,*.rtf")
  #End If
  
  If sPath(1) <> "" Then
    FileNum = FreeFile()
    m_PathTextData = sPath(1)
    Open sPath(1) For Input As #FileNum
    Do While Not EOF(FileNum)
      Line Input #FileNum, DataLine
      lLineCount = lLineCount + 1
      ReDim Preserve sLines(lLineCount)
      sLines(lLineCount) = DataLine
    Loop
    Close #FileNum
    Call clLogger.logDEBUG("M3 lines read: " & CStr(lLineCount), "M3Class.GetTextData")
  Else
    ReDim Preserve sLines(1)
    sLines(1) = vbNullString
    m_PathTextData = vbNullString
  End If
  GetTextData = sLines
  ' - - - - - - - - - - - - - - -
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function M3Class.GetTextData"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "M3Class.GetTextData")
  End If
  Resume FUNC_EXIT
End Function


  
Function SelectFile(ByVal Multiselect As Boolean, _
        ByVal DialogTitle As String, _
              ParamArray FileFilter() As Variant) As Variant
  'Open file dialog to select file(s)
  'Parameters:
  ' {Boolean} Multiselect: Allow multiselect
  ' {String} DialogTitle: Title of the dialog box
  ' {Optional Variant()} FileFilter: Filter array, 1:Name, 2:Extensions
  ' Example1: SelectFile(False, "Select File")
  ' Example2: SelectFile(True, "Select Files")
  ' Example3: SelectFile(True, "Select Files", "Excel files", "*.xlsx,*.xls,*.xlsm")
  'Returns:
  ' {Variant()} Path(s) to selected file(s)
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim sPath() As String
  Dim iChoice As Long
  Dim dialogBox As FileDialog
  Dim i As Long
  
  On Error GoTo FUNC_ERR
  
    clLogger.logDEBUG "Select file WIN mode", "M3Class.SelectFile"
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    dialogBox.AllowMultiSelect = Multiselect
    dialogBox.Title = DialogTitle
    dialogBox.Filters.Clear
    If UBound(FileFilter) = 1 Then
      dialogBox.Filters.Add FileFilter(0), FileFilter(1)
    End If
    iChoice = dialogBox.Show
    If iChoice <> 0 Then
      For i = 1 To dialogBox.SelectedItems.Count
        ReDim Preserve sPath(i)
        sPath(i) = dialogBox.SelectedItems.Item(i)
      Next i
    Else
      ReDim Preserve sPath(1)
      sPath(1) = vbNullString
      clLogger.logDEBUG "Nothing selected", "M3Class.SelectFile"
    End If
  SelectFile = sPath
  Set dialogBox = Nothing
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  If Err.Number Then
    Debug.Print "Error in Function M3Class.SelectFile >> " & _
          Err.Description
    clLogger.logERROR Err.Description, "M3Class.SelectFile"
  End If
  Resume FUNC_EXIT
End Function
Function SelectFileMac(ByVal FileFormats As String) As Variant
  'Open file dialog to select file
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim MyPath As String
  Dim MyScript As String
  Dim MyFiles As String
  Dim MySplit As Variant
  Dim N As Long
  Dim Fname As String
  Dim mybook As Workbook
  Dim FileFormat As String
  
  Dim sPath() As String
  
  On Error GoTo FUNC_ERR
  clLogger.logDEBUG "Select file Mac mode", "M3Class.SelectFileMac"
  
  FileFormat = FileFormats
  
  
  
  On Error Resume Next
  MyPath = MacScript("Return (path to desktop folder) As String")
  
  'Building the applescript string, Do not change this
  If Val(Application.version) < 15 Then
    'This is Mac Excel 2011
    MyScript = _
          "Set theFile to (choose file of Type" & _
                " " & FileFormat & " " & _
                      "With prompt ""Please select a file"" default location alias """ & _
                            MyPath & """ without multiple selections allowed) As String" & vbNewLine & _
                                  "Return theFile"
  Else
    'This is Mac Excel 2016
    MyScript = _
          "Set theFile to (choose file of Type" & _
                " " & FileFormat & " " & _
                      "With prompt ""Please select a file"" default location alias """ & _
                            MyPath & """ without multiple selections allowed) As String" & vbNewLine & _
                                  "Return posix path of theFile"
  End If
  
  MyFiles = MacScript(MyScript)
  On Error GoTo 0
  
  'If you select one or more files MyFiles is not empty
  'We can Do things With the file paths now like I show you below
  If MyFiles <> "" Then
    With Application
      .ScreenUpdating = False
      .EnableEvents = False
    End With
    
    MySplit = Split(MyFiles, Chr(10))
    For N = LBound(MySplit) To UBound(MySplit)
      
      'Get file Name only and test If it is Open
      Fname = Right(MySplit(N), Len(MySplit(N)) - InStrRev(MySplit(N), _
            Application.PathSeparator, , 1))
      
      If IsBookOpen(Fname) = False Then
        ReDim Preserve sPath(1)
        sPath(1) = MySplit(N)
        clLogger.logDEBUG "Selected: " & sPath(1), "M3Class.SelectFileMac" _
           
      Else
        clLogger.logDEBUG "Skipping this file : " & MySplit(N) & _
           " because it is already Open", "M3Class.SelectFileMac"
      End If
      
    Next N
    With Application
      .ScreenUpdating = True
      .EnableEvents = True
    End With
  End If
  
  SelectFileMac = sPath
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  If Err.Number Then
    Debug.Print "Error in Function M3Class.SelectFileMac >> " & _
       Err.Description
    clLogger.logERROR Err.Description, "M3Class.SelectFileMac"
  End If
  Resume FUNC_EXIT
End Function

Function IsBookOpen(ByRef szBookName As String) As Boolean
  'Check if workbook is open
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  On Error Resume Next
  IsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function
  
Function GetBarcodeDataFromBarcodeLine(ByRef BarcodeLine As String) As String
  'Extract barcode data from scanner barcode TXT line
  'Parameters:
  ' {String} BarcodeLine: A line from Scanner's TXT file
  ' Example(TXT): CZC737843V
  ' Example(PWD): \viewkind4\uc1\pard\cf0\f0\fs20 CZC737843V\par
  ' Example(PWD): CZC737843V\par
  'Returns:{String} Only the barcode: CZC737843V
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim sTempLine As String
  Dim bLineOk As Boolean
  Dim bTxtLine As Boolean
  Dim bPwdLine As Boolean
  On Error GoTo FUNC_ERR
  
  bLineOk = False
  bPwdLine = False
  bTxtLine = True
  GetBarcodeDataFromBarcodeLine = vbNullString
  If EndsWith(BarcodeLine, "\par") Then
    bPwdLine = True
  ElseIf StartsWith(BarcodeLine, "*") Or Contains(BarcodeLine, "}") _
    Or Contains(BarcodeLine, "{") Then
      bTxtLine = False
  End If
  If bPwdLine Then
    sTempLine = Replace(BarcodeLine, "\viewkind4\uc1\pard\cf0\f0\fs20", "")
    sTempLine = Replace(sTempLine, "\par", "")
    sTempLine = Trim(sTempLine)
    GetBarcodeDataFromBarcodeLine = sTempLine
  ElseIf bTxtLine Then
    GetBarcodeDataFromBarcodeLine = BarcodeLine
  End If
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function M3Class.GetBarcodeDataFromBarcodeLine"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, _
          "M3Class.GetBarcodeDataFromBarcodeLine")
  End If
  Resume FUNC_EXIT
End Function
  
Private Function Contains(BaseString As String, SearchTerm As String) As Boolean
  'Determine if BaseString contains SearchTerm
  'Parameters:
  '  {String} BaseString
  '  {String} SearchTerm
  'Returns:{Boolean}
  'Created by: Laszlo Tamas
  'Licence: MIT

  On Error GoTo FUNC_ERR

  Contains = InStr(BaseString, SearchTerm)
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function M3Class.Contains"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "M3Class.Contains")
  End If
  Resume FUNC_EXIT
End Function

Private Function EndsWith(BaseString As String, SearchTerm As String) As Boolean
  'Determine if SearchTerm is at the end of BaseString
  'Parameters:
  '  {String} BaseString
  '  {String} SearchTerm
  'Returns:{Boolean}
  'Created by: Laszlo Tamas
  'Licence: MIT

  Dim endingLen As Integer
  On Error GoTo FUNC_ERR
  endingLen = Len(SearchTerm)
  EndsWith = (Right(Trim(UCase(BaseString)), endingLen) = UCase(SearchTerm))
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function M3Class.EndsWith"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "M3Class.EndsWith")
  End If
  Resume FUNC_EXIT
End Function

Private Function StartsWith(BaseString As String, SearchTerm As String) As Boolean
  'Determine if SearchTerm is at the beginning of BaseString
  'Parameters:
  '  {String} BaseString
  '  {String} SearchTerm
  'Returns:{Boolean}
  'Created by: Laszlo Tamas
  'Licence: MIT

  Dim startLen As Integer
  On Error GoTo FUNC_ERR
  startLen = Len(SearchTerm)
  StartsWith = (Left(Trim(UCase(BaseString)), startLen) = UCase(SearchTerm))
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function M3Class.StartsWith"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "M3Class.StartsWith")
  End If
  Resume FUNC_EXIT
End Function

  
  
  
  
'Sub testM3Class()
' 'Test For Class M3Class
' Dim clClass As New M3Class
' Dim textdata() As String
' Dim i As Long
' textdata = clClass.GetTextData()
' If textdata(1) <> "" Then
' For i = 1 To UBound(textdata)
' clLogger.logDEBUG "ReadTextData test: >> " & Trim(CStr(i)) & ": " & textdata(i), "testM3Class"
' clLogger.logDEBUG "GetBarcodeDataFromBarcodeLine test: >> " & Trim(CStr(i)) & ": " & clClass.GetBarcodeDataFromBarcodeLine(textdata(i)), "testM3Class"
' Next i
' Else
' clLogger.logDEBUG "ReadTextData test: >> EMPTY", "testM3Class"
' End If
'End Sub
'Sub test_MAC_1_M3Class()
'  'Test for Class M3Class
'  clLogger.logDEBUG clMotorola.SelectFileMac("{""Public.plain-text""}"), "test_MAC_1_M3Class"
'End Sub
    
    
    
    
    
    
    
    
    
    
    




