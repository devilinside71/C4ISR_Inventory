VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MotorolaCS3070Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    'Class For Motorola CS3070 Bluetooth Barcode Reader
    'Created by: Laszlo Tamas
    'Licence: MIT
    #If Mac Then
      'Mac code here
    #ElseIf Win64 Then
      Private Declare PtrSafe Function ActivateKeyboardLayout Lib "USER32" (ByVal HKL As LongPtr, _
            ByVal Flag As Long) As LongPtr
    #ElseIf Win32 Then
      Private Declare Function ActivateKeyboardLayout Lib "user32.dll" (ByVal HKL As Long, _
            ByRef Flag As Long) As Long
    #End If
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
  clLogger.logDEBUG "Class MotorolaCS3070Class initialized", _
        "MotorolaCS3070Class.Class_Initialize"
End Sub
  
Private Sub Class_Terminate()
  clLogger.logDEBUG "Class MotorolaCS3070Class terminated", _
        "MotorolaCS3070Class.Class_Terminate"
End Sub
  
Function GetTextData() As Variant
  'Read text data stored On Scanner in USB mode
  'Returns:{Variant()} Data lines
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim sPath() As String
  Dim sLines() As String
  Dim DataLine As String
  Dim lLineCount As Long
  Dim FileNum As Long
  
  
  On Error GoTo FUNC_ERR
  
  
  #If Mac Then
    'Mac code here
    clLogger.logDEBUG "Select file Mac", "MotorolaCS3070Class.SelectFile"
    ReDim Preserve sPath(1)
    sPath = SelectFileMac(False, "{""public.plain-text""}")
  #Else
    sPath = SelectFile(False, "Select Barcode TXT File", _
          "Scanner Barcode TXT files", "*.txt")
  #End If
        
  #If Mac Then
    'Mac code here
    ReDim Preserve sLines(1)
    sLines(1) = vbNullString
    m_PathTextData = sPath(1)
  #Else
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
    Else
      ReDim Preserve sLines(1)
      sLines(1) = vbNullString
      m_PathTextData = vbNullString
    End If
  #End If
  GetTextData = sLines
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function MotorolaCS3070Class.GetTextData"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "MotorolaCS3070Class.GetTextData")
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
  
    clLogger.logDEBUG "Select file WIN mode", "MotorolaCS3070Class.SelectFile"
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
      clLogger.logDEBUG "Nothing selected", "MotorolaCS3070Class.SelectFile"
    End If
  SelectFile = sPath
  Set dialogBox = Nothing
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  If Err.Number Then
    Debug.Print "Error in Function MotorolaCS3070Class.SelectFile >> " & _
          Err.Description
    clLogger.logERROR Err.Description, "MotorolaCS3070Class.SelectFile"
  End If
  Resume FUNC_EXIT
End Function
Function SelectFileMac(ByVal Multiselect As Boolean, ByVal FileFormats As String) As Variant
  'Open file dialog to select file(s)
  'TODO: NOT IMPLEMENTED YET
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim MySplit As Variant
    Dim N As Long
    Dim Fname As String
    Dim mybook As Workbook
    Dim OneFile As Boolean
    Dim FileFormat As String
  
  Dim sPath() As String
  
  On Error GoTo FUNC_ERR
'    FileFormat = "{""org.openxmlformats.spreadsheetml.sheet"",""org.openxmlformats.spreadsheetml.sheet.macroenabled""}"
    FileFormat = FileFormats

    ' Set to True if you only want to be able to select one file
    ' And to False to be able to select one or more files
    If Multiselect Then
    OneFile = False
    Else
    OneFile = True
    End If
    
    On Error Resume Next
    MyPath = MacScript("return (path to desktop folder) as String")
  
    'Building the applescript string, do not change this
    If Val(Application.Version) < 15 Then
        'This is Mac Excel 2011
        If OneFile = True Then
            MyScript = _
                "set theFile to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
                "return theFile"
        Else
            MyScript = _
                "set applescript's text item delimiters to {ASCII character 10} " & vbNewLine & _
                "set theFiles to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file or files"" default location alias """ & _
                MyPath & """ with multiple selections allowed) as string" & vbNewLine & _
                "set applescript's text item delimiters to """" " & vbNewLine & _
                "return theFiles"
        End If
    Else
        'This is Mac Excel 2016
        If OneFile = True Then
            MyScript = _
                "set theFile to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
                "return posix path of theFile"
        Else
            MyScript = _
                "set theFiles to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file or files"" default location alias """ & _
                MyPath & """ with multiple selections allowed)" & vbNewLine & _
                "set thePOSIXFiles to {}" & vbNewLine & _
                "repeat with aFile in theFiles" & vbNewLine & _
                "set end of thePOSIXFiles to POSIX path of aFile" & vbNewLine & _
                "end repeat" & vbNewLine & _
                "set {TID, text item delimiters} to {text item delimiters, ASCII character 10}" & vbNewLine & _
                "set thePOSIXFiles to thePOSIXFiles as text" & vbNewLine & _
                "set text item delimiters to TID" & vbNewLine & _
                "return thePOSIXFiles"
        End If
    End If

    MyFiles = MacScript(MyScript)
    On Error GoTo 0

    'If you select one or more files MyFiles is not empty
    'We can do things with the file paths now like I show you below
    If MyFiles <> "" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With

        MySplit = Split(MyFiles, Chr(10))
        For N = LBound(MySplit) To UBound(MySplit)

            'Get file name only and test if it is open
            Fname = Right(MySplit(N), Len(MySplit(N)) - InStrRev(MySplit(N), _
                Application.PathSeparator, , 1))

            If bIsBookOpen(Fname) = False Then
              ReDim Preserve sPath(1)
              sPath(1) = MySplit(N)
              clLogger.logDEBUG "Selected: " & sPath(1), "MotorolaCS3070Class.SelectFileMac"
            Else
                clLogger.logDEBUG "We skip this file : " & MySplit(N) & " because it Is already open", "MotorolaCS3070Class.SelectFileMac"
            End If

            Next N
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With
    End If

  
  
'  ReDim Preserve sPath(1)
'  sPath(1) = "/Users/macos/Documents/GitHub/C4ISR_Inventory/BARCODE.txt"
  SelectFileMac = sPath
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  If Err.Number Then
    Debug.Print "Error in Function MotorolaCS3070Class.SelectFileMac >> " & _
          Err.Description
    clLogger.logERROR Err.Description, "MotorolaCS3070Class.SelectFileMac"
  End If
  Resume FUNC_EXIT
End Function
  
  
Function GetBarcodeDataFromBarcodeLine(ByRef BarcodeLine As String) As String
  'Extract barcode data from scnnar barcode TXT line
  'Parameters:
  ' {String} BarcodeLine: A line from Scanner's TXT file
  ' Example: 07/02/13,17:58:31,08,000006819309
  'Returns:{String} Only the barcode: 000006819309
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim sLineArr() As String
  
  On Error GoTo FUNC_ERR
  
  GetBarcodeDataFromBarcodeLine = vbNullString
  sLineArr = Split(BarcodeLine, ",")
  
  GetBarcodeDataFromBarcodeLine = sLineArr(3)
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function MotorolaCS3070Class.GetBarcodeDataFromBarcodeLine"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, _
          "MotorolaCS3070Class.GetBarcodeDataFromBarcodeLine")
  End If
  Resume FUNC_EXIT
End Function
  
Sub SwitchToENG()
  'Switch to ENG keyboard
  On Error GoTo PROC_ERR
  #If Mac Then
    clLogger.logNOTICE "Switch keyboard not implemented yet", _
          "MotorolaCS3070Class.SwitchToENG"
  #Else
    ActivateKeyboardLayout 1033, 0
    clLogger.logDEBUG "Switched to English(US) keyboard", _
          "MotorolaCS3070Class.SwitchToENG"
  #End If
  
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure MotorolaCS3070Class.SwitchToENG"
  If Err.Number Then
    clLogger.logERROR Err.Description, "MotorolaCS3070Class.SwitchToENG"
  End If
  Resume PROC_EXIT
End Sub
  
Sub SwitchToHUN()
  'Switch to HUN keyboard
  On Error GoTo PROC_ERR
  #If Mac Then
    clLogger.logNOTICE "Switch keyboard not implemented yet", _
          "MotorolaCS3070Class.SwitchToHUN"
  #Else
    ActivateKeyboardLayout 1038, 0
    clLogger.logDEBUG "Switched to Hungarian keyboard", _
          "MotorolaCS3070Class.SwitchToHUN"
  #End If
  
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure MotorolaCS3070Class.SwitchToHUN"
  If Err.Number Then
    clLogger.logERROR Err.Description, "MotorolaCS3070Class.SwitchToHUN"
  End If
  Resume PROC_EXIT
End Sub
Function bIsBookOpen(ByRef szBookName As String) As Boolean
    ' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function
  
  
  
  
  
'Sub testMotorolaCS3070Class()
' 'Test For Class MotorolaCS3070Class
' Dim clClass As New MotorolaCS3070Class
' Dim textdata() As String
' Dim i As Long
' textdata = clClass.GetTextData()
' If textdata(1) <> "" Then
' For i = 1 To UBound(textdata)
' clLogger.logDEBUG "ReadTextData test: >> " & Trim(CStr(i)) & ": " & textdata(i), "testMotorolaCS3070Class"
' clLogger.logDEBUG "GetBarcodeDataFromBarcodeLine test: >> " & Trim(CStr(i)) & ": " & clClass.GetBarcodeDataFromBarcodeLine(textdata(i)), "testMotorolaCS3070Class"
' Next i
' Else
' clLogger.logDEBUG "ReadTextData test: >> EMPTY", "testMotorolaCS3070Class"
' End If
'End Sub
    
    
    
    
    
    
    
    
    
    
    


