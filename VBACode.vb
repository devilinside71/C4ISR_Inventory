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