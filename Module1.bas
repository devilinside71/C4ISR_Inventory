Attribute VB_Name = "Module1"
Option Explicit
Private clLogger As New LoggerClass
Private clMotorola As New MotorolaCS3070Class

Sub testMotorolaCS3070Class()
  'Test for Class MotorolaCS3070Class
  Dim textdata() As String
  Dim i As Long
  textdata = clMotorola.GetTextData()
  If textdata(1) <> "" Then
    For i = 1 To UBound(textdata)
      clLogger.logDEBUG "ReadTextData test: >> " & Trim(CStr(i)) & ": " & textdata(i), "testMotorolaCS3070Class"
      clLogger.logDEBUG "GetBarcodeDataFromBarcodeLine test: >> " & Trim(CStr(i)) & ": " & clMotorola.GetBarcodeDataFromBarcodeLine(textdata(i)), "testMotorolaCS3070Class"
    Next i
  Else
    clLogger.logDEBUG "ReadTextData test: >> EMPTY", "testMotorolaCS3070Class"
  End If
End Sub

Sub testMotorolaCS3070ClassMac()
  'Test for Class MotorolaCS3070Class Mac

End Sub
