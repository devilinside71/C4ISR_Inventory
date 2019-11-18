Attribute VB_Name = "Module1"
Option Explicit
Private clLogger As New LoggerClass
Private clMotorola As New MotorolaCS3070Class

Sub test_1_MotorolaCS3070Class()
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
Sub test_MAC_1_MotorolaCS3070Class()
  'Test for Class MotorolaCS3070Class
  clLogger.logDEBUG clMotorola.SelectFileMac("{""Public.plain-text""}"), "test_MAC_1_MotorolaCS3070Class"
End Sub
Sub testM3Class()
 'Test For Class M3Class
 Dim clClass As New M3Class
 Dim textdata() As String
 Dim i As Long
 textdata = clClass.GetTextData()
 If textdata(1) <> "" Then
 For i = 1 To UBound(textdata)
 clLogger.logDEBUG "ReadTextData test: >> " & Trim(CStr(i)) & ": " & textdata(i), "testM3Class"
 clLogger.logDEBUG "GetBarcodeDataFromBarcodeLine test: >> " & Trim(CStr(i)) & ": " & clClass.GetBarcodeDataFromBarcodeLine(textdata(i)), "testM3Class"
 Next i
 Else
 clLogger.logDEBUG "ReadTextData test: >> EMPTY", "testM3Class"
 End If
End Sub
Sub test_MAC_1_M3Class()
  'Test for Class M3Class
  clLogger.logDEBUG clMotorola.SelectFileMac("{""Public.plain-text""}"), "test_MAC_1_M3Class"
End Sub

Sub others()
  'Test for Class MotorolaCS3070Class Mac
  
  
  'Description
  'Created by: Laszlo Tamas
  'Licence: MIT



'---------------------------------
#If Mac Then
  'Mac code here
#ElseIf Win64 Then
#ElseIf Win32 Then
#End If


'---------------------------------
#If Mac Then
  'Mac code here
#Else
#End If

'---------------------------------
#If Mac Then
  'Mac code here
#Else
#End If

'---------------------------------
#If Mac Then
  'Mac code here
#Else
  #If VBA7 Then
  #Else
  #End If
#End If


'---------------------------------
#If Mac Then
  'Mac code here
#Else
  #If VBA7 Then
    #If Win64 Then
    #Else
    #End If
  #Else
  #End If
#End If


End Sub
