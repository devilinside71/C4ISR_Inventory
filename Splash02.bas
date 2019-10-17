Attribute VB_Name = "Splash"
Option Explicit
  
Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000

#If Mac Then
    'Mac code here
#ElseIf Win64 Then
  Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
  Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As LongPtr
  Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#ElseIf Win32 Then
  Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
  Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
      
Sub HideTitleBar(frm As Object)
  Dim lngWindow As Long
  Dim lFrmHdl As Long
  lFrmHdl = FindWindowA(vbNullString, frm.Caption)
  lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
  lngWindow = lngWindow And (Not WS_CAPTION)
  Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
  Call DrawMenuBar(lFrmHdl)
End Sub
  
  

