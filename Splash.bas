Attribute VB_Name = "Splash"
Option Explicit

'https://jkp-ads.com/Articles/apideclarations.asp

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
  
#If Mac Then
  'Mac code here
#Else
  #If VBA7 Then
    #If Win64 Then
      Private Declare PtrSafe Function SetWindowLong Lib "USER32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
      Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
  #Else
    Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
          ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  #End If
#End If

#If Mac Then
  'Mac code here
#Else
  #If VBA7 Then
    #If Win64 Then
      Private Declare PtrSafe Function GetWindowLong Lib "USER32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) As LongPtr
    #Else
      Private Declare PtrSafe Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) As LongPtr
    #End If
  #Else
    Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
          ByVal nIndex As Long) As Long
  #End If
#End If

#If Mac Then
  'Mac code here
#Else
  #If VBA7 Then
    Public Declare PtrSafe Function DrawMenuBar Lib "USER32" (ByVal hWnd As LongPtr) As Long
  #Else
    Public Declare Function DrawMenuBar Lib "USER32" (ByVal hWnd As Long) As Long
  #End If
#End If

#If Mac Then
  'Mac code here
#ElseIf Win64 Then
  Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
#ElseIf Win32 Then
  Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
#End If
  
Sub HideTitleBar(frm As Object)
  #If Mac Then
    'Mac code here
  #ElseIf Win64 Then
    Dim lFrmHdl As LongPtr
  #ElseIf Win32 Then
    Dim lFrmHdl As Long
  #End If
  
  #If Mac Then
    'Mac code here
  #Else
    #If VBA7 Then
      Dim lngWindow As LongPtr
    #Else
      Dim lngWindow As Long
    #End If
  #End If
  
  lFrmHdl = FindWindow(vbNullString, frm.Caption)
  lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
  lngWindow = lngWindow And (Not WS_CAPTION)
  Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
  Call DrawMenuBar(lFrmHdl)
End Sub
  
  
  
  
  
  

  
  
  
  
  
  
  
  
  
  


