VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private clLogger As New LoggerClass
Private m_DefaultPath As String
Private Const cmDefaultPath As String = vbNullString

Public Property Let DefaultPath(DefaultPathParam As String)
  On Error GoTo PROP_ERR
  m_DefaultPath = DefaultPathParam
  clLogger.logDEBUG "MyClass.Property.DefaultPath has been set to: " & m_DefaultPath, "MyClass.Property.Let.DefaultPath"
PROP_EXIT:
  Exit Property
PROP_ERR:
  If Err.Number Then
    clLogger.logERROR Err.Description, "Property.Let.DefaultPath"
  End If
  Resume PROP_EXIT
End Property

Public Property Get DefaultPath() As String
  On Error GoTo PROP_ERR
  DefaultPath = m_DefaultPath
PROP_EXIT:
  Exit Property
PROP_ERR:
  If Err.Number Then
    clLogger.logERROR Err.Description, "Property.Get.DefaultPath"
  End If
  Resume PROP_EXIT
End Property

Private Sub Class_Initialize()
  m_DefaultPath = cmDefaultPath
  clLogger.logDEBUG "Class MyClass initialized", "MyClass.Class_Initialize"
  clLogger.logDEBUG "Class MyClass default value for DefaultPath: " & m_DefaultPath, "MyClass.Class_Initialize"
End Sub

Private Sub Class_Terminate()
  clLogger.logDEBUG "Class MyClass terminated", "MyClass.Class_Terminate"
End Sub

'****************************************************************************
'Sub testMyClass()
'  'Test for Class MyClass
'  Dim clClass As New MyClass
'  clClass.DefaultPath = "dhskjad"
'End Sub

