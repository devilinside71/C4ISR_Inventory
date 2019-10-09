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
'Class for Motorola CS3070 Bluetooth Barcode Reader
'Created by: Laszlo Tamas
'Licence: MIT
Private clLogger As New LoggerClass
Private m_DefaultPath As String
Private Const cmDefaultPath As String = vbNullString

Public Property Let DefaultPath(DefaultPathParam As String)
  On Error GoTo PROP_ERR
  m_DefaultPath = DefaultPathParam
  clLogger.logDEBUG "MotorolaCS3070Class.Property.DefaultPath has been set to: " & m_DefaultPath, "MotorolaCS3070Class.Property.Let.DefaultPath"
PROP_EXIT:
  Exit Property
PROP_ERR:
  If Err.Number Then
    clLogger.logERROR Err.Description, "MotorolaCS3070Class.Property.Let.DefaultPath"
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
    clLogger.logERROR Err.Description, "MotorolaCS3070Class.Property.Get.DefaultPath"
  End If
  Resume PROP_EXIT
End Property

Private Sub Class_Initialize()
  m_DefaultPath = cmDefaultPath
  clLogger.logDEBUG "Class MotorolaCS3070Class initialized", "MotorolaCS3070Class.Class_Initialize"
  clLogger.logDEBUG "Class MotorolaCS3070Class default value for DefaultPath: " & m_DefaultPath, "MotorolaCS3070Class.Class_Initialize"
End Sub

Private Sub Class_Terminate()
  clLogger.logDEBUG "Class MotorolaCS3070Class terminated", "MotorolaCS3070Class.Class_Terminate"
End Sub
