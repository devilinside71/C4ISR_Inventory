VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_strPathToLogfile As String
Private Const cmstrPathToLogfile As String = "program.log"
Private m_strDelimiter As String
Private Const cmstrDelimiter As String = ", "
Private Const cmbDebugMode As Boolean = False

Public Property Let PathToLogfile(ByRef path2logfile As String)
  
  On Error GoTo PROP_ERR
  
  m_strPathToLogfile = path2logfile
  If cmbDebugMode Then
    logDEBUG "ProjectClass.PathToLogfile has been set to: " & m_strPathToLogfile, "LoggerClass.PathToLogfile"
  End If
PROP_EXIT:
  Exit Property
  
PROP_ERR:
  Err.Raise Err.Number
  Resume PROP_EXIT
End Property
Public Property Get PathToLogfile() As String
  
  On Error GoTo PROP_ERR
  
  PathToLogfile = m_strPathToLogfile
  
PROP_EXIT:
  Exit Property
  
PROP_ERR:
  Err.Raise Err.Number
  Resume PROP_EXIT
End Property
Public Property Let Delimiter(ByRef delim As String)
  
  On Error GoTo PROP_ERR
  
  m_strDelimiter = delim
  If cmbDebugMode Then
    logDEBUG "ProjectClass.Delimiter has been set to: " & m_strDelimiter, "LoggerClass.Delimiter"
  End If
  
PROP_EXIT:
  Exit Property
  
PROP_ERR:
  Err.Raise Err.Number
  Resume PROP_EXIT
End Property
Public Property Get Delimiter() As String
  
  On Error GoTo PROP_ERR
  
  Delimiter = m_strDelimiter
  
PROP_EXIT:
  Exit Property
  
PROP_ERR:
  Err.Raise Err.Number
  Resume PROP_EXIT
End Property
Private Sub Class_Initialize()
  
  m_strPathToLogfile = ActiveWorkbook.Path & Application.PathSeparator & cmstrPathToLogfile
  m_strDelimiter = Chr(9)
  If cmbDebugMode Then
    Debug.Print "Class LoggerClass initialized"
    logDEBUG "LoggerClass Default value for PathToLogfile: " & m_strPathToLogfile, "LoggerClass.Class_Initialize"
    logDEBUG "LoggerClass Default value for Delimiter: " & m_strDelimiter, "LoggerClass.Class_Initialize"
  End If
End Sub
Private Sub Class_Terminate()
  If cmbDebugMode Then
    Debug.Print "Class LoggerClass terminated"
    logDEBUG "Class LoggerClass terminated", "LoggerClass.Class_Terminate"
  End If
End Sub
Sub Reset()
  
  m_strPathToLogfile = cmstrPathToLogfile
  m_strDelimiter = cmstrDelimiter
  If cmbDebugMode Then
    logDEBUG "ProjectClass Default value for PathToLogfile: " & m_strPathToLogfile, "LoggerClass.Reset"
    logDEBUG "ProjectClass Default value for Delimiter: " & m_strDelimiter, "LoggerClass.Reset"
  End If
End Sub
Sub logEMERG(ByRef message As String, Optional ByRef callername = "unknown-caller")
  'A panic condition.'Messages that contain information normally of use only when debugging a program.
  WriteLog "emerg", message, callername
End Sub
Sub logALERT(ByRef message As String, Optional ByRef callername = "unknown-caller")
  'Action must be taken immediately'Messages that contain information normally of use only when debugging a program.
  'A condition that should be corrected immediately, such as a corrupted system database.
  WriteLog "alert", message, callername
End Sub
Sub logCRITICAL(ByRef message As String, Optional ByRef callername = "unknown-caller")
  'Critical conditions'Messages that contain information normally of use only when debugging a program.
  'Hard device errors.
  WriteLog "crit", message, callername
End Sub
Sub logERROR(ByRef message As String, Optional ByRef callername = "unknown-caller")
  'Error conditions'Messages that contain information normally of use only when debugging a program.
  WriteLog "err", message, callername
End Sub
Sub logWARNING(ByRef message As String, Optional ByRef callername = "unknown-caller")
  'Warning conditions'Messages that contain information normally of use only when debugging a program.
  WriteLog "warn", message, callername
End Sub
Sub logNOTICE(ByRef message As String, Optional ByRef callername = "unknown-caller")
  'Normal but significant conditions'Messages that contain information normally of use only when debugging a program.
  'Conditions that are not error conditions, but that may require special handling.
  WriteLog "notice", message, callername
End Sub
Sub logINFO(ByRef message As String, Optional ByRef callername = "unknown-caller")
  'Informational messages'Messages that contain information normally of use only when debugging a program.
  WriteLog "info", message, callername
End Sub
Sub logDEBUG(ByRef message As String, Optional ByRef callername = "unknown-caller")
  'Debug-level messages'Messages that contain information normally of use only when debugging a program.
  'Messages that contain information normally of use only when debugging a program.
  WriteLog "debug", message, callername
End Sub
Private Sub WriteLog(ByRef errlevel As String, ByRef message As String, ByRef callername)
  Dim sTimeStamp As String
  Dim sHost As String
  Dim sFacility As String
  Dim sErrorLevel As String
  Dim sMessage As String
  Dim sLogLine As String
  Dim FileNumber
  
  sTimeStamp = Format(Now(), "yyyy-MM-dd hh:mm:ss")
  #If Mac Then
    sHost = GetUserNameMac()
  #Else
    sHost = Environ$("computername")
  #End If
  
  sFacility = callername
  
  
  sErrorLevel = errlevel
  
  sMessage = message
  sLogLine = sTimeStamp & m_strDelimiter & sHost & m_strDelimiter & _
    sFacility & m_strDelimiter & sErrorLevel & m_strDelimiter & sMessage
  #If Mac Then
    Debug.Print "LOG: " & sLogLine
  #Else
    FileNumber = FreeFile()
    Open m_strPathToLogfile For Append As #FileNumber
    Print #FileNumber, sLogLine
    Close #FileNumber
  #End If
End Sub

Private Function GetUserNameMac() As String
  Dim sMyScript As String
  
  sMyScript = "set userName to short user name of (system info)" & vbNewLine & "return userName"
  GetUserNameMac = MacScript(sMyScript)
End Function

Private Function GetComputerNameMac() As String
  'MAC code here
End Function






