VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashUserForm 
   Caption         =   "UserForm2"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9045
   OleObjectBlob   =   "SplashUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SplashUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Activate()
    SplashUserForm.Label2.Caption = "Ministry of Defence Electronics, Logistics and Property Management Co."
    SplashUserForm.Label3.Caption = "Copyright 2025"
    Application.Wait (Now + TimeValue("00:00:01"))
    SplashUserForm.Label4.Caption = "Loading Data..."
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:01"))
    SplashUserForm.Label4.Caption = "Creating Forms..."
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:01"))
    SplashUserForm.Label4.Caption = "Opening..."
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:01"))
    Unload SplashUserForm
End Sub
Private Sub UserForm_Initialize()
    HideTitleBar Me
End Sub
