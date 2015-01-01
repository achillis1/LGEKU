VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServiceCenter 
   Caption         =   "Service Center"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   OleObjectBlob   =   "frmServiceCenter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmServiceCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnrollment_Click()
    Me.Hide
    frmEnrollment.Show
End Sub

Private Sub cmdExit_Click()
    ThisWorkbook.Close
End Sub

Private Sub cmdExport_Click()
    Me.Hide
    frmExport.Show

End Sub

Private Sub cmdImport_Click()
    Me.Hide
    frmImport.Show
End Sub

Private Sub cmdSetting_Click()
    Me.Hide
    frmSetting.Show
End Sub

Private Sub cmdSystem_Click()
    Me.Hide
    frmSystem.Show
End Sub


Private Sub UserForm_Initialize()
    Application.Visible = False
End Sub

Private Sub UserForm_Terminate()
    Application.Visible = True
    'ThisWorkbook.Close
End Sub
