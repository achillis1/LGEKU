VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPassword 
   Caption         =   "Password"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   OleObjectBlob   =   "frmPassword.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private iTrytimes As Integer

Private Sub cmdCancel_Click()
    Me.Hide
    frmMain.Show vbModeless
End Sub

Private Sub cmdOK_Click()
    iTrytimes = iTrytimes + 1
    If iTrytimes > 2 Then
        MsgBox "You have tried more than 3 times. The workbook will be closed!"
        ThisWorkbook.Close SaveChanges:=1
    End If
    
    If txtPassword.Text = "Abcd123" Then
        'MsgBox "password is correct"
        Me.Hide
        Application.Visible = True
    Else
        MsgBox "password is not correct"
        txtPassword.Text = ""
        txtPassword.SetFocus
    End If
End Sub

Private Sub UserForm_Activate()
    iTrytimes = 0
    txtPassword.Text = ""
End Sub


