VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSystemHEAP 
   Caption         =   "Systems HEAP"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10740
   OleObjectBlob   =   "frmSystemHEAP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSystemHEAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
    Me.Hide
    frmMain.Show vbModeless
End Sub

Private Sub UserForm_Initialize()
    auditlastrow = Worksheets(AuditSheetName).Range("E" & Rows.Count).End(xlUp).Row
    Call updatesystemlist
    
End Sub

Private Sub updatesystemlist()
    lstSystemHEAP.Clear
    For i = 2 To auditlastrow
        system = Worksheets(AuditSheetName).Cells(i, 1).Value
        lstSystemHEAP.AddItem (system)
    Next i
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub
