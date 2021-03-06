VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Main"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDirectAccess_Click()
    Me.Hide
    frmPassword.Show vbModeless
End Sub

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdHeapInfo_Click()
    Me.Hide
    frmHeapInfo.Show vbModeless
End Sub

Private Sub cmdMeasure_Click()
    Me.Hide
    frmMeasure.Show vbModeless
End Sub


Private Sub lstpopulate()
    lastrow = EnrollmentFirstDataLine
    For i = EnrollmentFirstDataLine To lastrow
        ROSAID = Worksheets(SheetName).Cells(i, NexantEnrollments.Enrollment_ID_ROSA).Value
        HEAPID = Worksheets(SheetName).Cells(i, NexantEnrollments.Enrollment_ID_HEAP).Value
        If ROSAID = "" And HEAPID <> "" Then
            currentEnrollment = HEAPID
        End If
        
        If ROSAID <> "" And HEAPID = "" Then
            currentEnrollment = ROSAID
        End If
    Next i

End Sub


Private Sub cmdRosaInfo_Click()
    Me.Hide
    frmRosaInfo.Show vbModeless
End Sub

Private Sub cmdSystem_Click()
    Me.Hide
    frmSystem.Show vbModeless

End Sub

Private Sub CommandButton1_Click()
    Me.Hide
    frmSystemHEAP.Show vbModeless
End Sub

Private Sub UserForm_Activate()
    ROSAHEAP = Worksheets(SettingSN).Cells(2, 1).Value
    If ROSAHEAP = 0 Then
        cmdHeapInfo.Enabled = False
        cmdRosaInfo.Enabled = True
    Else
        cmdRosaInfo.Enabled = False
        cmdHeapInfo.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
'    Application.Visible = False
    EnrollmentFirstDataLine = 11
    AuditSheetName = "Audit"
    SheetName = "Enrollments"
    MeasureSheetName = "SelectedMeasures"
    PMSheetName = "PM"
    SettingSN = "Setting"
    currentEnrollment = ""
    currentrow = 0
    Call lstpopulate

    premiseid = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Premise_ID).Value
    accountnumber = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Account_Number).Value
    ROSAWONumber = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.WO_Number_ROSA).Value
    HEAPWONumber = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.WO_Number_HEAP).Value
    
    txtEnrollmentID.Text = currentEnrollment
    txtPremiseID.Text = premiseid
    txtAccountNumber.Text = accountnumber
    txtEnrollmentID.Enabled = False
    txtPremiseID.Enabled = False
    txtAccountNumber.Enabled = False


    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

Private Sub cmdProjectInfo_Click()
    Me.Hide
    Information_Form.Show vbModeless
End Sub
