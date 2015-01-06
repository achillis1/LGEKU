VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Main"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5100
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdMeasure_Click()
    Me.Hide
    frmMeasure.Show
End Sub

Private Sub cmdSystem_Click()
    Me.Hide
    frmSystem.Show
End Sub

Private Sub UserForm_Initialize()
'    Application.Visible = False
    EnrollmentFirstDataLine = 11
    ImportSheetName = "Enrollments"
    PMSheetName = "PM"
    InboundLastReadCol = 5
    currentEnrollment = ""
    
    ROSAID = Worksheets(ImportSheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Enrollment_ID_ROSA).Value
    HEAPID = Worksheets(ImportSheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Enrollment_ID_HEAP).Value
    
    If ROSAID = "" And HEAPID <> "" Then
        currentEnrollment = HEAPID
    End If
    
    If HEAPID = "" And ROSAID <> "" Then
        currentEnrollment = ROSAID
    End If
    
    If (HEAPID = "" And ROSAID = "") Or (HEAPID <> "" And ROSAID <> "") Then
        MsgBox "The enrollment ID is not valid. Please check the Enrollments sheet."
        Exit Sub
    End If
    

    premiseid = Worksheets(ImportSheetName).Cells(EnrollmentFirstDataLine, Premise_ID).Value
    accountnumber = Worksheets(ImportSheetName).Cells(EnrollmentFirstDataLine, Account_Number).Value
    
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
    Information_Form.Show
End Sub
