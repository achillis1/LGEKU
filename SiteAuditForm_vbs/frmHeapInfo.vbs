VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHeapInfo 
   Caption         =   "HEAP audit"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10665
   OleObjectBlob   =   "frmHeapInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHeapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
    frmMain.Show vbModeless
End Sub

Private Sub cmdSave_Click()
    Call HeapDataValidate
    Call writeheap
    MsgBox "The ROSA info is saved."
    Me.Hide
    frmMain.Show vbModeless
End Sub

Private Sub writeheap()
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Air_Leakage_Rating_HEAP).Value = txtAirLeakageRating.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Blower_door_pre_test_HEAP).Value = txtPreTest.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Blower_door_post_test_HEAP).Value = txtPostTest.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Number_of_Auditors_HEAP).Value = cboAuditorNumber.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.First_and_last_name_of_main_Auditor_HEAP).Value = txtAuditorName.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Schedule_Date_HEAP).Value = txtScheduleDate.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Schedule_Time_HEAP).Value = txtScheduleTime.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Business_Partner_Number_HEAP).Value = txtBPN.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Total_conditioned_square_footage_HEAP).Value = txtConditionedSQFT.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Building_occupancy_count_HEAP).Value = cboOccupancyCount.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Occupancy_frequency_HEAP).Value = txtOccupancyFrequency.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Number_of_stories_above_grade_HEAP).Value = cboStoriesAboveGrade.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Ownership_Type_HEAP).Value = cboOwnership.Text
    If optDogCatYes Then
        Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Dog_or_Cat_Flag_HEAP).Value = "X"
    Else
        Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Dog_or_Cat_Flag_HEAP).Value = ""
    End If
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Auditor_Notes_HEAP).Value = txtAuditorNotes.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Comments_HEAP).Value = txtAuditorComments.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.FILE_NAME_HEAP).Value = txtFilename.Text

End Sub
Private Sub HeapDataValidate()
    'to do
    '...
End Sub

Private Sub readheap()
txtAirLeakageRating.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Air_Leakage_Rating_HEAP).Value
txtPreTest.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Blower_door_pre_test_HEAP).Value
txtPostTest.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Blower_door_post_test_HEAP).Value
cboAuditorNumber.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Number_of_Auditors_HEAP).Value
txtAuditorName.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.First_and_last_name_of_main_Auditor_HEAP).Value
txtScheduleDate.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Schedule_Date_HEAP).Value
txtScheduleTime.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Schedule_Time_HEAP).Value
txtBPN.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Business_Partner_Number_HEAP).Value
txtConditionedSQFT.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Total_conditioned_square_footage_HEAP).Value
cboOccupancyCount.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Building_occupancy_count_HEAP).Value
txtOccupancyFrequency.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Occupancy_frequency_HEAP).Value
cboStoriesAboveGrade.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Number_of_stories_above_grade_HEAP).Value
cboOwnership.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Ownership_Type_HEAP).Value
    If Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Dog_or_Cat_Flag_HEAP).Value = "X" Then
        optDogCatYes.Value = True
    Else
        optDogCatYes.Value = False
    End If
txtAuditorNotes.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Auditor_Notes_HEAP).Value
   txtAuditorComments.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Comments_HEAP).Value
   txtFilename.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.FILE_NAME_HEAP).Value

End Sub

Private Sub UserForm_Activate()
    txtScheduleDate.Enabled = False
    txtScheduleTime.Enabled = False
    
    cboAuditorNumber.AddItem ("1")
    cboAuditorNumber.AddItem ("2")
    cboAuditorNumber.AddItem ("3")
    cboAuditorNumber.AddItem ("4")
    cboAuditorNumber.AddItem ("5")
    cboAuditorNumber.AddItem ("6")
    
    cboOccupancyCount.AddItem ("1")
    cboOccupancyCount.AddItem ("2")
    cboOccupancyCount.AddItem ("3")
    cboOccupancyCount.AddItem ("4")
    cboOccupancyCount.AddItem ("5")
    cboOccupancyCount.AddItem ("6")
    cboOccupancyCount.AddItem ("7")
    cboOccupancyCount.AddItem ("8")
    cboOccupancyCount.AddItem ("9")
    cboOccupancyCount.AddItem ("10")
    
    cboStoriesAboveGrade.AddItem ("1")
    cboStoriesAboveGrade.AddItem ("2")
    cboStoriesAboveGrade.AddItem ("3")
    
    cboOwnership.AddItem ("OWN")
    cboOwnership.AddItem ("RENT")
    
    txtEnrollmentID.Text = currentEnrollment
    txtWONumber.Text = HEAPWONumber
    txtEnrollmentID.Enabled = False
    txtWONumber.Enabled = False
    
    Call readheap
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub
