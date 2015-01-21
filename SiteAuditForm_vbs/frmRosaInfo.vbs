VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRosaInfo 
   Caption         =   "ROSA audit"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10620
   OleObjectBlob   =   "frmRosaInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRosaInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
    frmMain.Show vbModeless
End Sub

Private Sub cmdSave_Click()
    Call RosaDataValidate
    Call writerosa
    MsgBox "The ROSA info is saved."
    Me.Hide
    frmMain.Show vbModeless
End Sub
Private Sub writerosa()
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Air_Leakage_Rating_ROSA).Value = txtAirLeakageRating.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Blower_door_pre_test_ROSA).Value = txtPreTest.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Blower_door_post_test_ROSA).Value = txtPostTest.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Number_of_Auditors_ROSA).Value = cboAuditorNumber.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.First_and_last_name_of_main_Auditor_ROSA).Value = txtAuditorName.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Business_Partner_Number_ROSA).Value = txtBPN.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Total_conditioned_square_footage_ROSA).Value = txtConditionedSQFT.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Building_occupancy_count_ROSA).Value = cboOccupancyCount.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Occupancy_frequency_ROSA).Value = txtOccupancyFrequency.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Number_of_stories_above_grade_ROSA).Value = cboStoriesAboveGrade.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Ownership_Type_ROSA).Value = cboOwnership.Text
    If optDogCatYes Then
        Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Dog_or_Cat_Flag_ROSA).Value = "X"
    Else
        Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Dog_or_Cat_Flag_ROSA).Value = ""
    End If
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Auditor_Notes_ROSA).Value = txtAuditorNotes.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Comments_ROSA).Value = txtAuditorComments.Text
    Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.FILE_NAME_ROSA).Value = txtFilename.Text
End Sub

Private Sub readrosa()
txtAirLeakageRating.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Air_Leakage_Rating_ROSA).Value
txtPreTest.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Blower_door_pre_test_ROSA).Value
txtPostTest.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Blower_door_post_test_ROSA).Value
cboAuditorNumber.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Number_of_Auditors_ROSA).Value
txtAuditorName.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.First_and_last_name_of_main_Auditor_ROSA).Value
txtBPN.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Business_Partner_Number_ROSA).Value
txtConditionedSQFT.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Total_conditioned_square_footage_ROSA).Value
cboOccupancyCount.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Building_occupancy_count_ROSA).Value
txtOccupancyFrequency.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Occupancy_frequency_ROSA).Value
cboStoriesAboveGrade.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Number_of_stories_above_grade_ROSA).Value
cboOwnership.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Ownership_Type_ROSA).Value
    If Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Dog_or_Cat_Flag_ROSA).Value = "X" Then
        optDogCatYes.Value = True
    Else
        optDogCatYes.Value = False
    End If
txtAuditorNotes.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Auditor_Notes_ROSA).Value
txtAuditorComments.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Comments_ROSA).Value
txtFilename.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.FILE_NAME_ROSA).Value

End Sub
Private Sub RosaDataValidate()
End Sub

Private Sub UserForm_Activate()
    txtScheduleDate.Enabled = False
    txtScheduleTime.Enabled = False
    txtScheduleDate.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Schedule_Date_ROSA).Value
    txtScheduleTime.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.Schedule_Time_ROSA).Value
    
'    txtScheduleDate.BackColor = &H80000000
'    txtScheduleTime.BackColor = &H80000000
    
    cboAuditorNumber.Clear
    cboAuditorNumber.AddItem ("1")
    cboAuditorNumber.AddItem ("2")
    cboAuditorNumber.AddItem ("3")
    cboAuditorNumber.AddItem ("4")
    cboAuditorNumber.AddItem ("5")
    cboAuditorNumber.AddItem ("6")
    
    cboOccupancyCount.Clear
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
    
    cboStoriesAboveGrade.Clear
    cboStoriesAboveGrade.AddItem ("1")
    cboStoriesAboveGrade.AddItem ("2")
    cboStoriesAboveGrade.AddItem ("3")
    
    cboOwnership.Clear
    cboOwnership.AddItem ("OWN")
    cboOwnership.AddItem ("RENT")
    
    txtEnrollmentID.Text = currentEnrollment
    txtWONumber.Text = Worksheets(SheetName).Cells(EnrollmentFirstDataLine, NexantEnrollments.WO_Number_ROSA).Value
    txtEnrollmentID.Enabled = False
    txtWONumber.Enabled = False
    
    Call readrosa
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub

