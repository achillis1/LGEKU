VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Information_Form 
   Caption         =   "Information"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14220
   OleObjectBlob   =   "Information_Form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Information_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Account_Number_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Account_Number) = 12 And IsNumeric(Account_Number) = True Then
Account_Number.BackColor = &H80000005

Else

Account_Number.BackColor = &HFF&
MsgBox ("Account Number does not contain required 12 digits")
Cancel = True

End If


End Sub

Private Sub Cancel_Click()
Me.Hide
frmMain.Show vbModeless

End Sub

Private Sub Clear_Click()
Call UserForm_Initialize

End Sub

Private Sub Company_Acronym_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Company_Acronym.Value = "LGE" Or Company_Acronym.Value = "KU" Then
Company_Acronym.BackColor = &H80000005

Else

Company_Acronym.BackColor = &HFF&
MsgBox ("Company Acronym must be LGE or KU")
Cancel = True

End If

End Sub


Private Sub Customer_Home_Phone_exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Customer_Home_Phone) = 10 And IsNumeric(Customer_Home_Phone) = True Then
Customer_Home_Phone.Text = Format(Customer_Home_Phone, "(000)000-0000")
Customer_Home_Phone.BackColor = &H80000005

Else

If Len(Customer_Home_Phone) = 13 Then

Customer_Home_Phone.Text = Format(Customer_Home_Phone, "(000)000-0000")
Customer_Home_Phone.BackColor = &H80000005

Else


Customer_Home_Phone.BackColor = &HFF&
MsgBox ("Phone number should be in 'XXXXXXXXXX' format")
Cancel = True

End If
End If
End Sub

Private Sub Customer_mobile_phone_exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Customer_mobile_phone) = 10 And IsNumeric(Customer_mobile_phone) = True Then
Customer_mobile_phone.Text = Format(Customer_mobile_phone, "(000)000-0000")
Customer_mobile_phone.BackColor = &H80000005

Else
If Len(Customer_mobile_phone) = 13 Then

Customer_mobile_phone.Text = Format(Customer_mobile_phone, "(000)000-0000")
Customer_mobile_phone.BackColor = &H80000005

Else
Customer_mobile_phone.BackColor = &HFF&
MsgBox ("Phone number should be in 'XXXXXXXXXX' format")
Cancel = True

End If
End If
End Sub



Private Sub Mailing_Zipcode_exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Mailing_Zipcode) = 5 And IsNumeric(Mailing_Zipcode) Then
Mailing_Zipcode.BackColor = &H80000005

Else

If Len(Mailing_Zipcode) = 10 And IsNumeric(left(Mailing_Zipcode, 5)) = True And IsNumeric(Right(Mailing_Zipcode, 4)) = True Then
Mailing_Zipcode.BackColor = &H80000005

Else

Cancel = True
Mailing_Zipcode.BackColor = &HFF&
MsgBox ("Mailing Zipcode code must be in 'XXXXX' or 'XXXXX-XXXX' format")


End If

End If

End Sub



Private Sub Premise_ID_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Premise_ID) = 10 And IsNumeric(Premise_ID) = True Then
Premise_ID.BackColor = &H80000005

Else

Premise_ID.BackColor = &HFF&
MsgBox ("Premise ID does not contain required 10 digits")
Cancel = True

End If

End Sub

Private Sub Premise_Type_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Premise_Type = "R" Then
Premise_Type.BackColor = &H80000005

Else

Premise_Type.BackColor = &HFF&
MsgBox ("Premise Type is not R")
Cancel = True

End If

End Sub


Private Sub Primary_Contact_Address_Zip_exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Primary_Contact_Address_Zip) = 5 And IsNumeric(Primary_Contact_Address_Zip) Then
Primary_Contact_Address_Zip.BackColor = &H80000005

Else

If Len(Primary_Contact_Address_Zip) = 10 And IsNumeric(left(Primary_Contact_Address_Zip, 5)) = True And IsNumeric(Right(Primary_Contact_Address_Zip, 4)) = True Then
Primary_Contact_Address_Zip.BackColor = &H80000005

Else


Primary_Contact_Address_Zip.BackColor = &HFF&
MsgBox ("Mailing Zipcode code must be in 'XXXXX' or 'XXXXX-XXXX' format")
Cancel = True

End If

End If

End Sub

Private Sub Primary_Contact_mobile_phone_exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Primary_Contact_mobile_phone) = 10 And IsNumeric(Primary_Contact_mobile_phone) = True Then
Primary_Contact_mobile_phone.Text = Format(Primary_Contact_Phone, "(000)000-0000")
Primary_Contact_mobile_phone.BackColor = &H80000005

Else
If Len(Primary_Contact_mobile_phone) = 13 Then

Primary_Contact_mobile_phone.Text = Format(Primary_Contact_mobile_phone, "(000)000-0000")
Primary_Contact_mobile_phone.BackColor = &H80000005

Else

Primary_Contact_mobile_phone.BackColor = &HFF&
MsgBox ("Phone number should be in 'XXXXXXXXXX' format")
Cancel = True

End If
End If
End Sub

Private Sub Primary_Contact_Phone_exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Primary_Contact_Phone) = 10 And IsNumeric(Primary_Contact_Phone) = True Then
Primary_Contact_Phone.Text = Format(Primary_Contact_Phone, "(000)000-0000")
Primary_Contact_Phone.BackColor = &H80000005

Else
If Len(Primary_Contact_Phone) = 13 Then

Primary_Contact_Phone.Text = Format(Primary_Contact_Phone, "(000)000-0000")
Primary_Contact_Phone.BackColor = &H80000005

Else
Primary_Contact_Phone.BackColor = &HFF&
MsgBox ("Phone number should be in 'XXXXXXXXXX' format")
Cancel = True

End If
End If
End Sub




Private Sub Remit_to_Contact_Address_Zip_exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Remit_to_Contact_Address_Zip) = 5 And IsNumeric(Remit_to_Contact_Address_Zip) Then
Remit_to_Contact_Address_Zip.BackColor = &H80000005

Else

If Len(Remit_to_Contact_Address_Zip) = 10 And IsNumeric(left(Remit_to_Contact_Address_Zip, 5)) = True And IsNumeric(Right(Remit_to_Contact_Address_Zip, 4)) = True Then
Remit_to_Contact_Address_Zip.BackColor = &H80000005

Else


Remit_to_Contact_Address_Zip.BackColor = &HFF&
MsgBox ("Mailing Zipcode code must be in 'XXXXX' or 'XXXXX-XXXX' format")
Cancel = True

End If

End If

End Sub

Private Sub Remit_to_Contact_mobile_phone_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Remit_to_Contact_mobile_phone) = 10 And IsNumeric(Remit_to_Contact_mobile_phone) = True Then
Remit_to_Contact_mobile_phone.Text = Format(Remit_to_Contact_mobile_phone, "(000)000-0000")
Remit_to_Contact_mobile_phone.BackColor = &H80000005

Else
If Len(Remit_to_Contact_mobile_phone) = 13 Then

Remit_to_Contact_mobile_phone.Text = Format(Remit_to_Contact_mobile_phone, "(000)000-0000")
Remit_to_Contact_mobile_phone.BackColor = &H80000005

Else
Remit_to_Contact_mobile_phone.BackColor = &HFF&
MsgBox ("Phone number should be in 'XXXXXXXXXX' format")
Cancel = True

End If
End If
End Sub
Private Sub Remit_to_Contact_Phone_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Remit_to_Contact_Phone) = 10 And IsNumeric(Remit_to_Contact_Phone) = True Then
Remit_to_Contact_Phone.Text = Format(Remit_to_Contact_Phone, "(000)000-0000")
Remit_to_Contact_Phone.BackColor = &H80000005

Else
If Len(Remit_to_Contact_Phone) = 13 Then

Remit_to_Contact_Phone.Text = Format(Remit_to_Contact_Phone, "(000)000-0000")
Remit_to_Contact_Phone.BackColor = &H80000005

Else
Remit_to_Contact_Phone.BackColor = &HFF&
MsgBox ("Phone number should be in 'XXXXXXXXXX' format")
Cancel = True

End If
End If
End Sub



Private Sub Service_Zipcode_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(Service_Zipcode) = 5 And IsNumeric(Service_Zipcode) Then
Service_Zipcode.BackColor = &H80000005

Else

If Len(Service_Zipcode) = 10 And IsNumeric(left(Service_Zipcode, 5)) = True And IsNumeric(Right(Service_Zipcode, 4)) = True Then
Service_Zipcode.BackColor = &H80000005

Else


Service_Zipcode.BackColor = &HFF&
MsgBox ("Mailing Zipcode code must be in 'XXXXX' or 'XXXXX-XXXX' format")
Cancel = True

End If

End If

End Sub

Private Sub UserForm_Initialize()

Dim TimeandDate As String
Dim Result As Double
Dim Enrollment_ID_ROSA As String
Dim ColumnHeadings(46) As String
Dim ColumnValues(46) As String
Dim counter As Double
Dim lastrow As String
Dim HeadingOffset As Double

HeadingOffset = 10

Enrollment_ID_ROSA = currentEnrollment
currentrow = 1
'currentrow = EnrollmentFirstDataLine
If Enrollment_ID_ROSA <> "" Then
lastrow = Cells(Rows.Count, 2).End(xlUp).Row
'Result = Application.Match(Enrollment_ID_ROSA, Sheets("Enrollments").Range(Cells(currentrow, NexantEnrollments.Enrollment_ID_ROSA), Cells(lastrow, NexantEnrollments.Enrollment_ID_ROSA))) + HeadingOffset
Result = 11
'Result = Application.Match(Enrollment_ID_ROSA, Sheets("Enrollments").Range("B1", "B" & Range("B" & Rows.Count).End(xlUp).Row), 0)
Else

End If

If Result <> 0 Then

Account_Number.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Account_Number)
CFM_Reduction.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.CFM_Reduction)
Company_Acronym.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Company_Acronym)
Company_Code.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Company_Code)
Contact_Name.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Contact_Name)
Customer_Email.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Customer_Email)
Customer_Home_Phone.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Customer_Home_Phone)
Customer_mobile_phone.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Customer_mobile_phone)
Mailing_Zipcode.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Mailing_Zipcode)
Mailing_City.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Mailing_City)
Mailing_State.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Mailing_State)
Mailing_Street_Address.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Mailing_Street_Address)
Main_Account_Flag.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Main_Account_Flag)
Premise_ID.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Premise_ID)
Premise_Type.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Premise_Type)
Primary_Contact_Address.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_Contact_Address)
Primary_Contact_Address_City.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_Contact_Address_City)
Primary_Contact_Address_State.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_Contact_Address_State)
Primary_Contact_Address_Zip.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_Contact_Address_Zip)
Primary_Contact_Email.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_Contact_Email)
Primary_Contact_mobile_phone.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_Contact_mobile_phone)
Primary_contact_name.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_contact_name)
Primary_Contact_Phone.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_Contact_Phone)
Primary_Contact_phone_extension.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Primary_Contact_phone_extension)
Reason_for_audit.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Reason_for_audit)
Remit_to_Contact_Address.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_Address)
Remit_to_Contact_Address_City.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_Address_City)
Remit_to_Contact_Address_State.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_Address_State)
Remit_to_Contact_Address_Zip.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_Address_Zip)
Remit_to_Contact_Email.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_Email)
Remit_to_Contact_mobile_phone.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_mobile_phone)
Remit_to_Contact_Name.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_Name)
Remit_to_Contact_Phone.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_Phone)
Remit_to_Contact_phone_extension.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Remit_to_Contact_phone_extension)
Residence_Building_Class.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Residence_Building_Class)
Residence_Building_Type.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Residence_Building_Type)
Send_Report_to_Primary_Contact.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Send_Report_to_Primary_Contact)
Service_City.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Service_City)
'Service_Class.Value = Sheets("Enrollments").Cells(result,NexantEnrollments.Service_Class)
Service_customer_name.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Service_customer_name)
Service_State.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Service_State)
Service_Street_Address.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Service_Street_Address)
Service_Zipcode.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Service_Zipcode)
Verification_Class.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Verification_Class)
VKONT.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.VKONT)
Year_building_constructed.Value = Sheets("Enrollments").Cells(Result, NexantEnrollments.Year_building_constructed)

End If

End Sub

Private Sub VKONT_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(VKONT) = 12 Then
VKONT.BackColor = &H80000005

Else

VKONT.BackColor = &HFF&
MsgBox ("VKONT must contain 12 digits")
Cancel = True

End If

End Sub
