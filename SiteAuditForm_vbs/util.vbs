Attribute VB_Name = "util"
Public EnrollmentFirstDataLine As Integer
Public SheetName As String
Public PMSheetName As String
Public InboundLastReadCol As Integer
Public currentEnrollment As String
Public premiseid As String
Public accountnumber As String
Public lastrow As Integer
Public ir As Integer


Sub Main()
    frmMain.Show vbModeless
End Sub

'todo list
'validation
'other systems
Sub showworkbook()
Attribute showworkbook.VB_ProcData.VB_Invoke_Func = "g\n14"
    Application.Visible = True
End Sub
