Attribute VB_Name = "util"
Public EnrollmentFirstDataLine As Integer
Public SheetName As String 'Enrollments
Public AuditSheetName As String 'Audit
Public PMSheetName As String
Public InboundLastReadCol As Integer
Public currentEnrollment As String
Public premiseid As String
Public accountnumber As String
Public lastrow As Integer
Public auditlastrow As Integer
Public currentrow As Integer
Public auditcurrentrow As Integer

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

Public Function getenrolllastrow() As Long
    Dim lastROSA As Long
    Dim lastHEAP As Long
    
    lastROSA = Worksheets(SheetName).Range("B" & Rows.Count).End(xlUp).Row
    lastHEAP = Worksheets(SheetName).Range("C" & Rows.Count).End(xlUp).Row
    getenrolllastrow = WorksheetFunction.Max(lastROSA, lastHEAP)

End Function

