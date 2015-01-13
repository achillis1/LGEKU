VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMeasure 
   Caption         =   "Measure"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11250
   OleObjectBlob   =   "frmMeasure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Me.Hide
    frmMain.Show vbModeless
End Sub

Private Sub cmdLeftToRight_Click()
    If lstMeasuresAvailable.ListIndex <> -1 Then
        measurecurrentrow = selectedmeasurelastrow + 1
        selectedmeasurelastrow = selectedmeasurelastrow + 1
        'lstMeasuresAvailable.ListIndex 2
        measurename = Worksheets(MeasureSheetName).Cells(lstMeasuresAvailable.ListIndex + 2, MCOL.Measure_List).Value
        Worksheets(MeasureSheetName).Cells(measurecurrentrow, MCOL.Selected_Measures).Value = measurename
        lstMeasuresSelected.AddItem (Worksheets(MeasureSheetName).Cells(measurecurrentrow, MCOL.Measure_Name).Value)
    Else
        MsgBox "Please select a measure in the available measure list to proceed."
    End If
End Sub

Private Sub updateselectedlist()
    lstMeasuresSelected.Clear
    For i = 2 To selectedmeasurelastrow
        lstMeasuresSelected.AddItem (Worksheets(MeasureSheetName).Cells(i, MCOL.Measure_Name).Value)
    Next i
End Sub
Private Sub cmdRightToLeft_Click()
    If lstMeasuresSelected.ListIndex <> -1 Then
        For i = lstMeasuresSelected.ListIndex + 2 To selectedmeasurelastrow
            temp = Worksheets(MeasureSheetName).Cells(i + 1, MCOL.Selected_Measures).Value
            Worksheets(MeasureSheetName).Cells(i, MCOL.Selected_Measures).Value = temp
        Next i
        'Worksheets(MeasureSheetName).Cells(i, MCOL.Selected_Measures).Value = ""
        measurecurrentrow = selectedmeasurelastrow - 1
        selectedmeasurelastrow = selectedmeasurelastrow - 1
        
        Call updateselectedlist
    Else
        MsgBox "Please highlight a selected measure..."
    End If
End Sub

Private Sub lstMeasuresSelected_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstMeasuresSelected.ListIndex <> -1 Then
        measurename = Worksheets(MeasureSheetName).Cells(lstMeasuresSelected.ListIndex + 2, MCOL.Selected_Measures).Value
        Select Case measurename
            Case "Cooling System Equipment Improvement"
                Me.Hide
                UserForm1.Show vbModeless
            Case Else
                MsgBox "This measure form is not available yet. Please double click the cooling system equipment measure to show a sample measure form."
        End Select
    Else
        MsgBox "Please highlight a selected measure.."
    End If
End Sub

Private Sub UserForm_Initialize()
    measurelastrow = Worksheets(MeasureSheetName).Range("A" & Rows.Count).End(xlUp).Row
    selectedmeasurelastrow = Worksheets(MeasureSheetName).Range("C" & Rows.Count).End(xlUp).Row
    
    measurecurrentrow = 2
    
    lstMeasuresAvailable.Clear
    For i = 2 To measurelastrow
        lstMeasuresAvailable.AddItem (Worksheets(MeasureSheetName).Range("A" & CStr(i)).Value)
    Next i
    
    Call updateselectedlist
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub
