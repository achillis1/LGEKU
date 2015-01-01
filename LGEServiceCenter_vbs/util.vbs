Attribute VB_Name = "util"
Dim scWB As Workbook
Dim enrollWB As Workbook

Sub test1()
    frmServiceCenter.Show

End Sub
Sub ImportEnrollmentFile()
On Error Resume Next
    Dim thisWB As Workbook
    Dim importWB As Workbook
    Set thisWB = ActiveWorkbook

    'clear reading
'    lastrow = thisWB.Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
'    Worksheets("Sheet1").Range("A1:AV" & lastrow).Clear

'    If lastrow > 1 Then
'        lastrow = lastrow + 1
'    End If
    
    filetoopen = Application.GetOpenFilename(FileFilter:="OUT Files (*.txt), *.txt", Title:="Select OUT files")

    If filetoopen = False Then
        exit_macro1 = True
        exit_macro_reason = "User canceled during file selection."
        Exit Sub
    End If
    Workbooks.Open Filename:=filetoopen, Format:=1
    
    slashloc = InStrRev(filetoopen, "\")
    Filename = Mid(filetoopen, slashloc + 1, Len(filetoopen) - slashloc - 4)
    'C:\share\DriveZ\LGE\documents\HEAP\DSM_HEAP_ENROLL_OUT.TXT"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" + filetoopen, Destination:= _
        Cells(1, 1))
        .Name = Filename
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 936
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 2, 2, 2, 1, 2, 1, 2, 1, 1, 2, 2, 1, 1, 1, 2, 1, 1, 1, 2, 1, _
        1, 1, 1, 1, 2, 1, 2, 1, 1, 2, 2, 1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 2, 1, 2, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Set importWB = ActiveWorkbook
    
    lastrow = importWB.Worksheets(Filename).Range("A" & Rows.Count).End(xlUp).Row
    
    Dim existingdata As Range
    Set existingdata = importWB.Worksheets(Filename).Range("a2:AV" & lastrow - 1)
    
    Dim rng As Range
    lastenroll = thisWB.Worksheets("Enrollment").Range("A" & Rows.Count).End(xlUp).Row
    Set rng = thisWB.Worksheets("Enrollment").Range("A" & lastenroll + 1)
    existingdata.Copy
    rng.PasteSpecial xlValues
    
End Sub
