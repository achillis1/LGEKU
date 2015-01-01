VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AuditForm 
   Caption         =   "Add Systems"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12615
   OleObjectBlob   =   "AuditForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AuditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iHeating As Integer
Private iCooling As Integer
Private iHVAC As Integer
Private iWH As Integer
Private iThermostat As Integer
Private iWindow As Integer
Private iDoor As Integer
Private iLighting As Integer
Private iWall As Integer
Private iAttic As Integer
Private iBasement As Integer
Private iBW As Integer
Private iRefrigerator As Integer
Private iFreezer As Integer
Private iApplicance As Integer

Private thisWB As Workbook
Private prompt As String

Private SystemType As Object
Private lblSystemType As Object
Private FuelSource As Object
Private lblFuelSource As Object
Private lblSystemSize As Object
Private SystemSize As Object
Private lblSizeUnit As Object
Private SizeUnit As Object
Private lblSystemAge As Object
Private SystemAge As Object
Private lblEffRating As Object
Private EffRating As Object
Private lblEffRatingType As Object
Private EffRatingType As Object
Private lblPercentageCooled As Object
Private PercentageCooled As Object
Private lblFrequencyUse As Object
Private FrequencyUse As Object
Private lblTotalUnits As Object
Private TotalUnits As Object
Private lblQuantity As Object
Private Quantity As Object

Private strCurrentSystemName As String


Private Sub cboSystem_Change()
    'remove all dynamic controls on selection change
    For Each ctrl In AuditForm.Controls
        If ctrl.Name = "lblAddSystem" Or ctrl.Name = "cboSystem" Or ctrl.Name = "cmdOK" Or ctrl.Name = "cmdCancel" _
            Or ctrl.Name = "txtNote" Or ctrl.Name = "lstSelectedSystems" Or ctrl.Name = "cmdRemove" Or ctrl.Name = "cmdNew" _
            Or ctrl.Name = "cmdLoad" Or ctrl.Name = "cmdRename" Or ctrl.Name = "lblEnteredSystems" Then
        Else
            AuditForm.Controls.Remove (ctrl.Name)
        End If
    Next
    
    Select Case cboSystem.Text
        Case "HEATING"
            Call showheatingoptions
        Case "COOLING"
            Call showcoolingoptions
'        Case "HVAC DISTRIBUTION"
'            Call showhvacoptions
'        Case "WATER HEATER"
'            Call showwhoptions
'        Case "THERMOSTAT"
'            Call showthermostatoptions
'        Case "WINDOW"
'            Call showwindowoptions
'        Case "DOOR"
'            Call showdooroptions
'        Case "LIGHTING"
'            Call showlightingoptions
'        Case "WALL"
'            Call showwalloptions
'        Case "ATTIC"
'            Call showatticoptions
'        Case "BASEMENT"
'            Call showbasementoptions
'        Case "BASEMENT WALL"
'            Call showbwoptions
'        Case "REFRIGERATOR"
'            Call showrefrigeratoroptions
'        Case "FREEZER"
'            Call showfreezeroptions
'        Case "APPLIANCE"
'            Call showapplianceoptions
    End Select
    cmdOK.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    AuditForm.Hide
    Application.Visible = True
    ThisWorkbook.Saved = True
    ThisWorkbook.Close SaveChanges:=True
    Application.Quit
    
End Sub

Private Sub addcomboBox(ByVal cboBox As Object, ByVal cboname As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer)

    For Each ctrl In AuditForm.Controls
        If ctrl.Name = cboname Then
            AuditForm.Controls.Remove (cboname)
        End If
    Next

    Set cboBox = AuditForm.Controls.Add("Forms.ComboBox.1", cboname, True)
    With cboBox
        .top = top
        .left = left
        .width = width
    End With

End Sub

Private Sub addlabel(ByVal label As Object, ByVal labelname As String, ByVal caption As String, ByVal top As Integer, ByVal left As Integer)

    For Each ctrl In AuditForm.Controls
        If ctrl.Name = labelname Then
            AuditForm.Controls.Remove (labelname)
        End If
    Next
    
    Set label = AuditForm.Controls.Add("Forms.Label.1", labelname, True)
    With label
        .caption = caption
        .top = top
        .left = left
    End With
End Sub


Private Sub addtextbox(ByVal textbox As Object, ByVal textboxname As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer)

    For Each ctrl In AuditForm.Controls
        If ctrl.Name = textboxname Then
            AuditForm.Controls.Remove (textboxname)
        End If
    Next
    Set textbox = AuditForm.Controls.Add("Forms.TextBox.1", textboxname, True)
    With textbox
        .top = top
        .left = left
        .width = width
    End With
End Sub

Private Sub showcoolingoptions()
    ' COOLING TYPE
    Call addlabel(lblSystemType, "lblSystemType1", "Cooling Type*", 62, 15)
    Call addcomboBox(SystemType, "SystemType1", 62, 80, 150)
    AuditForm.Controls("SystemType1").AddItem ("CENTRAL AC")
    AuditForm.Controls("SystemType1").AddItem ("HEAT PUMP-AIR SOURCE")
    AuditForm.Controls("SystemType1").AddItem ("HEAT PUMP-WATER SOURCE")
    AuditForm.Controls("SystemType1").AddItem ("SPLIT SYSTEM")
    AuditForm.Controls("SystemType1").AddItem ("WINDOW AC")
    
    ' FUEL SOURCE
    Call addlabel(lblFuelSource, "lblFuelSource1", "Fuel Source*", 92, 15)
    Call addcomboBox(FuelSource, "FuelSource1", 92, 80, 150)
    AuditForm.Controls("FuelSource1").AddItem ("ELECTRIC")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "lblSystemSize1", "System Size*", 122, 15)
    Call addtextbox(SystemSize, "SystemSize1", 122, 80, 50)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "lblSizeUnit1", "System Size Unit*", 122, 150)
    Call addcomboBox(SizeUnit, "SizeUnit1", 122, 220, 70)
    AuditForm.Controls("SizeUnit1").AddItem ("BTU")
    AuditForm.Controls("SizeUnit1").AddItem ("MBTU")
    AuditForm.Controls("SizeUnit1").AddItem ("MMBTU")
    AuditForm.Controls("SizeUnit1").AddItem ("TON")

    'SYSTEM Age
    Call addlabel(lblSystemAge, "lblSystemAge1", "System Age*", 152, 15)
    Call addtextbox(SystemAge, "SystemAge1", 152, 80, 50)
    
    'SYSTEM Efficiency Rating
    Call addlabel(lblEffRating, "lblEffRating1", "Efficiency Rating*", 182, 5)
    Call addtextbox(EffRating, "EffRating1", 182, 80, 50)
    
    'SYSTEM Efficiency Rating Type
    Call addlabel(lblEffRatingType, "lblEffRatingType1", "Rating Type*", 182, 150)
    Call addcomboBox(EffRatingType, "EffRatingType1", 182, 220, 80)
    AuditForm.Controls("EffRatingType1").AddItem ("EER")
    AuditForm.Controls("EffRatingType1").AddItem ("SEER")
    AuditForm.Controls("EffRatingType1").AddItem ("COP")
    


    'TOTAL PERCENTAGE OF SPACE COOLED
    Call addlabel(lblPercentageCooled, "lblPercentageCooled1", "% of space cooled*", 210, 5)
    Call addtextbox(PercentageCooled, "PercentageCooled1", 210, 80, 50)
    
    'FREQUENCY OF SYSTEM USE

    Call addlabel(lblFrequencyUse, "lblFrequencyUse1", "Frequency of use*", 210, 150)
    Call addcomboBox(FrequencyUse, "FrequencyUse1", 210, 220, 70)
    AuditForm.Controls("FrequencyUse1").AddItem ("0%")
    AuditForm.Controls("FrequencyUse1").AddItem ("10-30%")
    AuditForm.Controls("FrequencyUse1").AddItem ("31-70%")
    AuditForm.Controls("FrequencyUse1").AddItem ("71-100%")
    
    'TOTAL UNITS USED
    Call addlabel(lblTotalUnits, "lblTotalUnits1", "Total units used", 240, 5)
    Call addtextbox(TotalUnits, "TotalUnits1", 240, 80, 50)
        
    'QUANTITY
    Call addlabel(lblQuantity, "lblQuantity1", "Quantity", 270, 5)
    Call addtextbox(Quantity, "Quantity1", 270, 80, 50)
        
End Sub
'        Case "HVAC DISTRIBUTION"
'            Call showhvacoptions
'        Case "WATER HEATER"
'            Call showwhoptions
'        Case "THERMOSTAT"
'            Call showthermostatoptions
'        Case "WINDOW"
'            Call showwindowoptions
'        Case "DOOR"
'            Call showdooroptions
'        Case "LIGHTING"
'            Call showlightingoptions
'        Case "WALL"
'            Call showwalloptions
'        Case "ATTIC"
'            Call showatticoptions
'        Case "BASEMENT"
'            Call showbasementoptions
'        Case "BASEMENT WALL"
'            Call showbwoptions
'        Case "REFRIGERATOR"
'            Call showrefrigeratoroptions
'        Case "FREEZER"
'            Call showfreezeroptions
'        Case "APPLIANCE"
'            Call showapplianceoptions

Private Sub showheatingoptions()
    
    ' HEATING TYPE
    Call addlabel(lblSystemType, "lblSystemType1", "Heating Type*", 62, 15)
    Call addcomboBox(SystemType, "SystemType1", 62, 80, 150)
    AuditForm.Controls("SystemType1").AddItem ("GAS FURNACE")
    AuditForm.Controls("SystemType1").AddItem ("HEAT PUMP-AIR SOURCE")
    AuditForm.Controls("SystemType1").AddItem ("HEAT PUMP-GROUND SOURCE")
    AuditForm.Controls("SystemType1").AddItem ("HEAT PUMP-DUAL FUEL")
    AuditForm.Controls("SystemType1").AddItem ("RESISTANCE ELECTRIC HEAT")
    AuditForm.Controls("SystemType1").AddItem ("HOT WATER BOILER")
    AuditForm.Controls("SystemType1").AddItem ("FORCED AIR")
    AuditForm.Controls("SystemType1").AddItem ("STEAM")
    AuditForm.Controls("SystemType1").AddItem ("WOOD/COAL STOVE")
    
    ' FUEL SOURCE
    Call addlabel(lblFuelSource, "lblFuelSource1", "Fuel Source*", 92, 15)
    Call addcomboBox(FuelSource, "FuelSource1", 92, 80, 150)
    AuditForm.Controls("FuelSource1").AddItem ("ELECTRIC")
    AuditForm.Controls("FuelSource1").AddItem ("GAS")
    AuditForm.Controls("FuelSource1").AddItem ("PROPANE")
    AuditForm.Controls("FuelSource1").AddItem ("CENTRAL STEAM")
    AuditForm.Controls("FuelSource1").AddItem ("COAL")
    AuditForm.Controls("FuelSource1").AddItem ("SOLAR")
    AuditForm.Controls("FuelSource1").AddItem ("WOOD")
    AuditForm.Controls("FuelSource1").AddItem ("OIL")
    AuditForm.Controls("FuelSource1").AddItem ("OTHER")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "lblSystemSize1", "System Size*", 122, 15)
    Call addtextbox(SystemSize, "SystemSize1", 122, 80, 50)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "lblSizeUnit1", "System Size Unit*", 122, 150)
    Call addcomboBox(SizeUnit, "SizeUnit1", 122, 220, 70)
    AuditForm.Controls("SizeUnit1").AddItem ("MBTU")
    AuditForm.Controls("SizeUnit1").AddItem ("MMBTU")
    AuditForm.Controls("SizeUnit1").AddItem ("TON")

    'SYSTEM Age
    Call addlabel(lblSystemAge, "lblSystemAge1", "System Age", 152, 15)
    Call addtextbox(SystemAge, "SystemAge1", 152, 80, 50)
    
    'SYSTEM Efficiency Rating
    Call addlabel(lblEffRating, "lblEffRating1", "Efficiency Rating", 182, 5)
    Call addtextbox(EffRating, "EffRating1", 182, 80, 50)
    
    'SYSTEM Efficiency Rating Type
    Call addlabel(lblEffRatingType, "lblEffRatingType1", "Rating Type*", 182, 150)
    Call addcomboBox(EffRatingType, "EffRatingType1", 182, 220, 80)
    AuditForm.Controls("EffRatingType1").AddItem ("AFUE")
    AuditForm.Controls("EffRatingType1").AddItem ("HSPF")
    AuditForm.Controls("EffRatingType1").AddItem ("COP")
End Sub

Private Sub errorstring(ByVal str1 As String)
    If prompt <> "" Then
        prompt = prompt + ", " + str1
    Else
        prompt = str1
    End If

End Sub

Private Function heatingvalidation() As Boolean
    Dim iReply As Integer

    If cboSystem.ListIndex < 0 Then
        prompt = "Heating system"
    End If
    
    stv = AuditForm.Controls("SystemType1").Value
    If stv = "GAS FURNACE" Or stv = "HEAT PUMP-AIR SOURCE" Or stv = "HEAT PUMP-GROUND SOURCE" _
        Or stv = "HEAT PUMP-DUAL FUEL" Or stv = "RESISTANCE ELECTRIC HEAT" Or stv = "HOT WATER BOILER" _
        Or stv = "FORCED AIR" Or stv = "STEAM" Or stv = "WOOD/COAL STOVE" Then
    Else
        errorstring ("System Type")
    End If
    
    fs = AuditForm.Controls("FuelSource1").Value
    If fs = "ELECTRIC" Or fs = "GAS" Or fs = "PROPANE" Or fs = "CENTRAL STEAM" Or fs = "COAL" Or fs = "SOLAR" _
        Or fs = "WOOD" Or fs = "OIL" Or fs = "OTHER" Then
    Else
        errorstring ("Fuel Source")
    End If
    
    If Not IsNumeric(AuditForm.Controls("SystemSize1").Value) Then
        errorstring ("System Size")
    End If

    su = AuditForm.Controls("SizeUnit1").Value
    If su = "MBTU" Or su = "MMBTU" Or su = "TON" Then
    Else
        errorstring ("Size Unit")
    End If
    
    If IsNumeric(AuditForm.Controls("SystemAge1").Value) Or AuditForm.Controls("SystemAge1").Value = "" Then
    Else
        errorstring ("System Age")
    End If
    
    If Not IsNumeric(AuditForm.Controls("EffRating1").Value) Then
        errorstring ("Efficiency Rating")
    End If
    
    et = AuditForm.Controls("EffRatingType1").Value
    If et = "AFUE" Or et = "HSPF" Or et = "COP" Then
    Else
        errorstring ("Efficiency Rating Type")
    End If
    
    If prompt <> "" Then
        iReply = MsgBox(prompt + " not filled out correctly", vbOKOnly, "Input error!")
        prompt = ""
        heatingvalidation = 0
        Exit Function
    Else
    heatingvalidation = 1
    End If
End Function

Private Sub cmdLoad_Click()
    Dim strSystem As String
    If lstSelectedSystems.ListIndex = -1 Then
        iReply = MsgBox("Please select the system to load", vbOKOnly, "Select a system!")
        Exit Sub
    End If
        
    ir = lstSelectedSystems.ListIndex
    strSystem = Worksheets("Audit").Cells(ir + 2, 5).Value
    cboSystem.Text = strSystem
    strCurrentSystemName = Worksheets("Audit").Cells(ir + 2, 1).Value
    Select Case strSystem
        Case "HEATING"
            AuditForm.Controls("SystemType1").Value = Worksheets("Audit").Cells(ir + 2, 7)
            AuditForm.Controls("FuelSource1").Value = Worksheets("Audit").Cells(ir + 2, 8)
            AuditForm.Controls("SystemSize1").Value = Worksheets("Audit").Cells(ir + 2, 11)
            AuditForm.Controls("SizeUnit1").Value = Worksheets("Audit").Cells(ir + 2, 12)
            AuditForm.Controls("SystemAge1").Value = Worksheets("Audit").Cells(ir + 2, 13)
            AuditForm.Controls("EffRating1").Value = Worksheets("Audit").Cells(ir + 2, 15)
            AuditForm.Controls("EffRatingType1").Value = Worksheets("Audit").Cells(ir + 2, 16)
        Case "COOLING"
            AuditForm.Controls("SystemType1").Value = Worksheets("Audit").Cells(ir + 2, 7)
            AuditForm.Controls("FuelSource1").Value = Worksheets("Audit").Cells(ir + 2, 8)
            AuditForm.Controls("SystemSize1").Value = Worksheets("Audit").Cells(ir + 2, 11)
            AuditForm.Controls("SizeUnit1").Value = Worksheets("Audit").Cells(ir + 2, 12)
            AuditForm.Controls("SystemAge1").Value = Worksheets("Audit").Cells(ir + 2, 13)
            AuditForm.Controls("EffRating1").Value = Worksheets("Audit").Cells(ir + 2, 15)
            AuditForm.Controls("EffRatingType1").Value = Worksheets("Audit").Cells(ir + 2, 16)
            AuditForm.Controls("PercentageCooled1").Value = Worksheets("Audit").Cells(ir + 2, 17)
            AuditForm.Controls("FrequencyUse1").Value = Worksheets("Audit").Cells(ir + 2, 18)
            AuditForm.Controls("TotalUnits1").Value = Worksheets("Audit").Cells(ir + 2, 19)
            AuditForm.Controls("Quantity1").Value = Worksheets("Audit").Cells(ir + 2, 14)
        Case "HVAC"
        '....
    End Select
    

End Sub

Private Sub cmdNew_Click()
    cboSystem.Text = ""
    strCurrentSystemName = ""
End Sub

Private Sub cmdOK_Click()
    Dim flag As Boolean
    Select Case cboSystem
        Case "HEATING"
            'flag = heatingvalidation
            If heatingvalidation = True Then
                Call saveheatingsystem
            End If
        Case "COOLING"
            Call savecoolingsystem
    End Select
    
End Sub

Private Sub saveheatingsystem()
    If iHeating < 3 Then
        iHeating = iHeating + 1
        lastrow = Worksheets("Audit").Range("E" & Rows.Count).End(xlUp).Row
        If strCurrentSystemName = "" Then
            strCurrentSystemName = "HEATING-" + CStr(iHeating)
        End If
        Worksheets("Audit").Cells(lastrow + 1, 1) = strCurrentSystemName
        Worksheets("Audit").Cells(lastrow + 1, 5) = "HEATING"
        Worksheets("Audit").Cells(lastrow + 1, 7) = AuditForm.Controls("SystemType1").Value
        Worksheets("Audit").Cells(lastrow + 1, 8) = AuditForm.Controls("FuelSource1").Value
        Worksheets("Audit").Cells(lastrow + 1, 11) = AuditForm.Controls("SystemSize1").Value
        Worksheets("Audit").Cells(lastrow + 1, 12) = AuditForm.Controls("SizeUnit1").Value
        Worksheets("Audit").Cells(lastrow + 1, 13) = AuditForm.Controls("SystemAge1").Value
        Worksheets("Audit").Cells(lastrow + 1, 15) = AuditForm.Controls("EffRating1").Value
        Worksheets("Audit").Cells(lastrow + 1, 16) = AuditForm.Controls("EffRatingType1").Value
        lstSelectedSystems.AddItem (strCurrentSystemName)
    Else
        MsgBox ("You can only enter at most 3 HEATING systems!")
    End If
    
End Sub

Private Sub savecoolingsystem()
    If iCooling < 3 Then
        iCooling = iCooling + 1
        lastrow = Worksheets("Audit").Range("E" & Rows.Count).End(xlUp).Row
        If strCurrentSystemName = "" Then
            strCurrentSystemName = "COOLING-" + CStr(iCooling)
        End If
        Worksheets("Audit").Cells(lastrow + 1, 1) = strCurrentSystemName
        Worksheets("Audit").Cells(lastrow + 1, 5) = "COOLING"
        Worksheets("Audit").Cells(lastrow + 1, 7) = AuditForm.Controls("SystemType1").Value
        Worksheets("Audit").Cells(lastrow + 1, 8) = AuditForm.Controls("FuelSource1").Value
        Worksheets("Audit").Cells(lastrow + 1, 11) = AuditForm.Controls("SystemSize1").Value
        Worksheets("Audit").Cells(lastrow + 1, 12) = AuditForm.Controls("SizeUnit1").Value
        Worksheets("Audit").Cells(lastrow + 1, 13) = AuditForm.Controls("SystemAge1").Value
        Worksheets("Audit").Cells(lastrow + 1, 15) = AuditForm.Controls("EffRating1").Value
        Worksheets("Audit").Cells(lastrow + 1, 16) = AuditForm.Controls("EffRatingType1").Value
        
        Worksheets("Audit").Cells(lastrow + 1, 17) = AuditForm.Controls("PercentageCooled1").Value
        Worksheets("Audit").Cells(lastrow + 1, 18) = AuditForm.Controls("FrequencyUse1").Value
        Worksheets("Audit").Cells(lastrow + 1, 19) = AuditForm.Controls("TotalUnits1").Value
        Worksheets("Audit").Cells(lastrow + 1, 14) = AuditForm.Controls("Quantity1").Value
        
        lstSelectedSystems.AddItem (strCurrentSystemName)
    Else
        MsgBox ("You can only enter at most 3 COOLING systems!")
    End If
    
End Sub

Private Sub cmdRemove_Click()
    ir = lstSelectedSystems.ListIndex
    lstSelectedSystems.RemoveItem (ir)
    Rows(ir + 2).Delete
    Select Case lstSelectedSystems.Text
        Case "HEATING"
            iHeating = iHeating - 1
        Case "COOLING"
            iCooling = iCooling - 1
        Case Else
        
    End Select
End Sub

Private Sub cmdRename_Click()
    Dim strSystem As String
    If lstSelectedSystems.ListIndex = -1 Then
        iReply = MsgBox("Please select the system to rename", vbOKOnly, "Select a system!")
        Exit Sub
    End If
        
    ir = lstSelectedSystems.ListIndex
    strSystem = Worksheets("Audit").Cells(ir + 2, 5).Value
    
    Dim message, title, defaultValue As String
    Dim myValue As String

    message = "Enter the system name"
    title = "System Name"
    defaultValue = "my favoriate system"
    myValue = InputBox(message, title, defaultValue)
    If myValue = "" Then myValue = defaultValue

    strCurrentSystemName = strSystem + "-" + myValue
    Worksheets("Audit").Cells(ir + 2, 1).Value = strCurrentSystemName
    
    lastrow = thisWB.Worksheets("Audit").Range("E" & Rows.Count).End(xlUp).Row
    lstSelectedSystems.Clear
    If lastrow > 1 Then
        For i = 2 To lastrow
            lstSelectedSystems.AddItem (Worksheets("Audit").Cells(i, 1))
        Next i
    End If
End Sub

Private Sub lstSelectedSystems_Change()
    If lstSelectedSystems.ListIndex <> -1 Then
        cmdRemove.Enabled = True
    End If
End Sub


Private Sub UserForm_Initialize()
    Dim rngSystem As Range

    Dim thisWS As Worksheet
    Dim rngItem As Range
    
    Set thisWB = ActiveWorkbook
    
    cboSystem.AddItem ("HEATING")
    cboSystem.AddItem ("COOLING")
    cboSystem.AddItem ("HVAC DISTRIBUTION")
    cboSystem.AddItem ("WATER HEATER")
    cboSystem.AddItem ("THERMOSTAT")
    cboSystem.AddItem ("WINDOW")
    cboSystem.AddItem ("DOOR")
    cboSystem.AddItem ("LIGHTING")
    cboSystem.AddItem ("WALL")
    cboSystem.AddItem ("ATTIC")
    cboSystem.AddItem ("BASEMENT")
    cboSystem.AddItem ("BASEMENT WALL")
    cboSystem.AddItem ("REFRIGERATOR")
    cboSystem.AddItem ("FREEZER")
    cboSystem.AddItem ("APPLIANCE")

    cmdOK.Enabled = False
    cmdCancel.Enabled = True

    iHeating = 0
    iCooling = 0
    iHVAC = 0
    iWH = 0
    iThermostat = 0
    iWindow = 0
    iDoor = 0
    iLighting = 0
    iWall = 0
    iAttic = 0
    iBasement = 0
    iBW = 0
    iRefrigerator = 0
    iFreezer = 0
    iApplicance = 0
   
    
    lastrow = thisWB.Worksheets("Audit").Range("E" & Rows.Count).End(xlUp).Row
    If lastrow > 1 Then
        For i = 2 To lastrow
            lstSelectedSystems.AddItem (Worksheets("Audit").Cells(i, 1))
            Select Case Worksheets("Audit").Cells(i, 5)
                Case "HEATING"
                    iHeating = iHeating + 1
                Case "COOLING"
                    iCooling = iCooling + 1
                Case "HVAC DISTRIBUTION"
                    iHVAC = iHVAC + 1
                Case "WATER HEATER"
                    iWH = iWH + 1
                Case "THERMOSTAT"
                    iThermostat = iThermostat + 1
                Case "WINDOW"
                    iWindow = iWindow + 1
                Case "DOOR"
                    iDoor = iDoor + 1
                Case "LIGHTING"
                    iLighting = iLighting + 1
                Case "WALL"
                    iWall = iWall + 1
                Case "ATTIC"
                    iAttic = iAttic + 1
                Case "BASEMENT"
                    iBasement = iBasement + 1
                Case "BASEMENT WALL"
                    iBW = iBW + 1
                Case "REFRIGERATOR"
                    iRefrigerator = iRefrigerator + 1
                Case "FREEZER"
                    iFreezer = iFreezer + 1
                Case "APPLIANCE"
                    iAppliance = iAppliance + 1
            End Select
        Next i
    End If

    If lstSelectedSystems.ListIndex = -1 Then
        cmdRemove.Enabled = False
    End If
    
    Application.Visible = False
    
End Sub

Private Sub UserForm_Terminate()
    Application.Visible = True
End Sub
