VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AuditForm 
   Caption         =   "Add Systems"
   ClientHeight    =   5835
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
Private lblInsIndicator As Object
Private InsIndicator As Object
Private lblInsType As Object
Private InsType As Object
Private lblSystemLocation As Object
Private SystemLocation As Object
Private lblSystemLength As Object
Private lblSystemLength1 As Object
Private SystemLength As Object
Private lblFlexCondition As Object
Private FlexCondition As Object

Private strCurrentSystemName As String

Private cboHeight As Integer
Private cboWidth As Integer
Private txtHeight As Integer
Private txtWidth As Integer
Private lblHeight As Integer
Private lblWidth As Integer

Private toTop As Integer
Private toTop1 As Integer
Private toTop2 As Integer
Private toTop3 As Integer
Private toTop4 As Integer
Private toTop5 As Integer
Private toTop6 As Integer
Private toTop7 As Integer
Private toTop8 As Integer
Private toTop9 As Integer

Private toLeft As Integer
Private toLeft1 As Integer
Private toLeft2 As Integer
Private toLeft3 As Integer
'Private toLeft4 As Integer
'Private toLeft5 As Integer
'Private toLeft6 As Integer
'Private toLeft7 As Integer
'Private toLeft8 As Integer
'Private toLeft9 As Integer

Private vertInterval As Integer

Private Sub cboSystem_Change()
    'remove all dynamic controls on selection change
    For Each ctrl In AuditForm.Controls
        If left(ctrl.Name, 3) = "dc_" Then
            AuditForm.Controls.Remove (ctrl.Name)
        End If
    Next
    
    Select Case cboSystem.Text
        Case "HEATING"
            Call showheatingoptions
        Case "COOLING"
            Call showcoolingoptions
        Case "HVAC DISTRIBUTION"
            Call showhvacoptions
        Case "WATER HEATER"
            Call showwhoptions
        Case "THERMOSTAT"
            Call showthermostatoptions
        Case "WINDOW"
            Call showwindowoptions
        Case "DOOR"
            Call showdooroptions
        Case "LIGHTING"
            Call showlightingoptions
        Case "WALL"
            Call showwalloptions
        Case "ATTIC"
            Call showatticoptions
        Case "BASEMENT"
            Call showbasementoptions
        Case "BASEMENT WALL"
            Call showbwoptions
        Case "REFRIGERATOR"
            Call showrefrigeratoroptions
        Case "FREEZER"
            Call showfreezeroptions
        Case "APPLIANCE"
            Call showapplianceoptions
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

Private Sub addcomboBox(ByVal cboBox As Object, ByVal cboname As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer)

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
        .height = height
    End With

End Sub

Private Sub addlabel(ByVal label As Object, ByVal labelname As String, ByVal caption As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer)

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
        .width = width
        .height = height
    End With
End Sub


Private Sub addtextbox(ByVal textbox As Object, ByVal textboxname As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer)

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
        .height = height
    End With
End Sub

Private Sub showwhoptions()
    'WATER HEATER TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Water Heater Type*", toTop, toLeft - 15, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 3, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("CONVENTIONAL STORAGE")
    AuditForm.Controls("dc_SystemType1").AddItem ("DEMAND")
    AuditForm.Controls("dc_SystemType1").AddItem ("TANKLESS\INSTANTANEOUS")
    AuditForm.Controls("dc_SystemType1").AddItem ("SOLAR/TANK")
    AuditForm.Controls("dc_SystemType1").AddItem ("GEOTHERMAL DESUPERHEATER/TANK")
    
End Sub

Private Sub showhvacoptions()
    ' HVAC DISTRIBUTION TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "HVAC Distribution Type*", toTop, toLeft - 10, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("DUCT ROUND")
    AuditForm.Controls("dc_SystemType1").AddItem ("DUCT RECTANGULAR")
    AuditForm.Controls("dc_SystemType1").AddItem ("IRON PIPE")
    AuditForm.Controls("dc_SystemType1").AddItem ("COPPER")
    AuditForm.Controls("dc_SystemType1").AddItem ("ELBOWS")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemSize, "dc_SystemSize1", toTop1, toLeft1, cboWidth, cboHeight)
    AuditForm.Controls("dc_SystemSize1").AddItem ("SMALL")
    AuditForm.Controls("dc_SystemSize1").AddItem ("MEDIUM")
    AuditForm.Controls("dc_SystemSize1").AddItem ("LARGE")
    AuditForm.Controls("dc_SystemSize1").AddItem ("EXTRA LARGE")
    
    'INSULATION EXIST INDICATOR
    Call addlabel(lblInsIndicator, "dc_lblInsIndicator1", "Insulation exist indicator*", toTop2, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsIndicator, "dc_InsIndicator1", toTop2, toLeft1, cboWidth, cboHeight)
    AuditForm.Controls("dc_InsIndicator1").AddItem ("Y")
    AuditForm.Controls("dc_InsIndicator1").AddItem ("N")
    AuditForm.Controls("dc_InsIndicator1").AddItem ("NOT NEEDED")
    
    'INSULATION TYPE
    Call addlabel(lblInsType, "dc_lblInsType1", "Insulation Type*", toTop3, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsType, "dc_InsType1", toTop3, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_InsType1").AddItem ("CELLULOSE")
    AuditForm.Controls("dc_InsType1").AddItem ("FIBERGLASS BATTS")
    AuditForm.Controls("dc_InsType1").AddItem ("FIBERGLASS BLOWN")
    AuditForm.Controls("dc_InsType1").AddItem ("LOOSE FIBERGLASS")
    AuditForm.Controls("dc_InsType1").AddItem ("MINERAL/ROCK WOOL")
    AuditForm.Controls("dc_InsType1").AddItem ("UREA FORMALDAHYDE")
    AuditForm.Controls("dc_InsType1").AddItem (".5 LB FOAM")
    AuditForm.Controls("dc_InsType1").AddItem ("2 LB FOAM")
    AuditForm.Controls("dc_InsType1").AddItem ("NONE")
    AuditForm.Controls("dc_InsType1").AddItem ("OTHER")
    

    'SYSTEM LOCATION
    Call addlabel(lblSystemLocation, "dc_lblSystemLocation1", "System Location*", toTop4, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemLocation, "dc_SystemLocation1", toTop4, toLeft1, cboWidth, cboHeight)
    AuditForm.Controls("dc_SystemLocation1").AddItem ("ATTIC")
    AuditForm.Controls("dc_SystemLocation1").AddItem ("BASEMENT")
    AuditForm.Controls("dc_SystemLocation1").AddItem ("CRAWL")
    
    'LENGTH
    Call addlabel(lblSystemLength, "dc_lblSystemLength1", "Length", toTop5, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemLength, "dc_SystemLength1", toTop5, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblSystemLength1, "dc_lblSystemLength2", "ft", toTop5, toLeft1 + txtWidth + 5, lblWidth, lblHeight)
        
    'CONDITION OF FLEX
    Call addlabel(lblFlexCondition, "dc_lblFlexCondition1", "Flex Condition", toTop6, toLeft, lblWidth, lblHeight)
    Call addcomboBox(FlexCondition, "dc_FlexCondition1", toTop6, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_FlexCondition1").AddItem ("COLLAPSED")
    AuditForm.Controls("dc_FlexCondition1").AddItem ("DAMAGED")
    AuditForm.Controls("dc_FlexCondition1").AddItem ("FUNCTIONAL")
    AuditForm.Controls("dc_FlexCondition1").AddItem ("NON-FUNCTIONAL COLLAPSED")
    AuditForm.Controls("dc_FlexCondition1").AddItem ("NON-FUNCTIONAL DAMAGED")
    
End Sub


Private Sub showthermostatoptions()
    'THERMOSTAT TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Thermostat Type*", toTop, toLeft - 10, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("DIGITAL")
    AuditForm.Controls("dc_SystemType1").AddItem ("DIGITAL PROGRAMMABLE")
    AuditForm.Controls("dc_SystemType1").AddItem ("STANDARD")
End Sub

Private Sub showwindowoptions()
    'WINDOW TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Window Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("SINGLE PANE")
    AuditForm.Controls("dc_SystemType1").AddItem ("SINGLE PANE W/STORM")
    AuditForm.Controls("dc_SystemType1").AddItem ("WINDOW")
    AuditForm.Controls("dc_SystemType1").AddItem ("DOUBLE PANE")
    AuditForm.Controls("dc_SystemType1").AddItem ("TRIPLE PANE")
    AuditForm.Controls("dc_SystemType1").AddItem ("DH")
    AuditForm.Controls("dc_SystemType1").AddItem ("CASEMENT")
    AuditForm.Controls("dc_SystemType1").AddItem ("FIXED")
End Sub

Private Sub showdooroptions()
    'DOOR TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Door Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("METAL/INSULATED")
    AuditForm.Controls("dc_SystemType1").AddItem ("FIBERGLASS/INSULATED")
    AuditForm.Controls("dc_SystemType1").AddItem ("WOOD")
    AuditForm.Controls("dc_SystemType1").AddItem ("SLIDER")
    AuditForm.Controls("dc_SystemType1").AddItem ("ATRIUM")
End Sub

Private Sub showlightingoptions()
    'WINDOW TYPE
    'NOT WINDOW TYPE
End Sub

Private Sub showwalloptions()
    'WALL TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Wall Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("ALUMINUM")
    AuditForm.Controls("dc_SystemType1").AddItem ("BRICK")
    AuditForm.Controls("dc_SystemType1").AddItem ("MASONITE")
    AuditForm.Controls("dc_SystemType1").AddItem ("OTHER")
    AuditForm.Controls("dc_SystemType1").AddItem ("STUCCO")
    AuditForm.Controls("dc_SystemType1").AddItem ("VINYL")
    AuditForm.Controls("dc_SystemType1").AddItem ("WOOD")
End Sub

Private Sub showatticoptions()
    'ATTIC TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Attic Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("FLOORED")
    AuditForm.Controls("dc_SystemType1").AddItem ("UNFLOORED")
    AuditForm.Controls("dc_SystemType1").AddItem ("KNEE WALL")
    AuditForm.Controls("dc_SystemType1").AddItem ("KW FLAT FLOORED")
    AuditForm.Controls("dc_SystemType1").AddItem ("KW FLAT UNFLOORED")
    AuditForm.Controls("dc_SystemType1").AddItem ("FLAT ROOF")
    AuditForm.Controls("dc_SystemType1").AddItem ("SLOPED")
End Sub

Private Sub showbasementoptions()
    'BASEMENT TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Basement Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("CRAWL-OPEN")
    AuditForm.Controls("dc_SystemType1").AddItem ("CRAWL-CLOSED")
    AuditForm.Controls("dc_SystemType1").AddItem ("FULL")
    AuditForm.Controls("dc_SystemType1").AddItem ("GARAGE")
    AuditForm.Controls("dc_SystemType1").AddItem ("SLAB")

End Sub
Private Sub showbwoptions()
    'BASEMENT WALL TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Basement Wall Type*", toTop, toLeft - 15, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("CINDER BLOCK")
    AuditForm.Controls("dc_SystemType1").AddItem ("CONCRETE POUR/FORMED")
    AuditForm.Controls("dc_SystemType1").AddItem ("FRAMED 2x4")
    AuditForm.Controls("dc_SystemType1").AddItem ("FRAMED 2x6")
End Sub
Private Sub showrefrigeratoroptions()
    'REFRIGERATOR TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Refrigerator Type*", toTop, toLeft - 10, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("SIDE BY SIDE")
    AuditForm.Controls("dc_SystemType1").AddItem ("FREEZER TOP")
    AuditForm.Controls("dc_SystemType1").AddItem ("FREEZER BOTTOM")
    AuditForm.Controls("dc_SystemType1").AddItem ("SINGLE DOOR")

End Sub
Private Sub showfreezeroptions()
    'FREEZER TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Freezer Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("UPRIGHT")
    AuditForm.Controls("dc_SystemType1").AddItem ("CHEST")

End Sub
Private Sub showapplianceoptions()
    'APPLIANCE TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Appliance Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("AQUARIUM")
    AuditForm.Controls("dc_SystemType1").AddItem ("ATTIC FAN")
    AuditForm.Controls("dc_SystemType1").AddItem ("BLACK & WHITE TV")
    AuditForm.Controls("dc_SystemType1").AddItem ("CEILING FAN")
    AuditForm.Controls("dc_SystemType1").AddItem ("CLOTHES WASHER")
    AuditForm.Controls("dc_SystemType1").AddItem ("COLOR TV")
    AuditForm.Controls("dc_SystemType1").AddItem ("COMPUTER")
    AuditForm.Controls("dc_SystemType1").AddItem ("DEHUMIDIFIER")
    AuditForm.Controls("dc_SystemType1").AddItem ("DISHWASHER")
    AuditForm.Controls("dc_SystemType1").AddItem ("ELECTRIC SPACE HEATER")
    AuditForm.Controls("dc_SystemType1").AddItem ("ELEC CLOTHES DRYER")
    AuditForm.Controls("dc_SystemType1").AddItem ("ELECTRIC BLANKET")
    AuditForm.Controls("dc_SystemType1").AddItem ("ELECTRIC COOKING")
    AuditForm.Controls("dc_SystemType1").AddItem ("FAX MACHINE")
    AuditForm.Controls("dc_SystemType1").AddItem ("GAS CLOTHES DRYER")
    AuditForm.Controls("dc_SystemType1").AddItem ("GAS COOKING")
    AuditForm.Controls("dc_SystemType1").AddItem ("HOT TUB")
    AuditForm.Controls("dc_SystemType1").AddItem ("HUMIDIFIER")
    AuditForm.Controls("dc_SystemType1").AddItem ("LASER PRINTER")
    AuditForm.Controls("dc_SystemType1").AddItem ("MICROWAVE")
    AuditForm.Controls("dc_SystemType1").AddItem ("MISCELLANEOUS")
    AuditForm.Controls("dc_SystemType1").AddItem ("POOL PUMP")
    AuditForm.Controls("dc_SystemType1").AddItem ("PRINTER")
    AuditForm.Controls("dc_SystemType1").AddItem ("STEREO")
    AuditForm.Controls("dc_SystemType1").AddItem ("SUMP PUMP")
    AuditForm.Controls("dc_SystemType1").AddItem ("WATERBED")
    AuditForm.Controls("dc_SystemType1").AddItem ("WELL PUMP")
End Sub
            
Private Sub showcoolingoptions()
    ' COOLING TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Cooling Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("CENTRAL AC")
    AuditForm.Controls("dc_SystemType1").AddItem ("HEAT PUMP-AIR SOURCE")
    AuditForm.Controls("dc_SystemType1").AddItem ("HEAT PUMP-WATER SOURCE")
    AuditForm.Controls("dc_SystemType1").AddItem ("SPLIT SYSTEM")
    AuditForm.Controls("dc_SystemType1").AddItem ("WINDOW AC")
    
'    loctextboxlen = 50
'    cboWidth = 70
    
    ' FUEL SOURCE
    Call addlabel(lblFuelSource, "dc_lblFuelSource1", "Fuel Source*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(FuelSource, "dc_FuelSource1", toTop1, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_FuelSource1").AddItem ("ELECTRIC")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop2, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemSize, "dc_SystemSize1", toTop2, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "dc_lblSizeUnit1", "System Size Unit*", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(SizeUnit, "dc_SizeUnit1", toTop2, toLeft3, cboWidth, cboHeight)
    AuditForm.Controls("dc_SizeUnit1").AddItem ("BTU")
    AuditForm.Controls("dc_SizeUnit1").AddItem ("MBTU")
    AuditForm.Controls("dc_SizeUnit1").AddItem ("MMBTU")
    AuditForm.Controls("dc_SizeUnit1").AddItem ("TON")

    'SYSTEM Age
    Call addlabel(lblSystemAge, "dc_lblSystemAge1", "System Age*", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemAge, "dc_SystemAge1", toTop3, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM Efficiency Rating
    Call addlabel(lblEffRating, "dc_lblEffRating1", "Efficiency Rating*", toTop4, toLeft - 10, lblWidth, lblHeight)
    Call addtextbox(EffRating, "dc_EffRating1", toTop4, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM Efficiency Rating Type
    Call addlabel(lblEffRatingType, "dc_lblEffRatingType1", "Rating Type*", toTop4, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(EffRatingType, "dc_EffRatingType1", toTop4, toLeft3, cboWidth, cboHeight)
    AuditForm.Controls("dc_EffRatingType1").AddItem ("EER")
    AuditForm.Controls("dc_EffRatingType1").AddItem ("SEER")
    AuditForm.Controls("dc_EffRatingType1").AddItem ("COP")
    
    'TOTAL PERCENTAGE OF SPACE COOLED
    Call addlabel(lblPercentageCooled, "dc_lblPercentageCooled1", "% of space cooled*", toTop5, toLeft - 13, lblWidth, lblHeight)
    Call addtextbox(PercentageCooled, "dc_PercentageCooled1", toTop5, toLeft1, txtWidth, txtHeight)
    
    'FREQUENCY OF SYSTEM USE

    Call addlabel(lblFrequencyUse, "dc_lblFrequencyUse1", "Frequency of use*", toTop5, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(FrequencyUse, "dc_FrequencyUse1", toTop5, toLeft3, cboWidth, cboHeight)
    AuditForm.Controls("dc_FrequencyUse1").AddItem ("0%")
    AuditForm.Controls("dc_FrequencyUse1").AddItem ("10-30%")
    AuditForm.Controls("dc_FrequencyUse1").AddItem ("31-70%")
    AuditForm.Controls("dc_FrequencyUse1").AddItem ("71-100%")
    
    'TOTAL UNITS USED
    Call addlabel(lblTotalUnits, "dc_lblTotalUnits1", "Total units used", toTop6, toLeft, lblWidth, lblHeight)
    Call addtextbox(TotalUnits, "dc_TotalUnits1", toTop6, toLeft1, txtWidth, txtHeight)
        
    'QUANTITY
    Call addlabel(lblQuantity, "dc_lblQuantity1", "Quantity", toTop6, toLeft2, lblWidth, lblHeight)
    Call addtextbox(Quantity, "dc_Quantity1", toTop6, toLeft3, txtWidth, txtHeight)
        
End Sub

Private Sub showheatingoptions()
    
    ' HEATING TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Heating Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_SystemType1").AddItem ("GAS FURNACE")
    AuditForm.Controls("dc_SystemType1").AddItem ("HEAT PUMP-AIR SOURCE")
    AuditForm.Controls("dc_SystemType1").AddItem ("HEAT PUMP-GROUND SOURCE")
    AuditForm.Controls("dc_SystemType1").AddItem ("HEAT PUMP-DUAL FUEL")
    AuditForm.Controls("dc_SystemType1").AddItem ("RESISTANCE ELECTRIC HEAT")
    AuditForm.Controls("dc_SystemType1").AddItem ("HOT WATER BOILER")
    AuditForm.Controls("dc_SystemType1").AddItem ("FORCED AIR")
    AuditForm.Controls("dc_SystemType1").AddItem ("STEAM")
    AuditForm.Controls("dc_SystemType1").AddItem ("WOOD/COAL STOVE")
    
    ' FUEL SOURCE
    Call addlabel(lblFuelSource, "dc_lblFuelSource1", "Fuel Source*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(FuelSource, "dc_FuelSource1", toTop1, toLeft1, cboWidth * 2, cboHeight)
    AuditForm.Controls("dc_FuelSource1").AddItem ("ELECTRIC")
    AuditForm.Controls("dc_FuelSource1").AddItem ("GAS")
    AuditForm.Controls("dc_FuelSource1").AddItem ("PROPANE")
    AuditForm.Controls("dc_FuelSource1").AddItem ("CENTRAL STEAM")
    AuditForm.Controls("dc_FuelSource1").AddItem ("COAL")
    AuditForm.Controls("dc_FuelSource1").AddItem ("SOLAR")
    AuditForm.Controls("dc_FuelSource1").AddItem ("WOOD")
    AuditForm.Controls("dc_FuelSource1").AddItem ("OIL")
    AuditForm.Controls("dc_FuelSource1").AddItem ("OTHER")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop2, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemSize, "dc_SystemSize1", toTop2, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "dc_lblSizeUnit1", "System Size Unit*", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(SizeUnit, "dc_SizeUnit1", toTop2, toLeft3, cboWidth, cboHeight)
    AuditForm.Controls("dc_SizeUnit1").AddItem ("MBTU")
    AuditForm.Controls("dc_SizeUnit1").AddItem ("MMBTU")
    AuditForm.Controls("dc_SizeUnit1").AddItem ("TON")

    'SYSTEM Age
    Call addlabel(lblSystemAge, "dc_lblSystemAge1", "System Age", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemAge, "dc_SystemAge1", toTop3, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM Efficiency Rating
    Call addlabel(lblEffRating, "dc_lblEffRating1", "Efficiency Rating", toTop4, toLeft - 10, lblWidth, lblHeight)
    Call addtextbox(EffRating, "dc_EffRating1", toTop4, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM Efficiency Rating Type
    Call addlabel(lblEffRatingType, "dc_lblEffRatingType1", "Rating Type*", toTop4, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(EffRatingType, "dc_EffRatingType1", toTop4, toLeft3, cboWidth, cboHeight)
    AuditForm.Controls("dc_EffRatingType1").AddItem ("AFUE")
    AuditForm.Controls("dc_EffRatingType1").AddItem ("HSPF")
    AuditForm.Controls("dc_EffRatingType1").AddItem ("COP")
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
    
    stv = AuditForm.Controls("dc_SystemType1").Value
    If stv = "GAS FURNACE" Or stv = "HEAT PUMP-AIR SOURCE" Or stv = "HEAT PUMP-GROUND SOURCE" _
        Or stv = "HEAT PUMP-DUAL FUEL" Or stv = "RESISTANCE ELECTRIC HEAT" Or stv = "HOT WATER BOILER" _
        Or stv = "FORCED AIR" Or stv = "STEAM" Or stv = "WOOD/COAL STOVE" Then
    Else
        errorstring ("System Type")
    End If
    
    fs = AuditForm.Controls("dc_FuelSource1").Value
    If fs = "ELECTRIC" Or fs = "GAS" Or fs = "PROPANE" Or fs = "CENTRAL STEAM" Or fs = "COAL" Or fs = "SOLAR" _
        Or fs = "WOOD" Or fs = "OIL" Or fs = "OTHER" Then
    Else
        errorstring ("Fuel Source")
    End If
    
    If Not IsNumeric(AuditForm.Controls("dc_SystemSize1").Value) Then
        errorstring ("System Size")
    End If

    su = AuditForm.Controls("dc_SizeUnit1").Value
    If su = "MBTU" Or su = "MMBTU" Or su = "TON" Then
    Else
        errorstring ("Size Unit")
    End If
    
    If IsNumeric(AuditForm.Controls("dc_SystemAge1").Value) Or AuditForm.Controls("dc_SystemAge1").Value = "" Then
    Else
        errorstring ("System Age")
    End If
    
    If Not IsNumeric(AuditForm.Controls("dc_EffRating1").Value) Then
        errorstring ("Efficiency Rating")
    End If
    
    et = AuditForm.Controls("dc_EffRatingType1").Value
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
            AuditForm.Controls("dc_SystemType1").Value = Worksheets("Audit").Cells(ir + 2, 7)
            AuditForm.Controls("dc_FuelSource1").Value = Worksheets("Audit").Cells(ir + 2, 8)
            AuditForm.Controls("dc_SystemSize1").Value = Worksheets("Audit").Cells(ir + 2, 11)
            AuditForm.Controls("dc_SizeUnit1").Value = Worksheets("Audit").Cells(ir + 2, 12)
            AuditForm.Controls("dc_SystemAge1").Value = Worksheets("Audit").Cells(ir + 2, 13)
            AuditForm.Controls("dc_EffRating1").Value = Worksheets("Audit").Cells(ir + 2, 15)
            AuditForm.Controls("dc_EffRatingType1").Value = Worksheets("Audit").Cells(ir + 2, 16)
        Case "COOLING"
            AuditForm.Controls("dc_SystemType1").Value = Worksheets("Audit").Cells(ir + 2, 7)
            AuditForm.Controls("dc_FuelSource1").Value = Worksheets("Audit").Cells(ir + 2, 8)
            AuditForm.Controls("dc_SystemSize1").Value = Worksheets("Audit").Cells(ir + 2, 11)
            AuditForm.Controls("dc_SizeUnit1").Value = Worksheets("Audit").Cells(ir + 2, 12)
            AuditForm.Controls("dc_SystemAge1").Value = Worksheets("Audit").Cells(ir + 2, 13)
            AuditForm.Controls("dc_EffRating1").Value = Worksheets("Audit").Cells(ir + 2, 15)
            AuditForm.Controls("dc_EffRatingType1").Value = Worksheets("Audit").Cells(ir + 2, 16)
            AuditForm.Controls("dc_PercentageCooled1").Value = Worksheets("Audit").Cells(ir + 2, 17)
            AuditForm.Controls("dc_FrequencyUse1").Value = Worksheets("Audit").Cells(ir + 2, 18)
            AuditForm.Controls("dc_TotalUnits1").Value = Worksheets("Audit").Cells(ir + 2, 19)
            AuditForm.Controls("dc_Quantity1").Value = Worksheets("Audit").Cells(ir + 2, 14)
        Case "HVAC DISTRIBUTION"
            AuditForm.Controls("dc_SystemType1").Value = Worksheets("Audit").Cells(ir + 1, 7)
            AuditForm.Controls("dc_SystemSize1").Value = Worksheets("Audit").Cells(ir + 1, 11)
            AuditForm.Controls("dc_InsIndicator1").Value = Worksheets("Audit").Cells(ir + 1, 20)
            AuditForm.Controls("dc_InsType1").Value = Worksheets("Audit").Cells(ir + 1, 21)
            AuditForm.Controls("dc_SystemLocation1").Value = Worksheets("Audit").Cells(ir + 1, 23)
            AuditForm.Controls("dc_SystemLength1").Value = Worksheets("Audit").Cells(ir + 1, 24)
            AuditForm.Controls("dc_FlexCondition1").Value = Worksheets("Audit").Cells(ir + 1, 30)
        
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
        Case "HVAC DISTRIBUTION"
            Call savehvacdistribution
            
    End Select
    
End Sub
Private Sub savehvacdistribution()
    If iHVAC < 3 Then
        iHVAC = iHVAC + 1
        lastrow = Worksheets("Audit").Range("E" & Rows.Count).End(xlUp).Row
        If strCurrentSystemName = "" Then
            strCurrentSystemName = "HVAC DISTRIBUTION-" + CStr(iHVAC)
        End If
        Worksheets("Audit").Cells(lastrow + 1, 1) = strCurrentSystemName
        Worksheets("Audit").Cells(lastrow + 1, 5) = "HVAC DISTRIBUTION"
        Worksheets("Audit").Cells(lastrow + 1, 7) = AuditForm.Controls("dc_SystemType1").Value
        Worksheets("Audit").Cells(lastrow + 1, 11) = AuditForm.Controls("dc_SystemSize1").Value
        Worksheets("Audit").Cells(lastrow + 1, 20) = AuditForm.Controls("dc_InsIndicator1").Value
        Worksheets("Audit").Cells(lastrow + 1, 21) = AuditForm.Controls("dc_InsType1").Value
        Worksheets("Audit").Cells(lastrow + 1, 23) = AuditForm.Controls("dc_SystemLocation1").Value
        Worksheets("Audit").Cells(lastrow + 1, 24) = AuditForm.Controls("dc_SystemLength1").Value
        Worksheets("Audit").Cells(lastrow + 1, 30) = AuditForm.Controls("dc_FlexCondition1").Value
        lstSelectedSystems.AddItem (strCurrentSystemName)
        strCurrentSystemName = ""
    Else
        MsgBox ("You can only enter at most 3 HVAC Distribution systems!")
    End If
    
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
        Worksheets("Audit").Cells(lastrow + 1, 7) = AuditForm.Controls("dc_SystemType1").Value
        Worksheets("Audit").Cells(lastrow + 1, 8) = AuditForm.Controls("dc_FuelSource1").Value
        Worksheets("Audit").Cells(lastrow + 1, 11) = AuditForm.Controls("dc_SystemSize1").Value
        Worksheets("Audit").Cells(lastrow + 1, 12) = AuditForm.Controls("dc_SizeUnit1").Value
        Worksheets("Audit").Cells(lastrow + 1, 13) = AuditForm.Controls("dc_SystemAge1").Value
        Worksheets("Audit").Cells(lastrow + 1, 15) = AuditForm.Controls("dc_EffRating1").Value
        Worksheets("Audit").Cells(lastrow + 1, 16) = AuditForm.Controls("dc_EffRatingType1").Value
        lstSelectedSystems.AddItem (strCurrentSystemName)
        strCurrentSystemName = ""
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
        Worksheets("Audit").Cells(lastrow + 1, 7) = AuditForm.Controls("dc_SystemType1").Value
        Worksheets("Audit").Cells(lastrow + 1, 8) = AuditForm.Controls("dc_FuelSource1").Value
        Worksheets("Audit").Cells(lastrow + 1, 11) = AuditForm.Controls("dc_SystemSize1").Value
        Worksheets("Audit").Cells(lastrow + 1, 12) = AuditForm.Controls("dc_SizeUnit1").Value
        Worksheets("Audit").Cells(lastrow + 1, 13) = AuditForm.Controls("dc_SystemAge1").Value
        Worksheets("Audit").Cells(lastrow + 1, 15) = AuditForm.Controls("dc_EffRating1").Value
        Worksheets("Audit").Cells(lastrow + 1, 16) = AuditForm.Controls("dc_EffRatingType1").Value
        
        Worksheets("Audit").Cells(lastrow + 1, 17) = AuditForm.Controls("dc_PercentageCooled1").Value
        Worksheets("Audit").Cells(lastrow + 1, 18) = AuditForm.Controls("dc_FrequencyUse1").Value
        Worksheets("Audit").Cells(lastrow + 1, 19) = AuditForm.Controls("dc_TotalUnits1").Value
        Worksheets("Audit").Cells(lastrow + 1, 14) = AuditForm.Controls("dc_Quantity1").Value
        lstSelectedSystems.AddItem (strCurrentSystemName)
        strCurrentSystemName = ""
    Else
        MsgBox ("You can only enter at most 3 COOLING systems!")
    End If
    
End Sub

Private Sub cmdRemove_Click()
    ir = lstSelectedSystems.ListIndex
    Select Case Worksheets("Audit").Cells(ir + 2, 5).Value
        Case "HEATING"
            iHeating = iHeating - 1
        Case "COOLING"
            iCooling = iCooling - 1
        Case "HVAC DISTRIBUTION"
            iHVAC = iHVAC - 1
        Case "WATER HEATER"
            iWH = iWH - 1
        Case "THERMOSTAT"
            iThermostat = iThermostat - 1
        Case "WINDOW"
            iWindow = iWindow - 1
        Case "DOOR"
            iDoor = iDoor - 1
        Case "LIGHTING"
            iLighting = iLighting - 1
        Case "WALL"
            iWall = iWall - 1
        Case "ATTIC"
            iAttic = iAttic - 1
        Case "BASEMENT"
            iBasement = iBasement - 1
        Case "BASEMENT WALL"
            iBW = iBW - 1
        Case "REFRIGERATOR"
            iRefrigerator = iRefrigerator - 1
        Case "FREEZER"
            iFreezer = iFreezer - 1
        Case "APPLIANCE"
            iAppliance = iAppliance - 1
        Case Else
    End Select
    lstSelectedSystems.RemoveItem (ir)
    Worksheets("Audit").Rows(ir + 2).Delete

End Sub

Private Sub cmdRemoveAll_Click()
    lastrow = Worksheets("Audit").Range("E" & Rows.Count).End(xlUp).Row
    cboSystem.Text = ""
    strCurrentSystemName = ""
    Worksheets("Audit").Range("A2:AZ" & lastrow).Clear
    lstSelectedSystems.Clear
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
    
    vertInterval = 25
    
    toTop = 40
    toTop1 = toTop + vertInterval
    toTop2 = toTop + 2 * vertInterval
    toTop3 = toTop + 3 * vertInterval
    toTop4 = toTop + 4 * vertInterval
    toTop5 = toTop + 5 * vertInterval
    toTop6 = toTop + 6 * vertInterval
    toTop7 = toTop + 7 * vertInterval
    toTop8 = toTop + 8 * vertInterval
    toTop9 = toTop + 9 * vertInterval
    
    toLeft = 20
    toLeft1 = 85
    toLeft2 = 150
    toLeft3 = 220
    
    cboHeight = 15
    cboWidth = 70
    txtHeight = 15
    txtWidth = 50
    lblHeight = 20
    lblWidth = 80

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
