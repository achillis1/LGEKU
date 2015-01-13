VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSystem 
   Caption         =   "Systems ROSA"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12855
   OleObjectBlob   =   "frmSystem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSystem"
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
Private iAppliance As Integer

Private thisWB As Workbook
Private prompt As String

Private lblSystemApplicable As Object
Private SystemApplicable As Object
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
Private lblTankRValue As Object
Private TankRValue As Object
Private lblPercentageLoad As Object
Private PercentageLoad As Object
Private lblTemperatureSetting As Object
Private lblTemperatureSetting1 As Object
Private TemperatureSetting As Object
Private lblEnergyFactor As Object
Private EnergyFactor As Object
Private lblAverageCoolingTemperature As Object
Private AverageCoolingTemperature As Object
Private lblAverageCoolingTemperature1 As Object
Private lblAverageHeatingTemperature As Object
Private AverageHeatingTemperature As Object
Private lblAverageHeatingTemperature2 As Object
Private lblDaytimeSetback As Object
Private DaytimeSetback As Object
Private lblEveningSetback As Object
Private EveningSetback As Object
Private lblNightSetback As Object
Private NightSetback As Object
Private lblHeatingDayTemperature As Object
Private HeatingDayTemperature As Object
Private lblHeatingDayTemperature1 As Object
Private lblHeatingEveningTemperature As Object
Private HeatingEveningTemperature As Object
Private lblHeatingEveningTemperature1 As Object
Private lblHeatingNightTemperature As Object
Private HeatingNightTemperature As Object
Private lblHeatingNightTemperature1 As Object
Private lblCoolingDayTemperature As Object
Private CoolingDayTemperature As Object
Private lblCoolingDayTemperature1 As Object
Private lblCoolingEveningTemperature As Object
Private CoolingEveningTemperature As Object
Private lblCoolingEveningTemperature1 As Object
Private lblCoolingNightTemperature As Object
Private CoolingNightTemperature As Object
Private lblCoolingNightTemperature1 As Object
Private lblACCtrlPresent As Object
Private ACCtrlPresent As Object
Private lblWindowDoorCondition As Object
Private WindowDoorCondition As Object
Private lblSurfaceArea As Object
Private SurfaceArea As Object
Private lblWindowUVCoated As Object
Private WindowUVCoated As Object
Private lblNumberOfGlazing As Object
Private NumberOfGlazing As Object
Private lblTotalWeeklyHours As Object
Private TotalWeeklyHours As Object
Private lblBulbWattage As Object
Private BulbWattage As Object
Private lblSystemHeight As Object
Private SystemHeight As Object
Private lblSystemHeight1 As Object
Private lblVentIndicator As Object
Private VentIndicator As Object
Private lblAccessType As Object
Private AccessType As Object
Private lblSystemDepth As Object
Private SystemDepth As Object
Private lblSystemDepth1 As Object
Private lblBasementAC As Object
Private BasementAC As Object
Private lblRJInsRecommended As Object
Private RJInsRecommended As Object
Private lblPerimeterFootage As Object
Private PerimeterFootage As Object
Private lblDefrostType As Object
Private DefrostType As Object
Private lblSystemMake As Object
Private SystemMake As Object
Private lblSystemMeteredUsage As Object
Private SystemMeteredUsage As Object

Private strCurrentSystemName As String

Private cboHeight As Integer
Private cboWidth As Integer
Private txtHeight As Integer
Private txtWidth As Integer
Private lblHeight As Integer
Private lblWidth As Integer

Private toTop As Integer
Private toTop0 As Integer
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

Private vertInterval As Integer

Private applianceStartCol As Integer
Private applianceNum As Integer
Private applianceLimit As Integer

Private atticStartCol As Integer
Private atticNum As Integer
Private atticLimit As Integer

Private basementStartCol As Integer
Private basementNum As Integer
Private basementLimit As Integer

Private basementwallStartCol As Integer
Private basementwallNum As Integer
Private basementwallLimit As Integer

Private coolingStartCol As Integer
Private coolingNum As Integer
Private coolingLimit As Integer

Private doorStartCol As Integer
Private doorNum As Integer
Private doorLimit As Integer

Private freezerStartCol As Integer
Private freezerNum As Integer
Private freezerLimit As Integer

Private heatingStartCol As Integer
Private heatingNum As Integer
Private heatingLimit As Integer

Private hvacdistStartCol As Integer
Private hvacdistNum As Integer
Private hvacdistLimit As Integer

Private lightingStartCol As Integer
Private lightingNum As Integer
Private lightingLimit As Integer

Private refrigStartCol As Integer
Private refrigNum As Integer
Private refrigLimit As Integer

Private thermostatStartCol As Integer
Private thermostatNum As Integer
Private thermostatLimit As Integer

Private wallStartCol As Integer
Private wallNum As Integer
Private wallLimit As Integer

Private waterheaterStartCol As Integer
Private waterheaterNum As Integer
Private waterheaterLimit As Integer

Private sysnum As Variant
Private syslimit As Variant

Private bSystemLoad As Boolean
Private oldSystemName As String

Private Sub cboSystem_Change()
    
    For Each ctrl In frmSystem.Controls
        If left(ctrl.Name, 3) = "dc_" Then
            frmSystem.Controls.Remove (ctrl.Name)
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
    frmSystem.Hide
    frmMain.Show vbModeless
'    Application.Visible = True
'    ThisWorkbook.Saved = True
    'ThisWorkbook.Close SaveChanges:=True
    
   ' Application.Quit
End Sub

Private Sub addcomboBox(ByVal cboBox As Object, ByVal cboname As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer)

    For Each ctrl In frmSystem.Controls
        If ctrl.Name = cboname Then
            frmSystem.Controls.Remove (cboname)
        End If
    Next

    Set cboBox = frmSystem.Controls.Add("Forms.ComboBox.1", cboname, True)
    With cboBox
        .top = top
        .left = left
        .width = width
        .height = height
    End With

End Sub

Private Sub addlabel(ByVal label As Object, ByVal labelname As String, ByVal caption As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer)

    For Each ctrl In frmSystem.Controls
        If ctrl.Name = labelname Then
            frmSystem.Controls.Remove (labelname)
        End If
    Next
    
    Set label = frmSystem.Controls.Add("Forms.Label.1", labelname, True)
    With label
        .caption = caption
        .top = top
        .left = left
        .width = width
        .height = height
    End With
End Sub


Private Sub addtextbox(ByVal textbox As Object, ByVal textboxname As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer)

    For Each ctrl In frmSystem.Controls
        If ctrl.Name = textboxname Then
            frmSystem.Controls.Remove (textboxname)
        End If
    Next
    Set textbox = frmSystem.Controls.Add("Forms.TextBox.1", textboxname, True)
    With textbox
        .top = top
        .left = left
        .width = width
        .height = height
    End With
End Sub


Private Sub showwhoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'WATER HEATER TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Water Heater Type*", toTop, toLeft - 15, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 3, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("CONVENTIONAL STORAGE")
    frmSystem.Controls("dc_SystemType1").AddItem ("DEMAND")
    frmSystem.Controls("dc_SystemType1").AddItem ("TANKLESS\INSTANTANEOUS")
    frmSystem.Controls("dc_SystemType1").AddItem ("SOLAR/TANK")
    frmSystem.Controls("dc_SystemType1").AddItem ("GEOTHERMAL DESUPERHEATER/TANK")
    
    ' FUEL SOURCE
    Call addlabel(lblFuelSource, "dc_lblFuelSource1", "Fuel Source*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(FuelSource, "dc_FuelSource1", toTop1, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_FuelSource1").AddItem ("ELECTRIC")
    frmSystem.Controls("dc_FuelSource1").AddItem ("GAS")
    frmSystem.Controls("dc_FuelSource1").AddItem ("PROPANE")
    frmSystem.Controls("dc_FuelSource1").AddItem ("SOLAR")
    frmSystem.Controls("dc_FuelSource1").AddItem ("WOOD")
    frmSystem.Controls("dc_FuelSource1").AddItem ("OIL")
    frmSystem.Controls("dc_FuelSource1").AddItem ("OTHER")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop2, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemSize, "dc_SystemSize1", toTop2, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "dc_lblSizeUnit1", "System Size Unit*", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(SizeUnit, "dc_SizeUnit1", toTop2, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_SizeUnit1").AddItem ("GALLONS")
    
    'SYSTEM Age
    Call addlabel(lblSystemAge, "dc_lblSystemAge1", "System Age*", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemAge, "dc_SystemAge1", toTop3, toLeft1, txtWidth, txtHeight)
    
    'INSULATION EXIST INDICATOR
    Call addlabel(lblInsIndicator, "dc_lblInsIndicator1", "Insulation exist indicator*", toTop3, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(InsIndicator, "dc_InsIndicator1", toTop3, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_InsIndicator1").AddItem ("Y")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("N")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("NOT NEEDED")
    
    'INSULATION TYPE
    Call addlabel(lblInsType, "dc_lblInsType1", "Insulation Type*", toTop4, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsType, "dc_InsType1", toTop4, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BATTS")
    frmSystem.Controls("dc_InsType1").AddItem ("MINERAL/ROCK WOOL")
    frmSystem.Controls("dc_InsType1").AddItem ("NONE")
    frmSystem.Controls("dc_InsType1").AddItem ("OTHER")
    

    'TANK R-VALUE
    Call addlabel(lblTankRValue, "dc_lblTankRValue1", "Tank R-Value", toTop5, toLeft, lblWidth, lblHeight)
    Call addtextbox(TankRValue, "dc_TankRValue1", toTop5, toLeft1, txtWidth, txtHeight)
    
    'PERCENTAGE OF LOAD
    Call addlabel(lblPercentageLoad, "dc_lblPercentageLoad1", "% of Load*", toTop5, toLeft2 + 20, lblWidth, lblHeight)
    Call addtextbox(PercentageLoad, "dc_PercentageLoad1", toTop5, toLeft3, txtWidth, txtHeight)
    
    'CURRENT TEMPERATURE SETTING
    Call addlabel(lblTemperatureSetting, "dc_lblTemperatureSetting1", "Current Temperature Setting*", toTop6, toLeft - 13, lblWidth, lblHeight)
    Call addtextbox(TemperatureSetting, "dc_TemperatureSetting1", toTop6, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblTemperatureSetting1, "dc_lblTemperatureSetting2", "F", toTop6, toLeft1 + txtWidth + 5, lblWidth, lblHeight)
    
    'ENERGY FACTOR
    Call addlabel(lblEnergyFactor, "dc_lblEnergyFactor1", "Energy Factor", toTop7, toLeft, lblWidth, lblHeight)
    Call addtextbox(EnergyFactor, "dc_EnergyFactor1", toTop7, toLeft1, txtWidth, txtHeight)
End Sub

Private Sub showhvacoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    ' HVAC DISTRIBUTION TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "HVAC Distribution Type*", toTop, toLeft - 10, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("DUCT ROUND")
    frmSystem.Controls("dc_SystemType1").AddItem ("DUCT RECTANGULAR")
    frmSystem.Controls("dc_SystemType1").AddItem ("IRON PIPE")
    frmSystem.Controls("dc_SystemType1").AddItem ("COPPER")
    frmSystem.Controls("dc_SystemType1").AddItem ("ELBOWS")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemSize, "dc_SystemSize1", toTop1, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemSize1").AddItem ("SMALL")
    frmSystem.Controls("dc_SystemSize1").AddItem ("MEDIUM")
    frmSystem.Controls("dc_SystemSize1").AddItem ("LARGE")
    frmSystem.Controls("dc_SystemSize1").AddItem ("EXTRA LARGE")
    
    'INSULATION EXIST INDICATOR
    Call addlabel(lblInsIndicator, "dc_lblInsIndicator1", "Insulation exist indicator*", toTop2, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsIndicator, "dc_InsIndicator1", toTop2, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_InsIndicator1").AddItem ("Y")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("N")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("NOT NEEDED")
    
    'INSULATION TYPE
    Call addlabel(lblInsType, "dc_lblInsType1", "Insulation Type*", toTop3, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsType, "dc_InsType1", toTop3, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_InsType1").AddItem ("CELLULOSE")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BATTS")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BLOWN")
    frmSystem.Controls("dc_InsType1").AddItem ("LOOSE FIBERGLASS")
    frmSystem.Controls("dc_InsType1").AddItem ("MINERAL/ROCK WOOL")
    frmSystem.Controls("dc_InsType1").AddItem ("UREA FORMALDAHYDE")
    frmSystem.Controls("dc_InsType1").AddItem (".5 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("2 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("NONE")
    frmSystem.Controls("dc_InsType1").AddItem ("OTHER")
    

    'SYSTEM LOCATION
    Call addlabel(lblSystemLocation, "dc_lblSystemLocation1", "System Location*", toTop4, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemLocation, "dc_SystemLocation1", toTop4, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemLocation1").AddItem ("ATTIC")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("BASEMENT")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("CRAWL")
    
    'LENGTH
    Call addlabel(lblSystemLength, "dc_lblSystemLength1", "Length", toTop5, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemLength, "dc_SystemLength1", toTop5, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblSystemLength1, "dc_lblSystemLength2", "ft", toTop5, toLeft1 + txtWidth + 5, lblWidth, lblHeight)
        
    'CONDITION OF FLEX
    Call addlabel(lblFlexCondition, "dc_lblFlexCondition1", "Flex Condition", toTop6, toLeft, lblWidth, lblHeight)
    Call addcomboBox(FlexCondition, "dc_FlexCondition1", toTop6, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_FlexCondition1").AddItem ("COLLAPSED")
    frmSystem.Controls("dc_FlexCondition1").AddItem ("DAMAGED")
    frmSystem.Controls("dc_FlexCondition1").AddItem ("FUNCTIONAL")
    frmSystem.Controls("dc_FlexCondition1").AddItem ("NON-FUNCTIONAL COLLAPSED")
    frmSystem.Controls("dc_FlexCondition1").AddItem ("NON-FUNCTIONAL DAMAGED")
    
End Sub


Private Sub showthermostatoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'THERMOSTAT TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Thermostat Type*", toTop, toLeft - 10, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("DIGITAL")
    frmSystem.Controls("dc_SystemType1").AddItem ("DIGITAL PROGRAMMABLE")
    frmSystem.Controls("dc_SystemType1").AddItem ("STANDARD")
    
    'PERCENTAGE OF LOAD
    Call addlabel(lblPercentageLoad, "dc_lblPercentageLoad1", "% of Load", toTop, toLeft2 + 100, lblWidth, lblHeight)
    Call addtextbox(PercentageLoad, "dc_PercentageLoad1", toTop, toLeft3 + 80, txtWidth, txtHeight)
    

    'AVERAGE COOLING TEMPERATURE SETTING
    Call addlabel(lblAverageCoolingTemperature, "dc_lblAverageCoolingTemperature1", "Avg. Cooling Temp*", toTop1, toLeft - 13, lblWidth, lblHeight)
    Call addtextbox(AverageCoolingTemperature, "dc_AverageCoolingTemperature1", toTop1, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblAverageCoolingTemperature1, "dc_lblAverageCoolingTemperature2", "F", toTop1, toLeft1 + txtWidth + 5, lblWidth - 45, lblHeight)
    
    'AVERAGE HEATING TEMPERATURE SETTING
    Call addlabel(lblAverageHeatingTemperature, "dc_lblAverageHeatingTemperature1", "Avg. Heating Temp*", toTop1, toLeft2 + 20, lblWidth, lblHeight + 10)
    Call addtextbox(AverageHeatingTemperature, "dc_AverageHeatingTemperature1", toTop1, toLeft3 + 25, txtWidth, txtHeight)
    Call addlabel(lblAverageHeatingTemperature2, "dc_lblAverageHeatingTemperature3", "F", toTop1, toLeft3 + txtWidth + 30, lblWidth, lblHeight)

    'DAYTIME SETBACK
    Call addlabel(lblDaytimeSetback, "dc_lblDaytimeSetback1", "Daytime set back?", toTop2, toLeft - 10, lblWidth, lblHeight)
    Call addcomboBox(DaytimeSetback, "dc_DaytimeSetback1", toTop2, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_DaytimeSetback1").AddItem ("Y")
    frmSystem.Controls("dc_DaytimeSetback1").AddItem ("N")
    
    'EVENING SETBACK
    Call addlabel(lblEveningSetback, "dc_lblEveningSetback1", "Evening set back?", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(EveningSetback, "dc_EveningSetback1", toTop2, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_EveningSetback1").AddItem ("Y")
    frmSystem.Controls("dc_EveningSetback1").AddItem ("N")
    
    'NIGHTTIME SETBACK
    Call addlabel(lblNightSetback, "dc_lblNightSetback1", "Night set back?", toTop3, toLeft, lblWidth, lblHeight)
    Call addcomboBox(NightSetback, "dc_NightSetback1", toTop3, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_NightSetback1").AddItem ("Y")
    frmSystem.Controls("dc_NightSetback1").AddItem ("N")
    
    'HEATING DAY TEMP SETTING
    Call addlabel(lblHeatingDayTemperature, "dc_lblHeatingDayTemperature1", "Heating Day Temp.", toTop3, toLeft2, lblWidth, lblHeight)
    Call addtextbox(HeatingDayTemperature, "dc_HeatingDayTemperature1", toTop3, toLeft3, txtWidth, txtHeight)
    Call addlabel(lblHeatingDayTemperature1, "dc_lblHeatingDayTemperature2", "F", toTop3, toLeft3 + txtWidth + 5, lblWidth, lblHeight)
    
    'HEATING EVENING TEMP SETTING
    Call addlabel(lblHeatingEveningTemperature, "dc_lblHeatingEveningTemperature1", "Heating Evening Temp.", toTop4, toLeft, lblWidth, lblHeight)
    Call addtextbox(HeatingEveningTemperature, "dc_HeatingEveningTemperature1", toTop4, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblHeatingEveningTemperature1, "dc_lblHeatingEveningTemperature2", "F", toTop4, toLeft1 + txtWidth + 5, lblWidth, lblHeight)
    
    'HEATING NIGHT TEMP SETTING
    Call addlabel(lblHeatingNightTemperature, "dc_lblHeatingNightTemperature1", "Heating Night Temp.", toTop4, toLeft2, lblWidth, lblHeight)
    Call addtextbox(HeatingNightTemperature, "dc_HeatingNightTemperature1", toTop4, toLeft3, txtWidth, txtHeight)
    Call addlabel(lblHeatingNightTemperature1, "dc_lblHeatingNightTemperature2", "F", toTop4, toLeft3 + txtWidth + 5, lblWidth, lblHeight)
    
    'COOLING DAY TEMP SETTING
    Call addlabel(lblCoolingDayTemperature, "dc_lblCoolingDayTemperature1", "Cooling Day Temp.", toTop5, toLeft - 13, lblWidth, lblHeight)
    Call addtextbox(CoolingDayTemperature, "dc_CoolingDayTemperature1", toTop5, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblCoolingDayTemperature1, "dc_lblCoolingDayTemperature2", "F", toTop5, toLeft1 + txtWidth + 5, lblWidth, lblHeight)
    
    'COOLING EVENING TEMP SETTING
    Call addlabel(lblCoolingEveningTemperature, "dc_lblCoolingEveningTemperature1", "Cooling Evening Temp.", toTop5, toLeft2 + 10, lblWidth, lblHeight)
    Call addtextbox(CoolingEveningTemperature, "dc_CoolingEveningTemperature1", toTop5, toLeft3, txtWidth, txtHeight)
    Call addlabel(lblCoolingEveningTemperature1, "dc_lblCoolingEveningTemperature2", "F", toTop5, toLeft3 + txtWidth + 5, lblWidth, lblHeight)
    
    'COOLING NIGHT TEMP SETTING
    Call addlabel(lblCoolingNightTemperature, "dc_lblCoolingNightTemperature1", "Cooling Night Temp.", toTop6, toLeft - 15, lblWidth, lblHeight)
    Call addtextbox(CoolingNightTemperature, "dc_CoolingNightTemperature1", toTop6, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblCoolingNightTemperature1, "dc_lblCoolingNightTemperature2", "F", toTop6, toLeft1 + txtWidth + 5, lblWidth, lblHeight)
    

    'AC LOAD CONTROL PRESENT
    Call addlabel(lblACCtrlPresent, "dc_lblACCtrlPresent1", "AC Load Control Present?", toTop7, toLeft - 15, lblWidth, lblHeight)
    Call addcomboBox(ACCtrlPresent, "dc_ACCtrlPresent1", toTop7, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_ACCtrlPresent1").AddItem ("Y")
    frmSystem.Controls("dc_ACCtrlPresent1").AddItem ("N")
End Sub

Private Sub showwindowoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'WINDOW TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Window Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("SINGLE PANE")
    frmSystem.Controls("dc_SystemType1").AddItem ("SINGLE PANE W/STORM")
    frmSystem.Controls("dc_SystemType1").AddItem ("WINDOW")
    frmSystem.Controls("dc_SystemType1").AddItem ("DOUBLE PANE")
    frmSystem.Controls("dc_SystemType1").AddItem ("TRIPLE PANE")
    frmSystem.Controls("dc_SystemType1").AddItem ("DH")
    frmSystem.Controls("dc_SystemType1").AddItem ("CASEMENT")
    frmSystem.Controls("dc_SystemType1").AddItem ("FIXED")
    
    'QUANTITY
    Call addlabel(lblQuantity, "dc_lblQuantity1", "Quantity", toTop1, toLeft, lblWidth, lblHeight)
    Call addtextbox(Quantity, "dc_Quantity1", toTop1, toLeft1, txtWidth, txtHeight)
    
    'WINDOW CONDITION
    Call addlabel(lblWindowDoorCondition, "dc_lblWindowDoorCondition1", "Window Condition*", toTop1, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(WindowDoorCondition, "dc_WindowDoorCondition1", toTop1, toLeft3, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_WindowDoorCondition1").AddItem ("NO APPARENT DAMAGE")
    frmSystem.Controls("dc_WindowDoorCondition1").AddItem ("SEALS BROKEN")
    frmSystem.Controls("dc_WindowDoorCondition1").AddItem ("AIR DRAFTS")
    
    'TOTAL WINDOW SURFACE AREA
    Call addlabel(lblSurfaceArea, "dc_lblSurfaceArea1", "Total Windows Area", toTop2, toLeft - 15, lblWidth, lblHeight)
    Call addtextbox(SurfaceArea, "dc_SurfaceArea1", toTop2, toLeft1, txtWidth, txtHeight)
    
    'IS WINDOW UV COATED
    Call addlabel(lblWindowUVCoated, "dc_lblWindowUVCoated1", "Window UV coated?", toTop2, toLeft2 - 5, lblWidth, lblHeight)
    Call addcomboBox(WindowUVCoated, "dc_WindowUVCoated1", toTop2, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_WindowUVCoated1").AddItem ("Y")
    frmSystem.Controls("dc_WindowUVCoated1").AddItem ("N")
    
    'NUMBER OF GLAZING
    Call addlabel(lblNumberOfGlazing, "dc_lblNumberOfGlazing1", "Number of glazings", toTop3, toLeft - 15, lblWidth, lblHeight)
    Call addcomboBox(NumberOfGlazing, "dc_NumberOfGlazing1", toTop3, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_NumberOfGlazing1").AddItem ("1")
    frmSystem.Controls("dc_NumberOfGlazing1").AddItem ("2")
    frmSystem.Controls("dc_NumberOfGlazing1").AddItem ("3")
    
End Sub

Private Sub showdooroptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'DOOR TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Door Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("METAL/INSULATED")
    frmSystem.Controls("dc_SystemType1").AddItem ("FIBERGLASS/INSULATED")
    frmSystem.Controls("dc_SystemType1").AddItem ("WOOD")
    frmSystem.Controls("dc_SystemType1").AddItem ("SLIDER")
    frmSystem.Controls("dc_SystemType1").AddItem ("ATRIUM")
    
    'QUANTITY
    Call addlabel(lblQuantity, "dc_lblQuantity1", "Quantity", toTop1, toLeft, lblWidth, lblHeight)
    Call addtextbox(Quantity, "dc_Quantity1", toTop1, toLeft1, txtWidth, txtHeight)
    
    'WINDOW CONDITION
    Call addlabel(lblWindowDoorCondition, "dc_lblWindowDoorCondition1", "Door Condition*", toTop1, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(WindowDoorCondition, "dc_WindowDoorCondition1", toTop1, toLeft3, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_WindowDoorCondition1").AddItem ("NO APPARENT DAMAGE")
    frmSystem.Controls("dc_WindowDoorCondition1").AddItem ("SEALS BROKEN")
    frmSystem.Controls("dc_WindowDoorCondition1").AddItem ("AIR DRAFTS")
End Sub

Private Sub showlightingoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'WINDOW TYPE
    'NOT WINDOW TYPE
    
    'QUANTITY
    Call addlabel(lblQuantity, "dc_lblQuantity1", "Quantity*", toTop, toLeft, lblWidth, lblHeight)
    Call addtextbox(Quantity, "dc_Quantity1", toTop, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM LOCATION
    Call addlabel(lblSystemLocation, "dc_lblSystemLocation1", "System Location*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemLocation, "dc_SystemLocation1", toTop1, toLeft1, 2 * cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemLocation1").AddItem ("BASEMENT")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("BEDROOM")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("DINING ROOM")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("EXTERIOR")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("FAMILY/SITTING ROOM")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("HALLWAY")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("KITCHEN")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("LIVING ROOM")
    frmSystem.Controls("dc_SystemLocation1").AddItem ("BATHROOM TOILET")
    

    'TOTAL WEEKLY OPERATING HOURS
    Call addlabel(lblTotalWeeklyHours, "dc_lblTotalWeeklyHours1", "Total weekly hours", toTop2, toLeft - 5, lblWidth, lblHeight)
    Call addtextbox(TotalWeeklyHours, "dc_TotalWeeklyHours1", toTop2, toLeft1, txtWidth, txtHeight)
    
    'EXISTING BULB WATTAGE
    Call addlabel(lblBulbWattage, "dc_lblBulbWattage1", "Existing Bulb Watt", toTop2, toLeft - 10, lblWidth, lblHeight)
    Call addtextbox(BulbWattage, "dc_BulbWattage1", toTop2, toLeft1, txtWidth, txtHeight)
End Sub

Private Sub showwalloptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'WALL TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Wall Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("ALUMINUM")
    frmSystem.Controls("dc_SystemType1").AddItem ("BRICK")
    frmSystem.Controls("dc_SystemType1").AddItem ("MASONITE")
    frmSystem.Controls("dc_SystemType1").AddItem ("OTHER")
    frmSystem.Controls("dc_SystemType1").AddItem ("STUCCO")
    frmSystem.Controls("dc_SystemType1").AddItem ("VINYL")
    frmSystem.Controls("dc_SystemType1").AddItem ("WOOD")
    
    'INSULATION EXIST INDICATOR
    Call addlabel(lblInsIndicator, "dc_lblInsIndicator1", "Insulation exist indicator*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsIndicator, "dc_InsIndicator1", toTop1, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_InsIndicator1").AddItem ("Y")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("N")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("NOT NEEDED")
    
    'INSULATION TYPE
    Call addlabel(lblInsType, "dc_lblInsType1", "Insulation Type*", toTop1, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(InsType, "dc_InsType1", toTop1, toLeft3, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_InsType1").AddItem ("CELLULOSE")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BATTS")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BLOWN")
    frmSystem.Controls("dc_InsType1").AddItem ("LOOSE FIBERGLASS")
    frmSystem.Controls("dc_InsType1").AddItem ("MINERAL/ROCK WOOL")
    frmSystem.Controls("dc_InsType1").AddItem ("UREA FORMALDAHYDE")
    frmSystem.Controls("dc_InsType1").AddItem (".5 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("2 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("NONE")
    frmSystem.Controls("dc_InsType1").AddItem ("OTHER")
    
    'TANK R-VALUE
    Call addlabel(lblTankRValue, "dc_lblTankRValue1", "Wall R-Value", toTop2, toLeft, lblWidth, lblHeight)
    Call addtextbox(TankRValue, "dc_TankRValue1", toTop2, toLeft1, txtWidth, txtHeight)
    
    'LENGTH
    Call addlabel(lblSystemLength, "dc_lblSystemLength1", "Length", toTop2, toLeft2, lblWidth, lblHeight)
    Call addtextbox(SystemLength, "dc_SystemLength1", toTop2, toLeft3, txtWidth, txtHeight)
    Call addlabel(lblSystemLength1, "dc_lblSystemLength2", "ft", toTop2, toLeft3 + txtWidth + 5, lblWidth, lblHeight)

    'HEIGHT
    Call addlabel(lblSystemHeight, "dc_lblSystemHeight1", "Height", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemHeight, "dc_SystemHeight1", toTop3, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblSystemHeight1, "dc_lblSystemHeight2", "ft", toTop3, toLeft1 + txtWidth + 5, lblWidth, lblHeight)
End Sub

Private Sub showatticoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'ATTIC TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Attic Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("FLOORED")
    frmSystem.Controls("dc_SystemType1").AddItem ("UNFLOORED")
    frmSystem.Controls("dc_SystemType1").AddItem ("KNEE WALL")
    frmSystem.Controls("dc_SystemType1").AddItem ("KW FLAT FLOORED")
    frmSystem.Controls("dc_SystemType1").AddItem ("KW FLAT UNFLOORED")
    frmSystem.Controls("dc_SystemType1").AddItem ("FLAT ROOF")
    frmSystem.Controls("dc_SystemType1").AddItem ("SLOPED")
    
    'TOTAL WINDOW SURFACE AREA
    Call addlabel(lblSurfaceArea, "dc_lblSurfaceArea1", "Square Footage", toTop1, toLeft, lblWidth, lblHeight)
    Call addtextbox(SurfaceArea, "dc_SurfaceArea1", toTop1, toLeft1, txtWidth, txtHeight)
    
    'INSULATION EXIST INDICATOR
    Call addlabel(lblInsIndicator, "dc_lblInsIndicator1", "Insulation exist indicator*", toTop1, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(InsIndicator, "dc_InsIndicator1", toTop1, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_InsIndicator1").AddItem ("Y")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("N")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("NOT NEEDED")
    
    'INSULATION TYPE
    Call addlabel(lblInsType, "dc_lblInsType1", "Insulation Type*", toTop2, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsType, "dc_InsType1", toTop2, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_InsType1").AddItem ("CELLULOSE")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BATTS")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BLOWN")
    frmSystem.Controls("dc_InsType1").AddItem ("LOOSE FIBERGLASS")
    frmSystem.Controls("dc_InsType1").AddItem ("MINERAL/ROCK WOOL")
    frmSystem.Controls("dc_InsType1").AddItem ("UREA FORMALDAHYDE")
    frmSystem.Controls("dc_InsType1").AddItem (".5 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("2 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("NONE")
    frmSystem.Controls("dc_InsType1").AddItem ("OTHER")
    
    'R-VALUE
    Call addlabel(lblTankRValue, "dc_lblTankRValue1", "Attic R-Value", toTop2, toLeft2, lblWidth, lblHeight)
    Call addtextbox(TankRValue, "dc_TankRValue1", toTop2, toLeft3, txtWidth, txtHeight)
    
    'LENGTH
    Call addlabel(lblSystemLength, "dc_lblSystemLength1", "Length", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemLength, "dc_SystemLength1", toTop3, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblSystemLength1, "dc_lblSystemLength2", "ft", toTop2, toLeft3 + txtWidth + 5, lblWidth, lblHeight)

    'HEIGHT
    Call addlabel(lblSystemHeight, "dc_lblSystemHeight1", "Height", toTop3, toLeft2, lblWidth, lblHeight)
    Call addtextbox(SystemHeight, "dc_SystemHeight1", toTop3, toLeft3, txtWidth, txtHeight)
    Call addlabel(lblSystemHeight1, "dc_lblSystemHeight2", "ft", toTop3, toLeft3 + txtWidth + 5, lblWidth, lblHeight)
    
    'VENT REQUIRED
    Call addlabel(lblVentIndicator, "dc_lblVentIndicator1", "Vent Required*", toTop4, toLeft, lblWidth, lblHeight)
    Call addcomboBox(VentIndicator, "dc_VentIndicator1", toTop4, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_VentIndicator1").AddItem ("Y")
    frmSystem.Controls("dc_VentIndicator1").AddItem ("N")
    

    'ACCESS TYPE
    Call addlabel(lblAccessType, "dc_lblAccessType1", "Access Type*", toTop4, toLeft, lblWidth, lblHeight)
    Call addcomboBox(AccessType, "dc_AccessType1", toTop4, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_AccessType1").AddItem ("CEILING")
    frmSystem.Controls("dc_AccessType1").AddItem ("EXTERIOR")
    frmSystem.Controls("dc_AccessType1").AddItem ("KNEE WALL")
    frmSystem.Controls("dc_AccessType1").AddItem ("NO ACCESS AVAILABLE")
    frmSystem.Controls("dc_AccessType1").AddItem ("PULL DOWN STAIRS")
    frmSystem.Controls("dc_AccessType1").AddItem ("TEMPORARY")
    frmSystem.Controls("dc_AccessType1").AddItem ("WALK UP STAIRWAY")
    
    'DEPTH
    Call addlabel(lblSystemDepth, "dc_lblSystemDepth1", "Depth", toTop5, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemDepth, "dc_SystemDepth1", toTop5, toLeft1, txtWidth, txtHeight)
    Call addlabel(lblSystemDepth1, "dc_lblSystemDepth2", "ft", toTop5, toLeft1 + txtWidth + 5, lblWidth, lblHeight)
End Sub

Private Sub showbasementoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'BASEMENT TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Basement Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("CRAWL-OPEN")
    frmSystem.Controls("dc_SystemType1").AddItem ("CRAWL-CLOSED")
    frmSystem.Controls("dc_SystemType1").AddItem ("FULL")
    frmSystem.Controls("dc_SystemType1").AddItem ("GARAGE")
    frmSystem.Controls("dc_SystemType1").AddItem ("SLAB")

    'TOTAL AREA
    Call addlabel(lblSurfaceArea, "dc_lblSurfaceArea1", "Square footage", toTop1, toLeft - 10, lblWidth, lblHeight)
    Call addtextbox(SurfaceArea, "dc_SurfaceArea1", toTop1, toLeft1, txtWidth, txtHeight)
    

    'Perimeter Footage
    Call addlabel(lblPerimeterFootage, "dc_lblPerimeterFootage1", "Perimeter footage", toTop1, toLeft2, lblWidth, lblHeight)
    Call addtextbox(PerimeterFootage, "dc_PerimeterFootage1", toTop1, toLeft3, txtWidth, txtHeight)
    
    'INSULATION EXIST INDICATOR
    Call addlabel(lblInsIndicator, "dc_lblInsIndicator1", "Insulation exist indicator*", toTop2, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsIndicator, "dc_InsIndicator1", toTop2, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_InsIndicator1").AddItem ("Y")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("N")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("NOT NEEDED")
    
    'INSULATION TYPE
    Call addlabel(lblInsType, "dc_lblInsType1", "Insulation Type*", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(InsType, "dc_InsType1", toTop2, toLeft3, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_InsType1").AddItem ("CELLULOSE")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BATTS")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BLOWN")
    frmSystem.Controls("dc_InsType1").AddItem ("LOOSE FIBERGLASS")
    frmSystem.Controls("dc_InsType1").AddItem ("MINERAL/ROCK WOOL")
    frmSystem.Controls("dc_InsType1").AddItem ("UREA FORMALDAHYDE")
    frmSystem.Controls("dc_InsType1").AddItem (".5 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("2 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("NONE")
    frmSystem.Controls("dc_InsType1").AddItem ("OTHER")
    
    'R-VALUE
    Call addlabel(lblTankRValue, "dc_lblTankRValue1", "Floor R-Value", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(TankRValue, "dc_TankRValue1", toTop3, toLeft1, txtWidth, txtHeight)
    

    'BASEMENT AIR CONDITIONED
    Call addlabel(lblBasementAC, "dc_lblBasementAC1", "Basement AC", toTop3, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(BasementAC, "dc_BasementAC1", toTop3, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_BasementAC1").AddItem ("Y")
    frmSystem.Controls("dc_BasementAC1").AddItem ("N")
    
    'RIM JOIST INSULATION RECOMMENDED
    Call addlabel(lblRJInsRecommended, "dc_lblRJInsRecommended1", "Rim joist insulation recommended?", toTop4, toLeft - 15, lblWidth, lblHeight)
    Call addcomboBox(RJInsRecommended, "dc_RJInsRecommended1", toTop4, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_RJInsRecommended1").AddItem ("Y")
    frmSystem.Controls("dc_RJInsRecommended1").AddItem ("N")
    
End Sub
Private Sub showbwoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'BASEMENT WALL TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Basement Wall Type*", toTop, toLeft - 15, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("CINDER BLOCK")
    frmSystem.Controls("dc_SystemType1").AddItem ("CONCRETE POUR/FORMED")
    frmSystem.Controls("dc_SystemType1").AddItem ("FRAMED 2x4")
    frmSystem.Controls("dc_SystemType1").AddItem ("FRAMED 2x6")
    
    'R-VALUE
    Call addlabel(lblTankRValue, "dc_lblTankRValue1", "Basement wall R-Value", toTop1, toLeft, lblWidth, lblHeight)
    Call addtextbox(TankRValue, "dc_TankRValue1", toTop1, toLeft1, txtWidth, txtHeight)
    
    'INSULATION EXIST INDICATOR
    Call addlabel(lblInsIndicator, "dc_lblInsIndicator1", "Insulation exist indicator*", toTop1, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(InsIndicator, "dc_InsIndicator1", toTop1, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_InsIndicator1").AddItem ("Y")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("N")
    frmSystem.Controls("dc_InsIndicator1").AddItem ("NOT NEEDED")

    'INSULATION TYPE
    Call addlabel(lblInsType, "dc_lblInsType1", "Insulation Type*", toTop2, toLeft, lblWidth, lblHeight)
    Call addcomboBox(InsType, "dc_InsType1", toTop2, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_InsType1").AddItem ("CELLULOSE")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BATTS")
    frmSystem.Controls("dc_InsType1").AddItem ("FIBERGLASS BLOWN")
    frmSystem.Controls("dc_InsType1").AddItem ("LOOSE FIBERGLASS")
    frmSystem.Controls("dc_InsType1").AddItem ("MINERAL/ROCK WOOL")
    frmSystem.Controls("dc_InsType1").AddItem ("UREA FORMALDAHYDE")
    frmSystem.Controls("dc_InsType1").AddItem (".5 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("2 LB FOAM")
    frmSystem.Controls("dc_InsType1").AddItem ("NONE")
    frmSystem.Controls("dc_InsType1").AddItem ("OTHER")
End Sub
Private Sub showrefrigeratoroptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'REFRIGERATOR TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Refrigerator Type*", toTop, toLeft - 10, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("SIDE BY SIDE")
    frmSystem.Controls("dc_SystemType1").AddItem ("FREEZER TOP")
    frmSystem.Controls("dc_SystemType1").AddItem ("FREEZER BOTTOM")
    frmSystem.Controls("dc_SystemType1").AddItem ("SINGLE DOOR")

    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop1, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemSize, "dc_SystemSize1", toTop1, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "dc_lblSizeUnit1", "System Size Unit*", toTop1, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(SizeUnit, "dc_SizeUnit1", toTop1, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_SizeUnit1").AddItem ("GALLONS")
    
    'SYSTEM Age
    Call addlabel(lblSystemAge, "dc_lblSystemAge1", "System Age*", toTop2, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemAge, "dc_SystemAge1", toTop2, toLeft1, txtWidth, txtHeight)
    

    'DEFROST TYPE
    Call addlabel(lblDefrostType, "dc_lblDefrostType1", "Defrost Type*", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(DefrostType, "dc_DefrostType1", toTop2, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_DefrostType1").AddItem ("AUTOMATIC")
    frmSystem.Controls("dc_DefrostType1").AddItem ("MANUAL")

    'MAKE
    Call addlabel(lblSystemMake, "dc_lblSystemMake1", "Make (Manufacturer)*", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemMake, "dc_SystemMake1", toTop3, toLeft1, txtWidth, txtHeight)


    'SYSTEM Age
    Call addlabel(lblSystemMeteredUsage, "dc_lblSystemMeteredUsage1", "Metered Usage*", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemMeteredUsage, "dc_SystemMeteredUsage1", toTop3, toLeft1, txtWidth, txtHeight)

End Sub
Private Sub showfreezeroptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'FREEZER TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Freezer Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("UPRIGHT")
    frmSystem.Controls("dc_SystemType1").AddItem ("CHEST")

    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop1, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemSize, "dc_SystemSize1", toTop1, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "dc_lblSizeUnit1", "System Size Unit*", toTop1, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(SizeUnit, "dc_SizeUnit1", toTop1, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_SizeUnit1").AddItem ("GALLONS")
    
    'SYSTEM Age
    Call addlabel(lblSystemAge, "dc_lblSystemAge1", "System Age*", toTop2, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemAge, "dc_SystemAge1", toTop2, toLeft1, txtWidth, txtHeight)
    

    'DEFROST TYPE
    Call addlabel(lblDefrostType, "dc_lblDefrostType1", "Defrost Type*", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(DefrostType, "dc_DefrostType1", toTop2, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_DefrostType1").AddItem ("AUTOMATIC")
    frmSystem.Controls("dc_DefrostType1").AddItem ("MANUAL")
End Sub
Private Sub showapplianceoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    'APPLIANCE TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Appliance Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("AQUARIUM")
    frmSystem.Controls("dc_SystemType1").AddItem ("ATTIC FAN")
    frmSystem.Controls("dc_SystemType1").AddItem ("BLACK & WHITE TV")
    frmSystem.Controls("dc_SystemType1").AddItem ("CEILING FAN")
    frmSystem.Controls("dc_SystemType1").AddItem ("CLOTHES WASHER")
    frmSystem.Controls("dc_SystemType1").AddItem ("COLOR TV")
    frmSystem.Controls("dc_SystemType1").AddItem ("COMPUTER")
    frmSystem.Controls("dc_SystemType1").AddItem ("DEHUMIDIFIER")
    frmSystem.Controls("dc_SystemType1").AddItem ("DISHWASHER")
    frmSystem.Controls("dc_SystemType1").AddItem ("ELECTRIC SPACE HEATER")
    frmSystem.Controls("dc_SystemType1").AddItem ("ELEC CLOTHES DRYER")
    frmSystem.Controls("dc_SystemType1").AddItem ("ELECTRIC BLANKET")
    frmSystem.Controls("dc_SystemType1").AddItem ("ELECTRIC COOKING")
    frmSystem.Controls("dc_SystemType1").AddItem ("FAX MACHINE")
    frmSystem.Controls("dc_SystemType1").AddItem ("GAS CLOTHES DRYER")
    frmSystem.Controls("dc_SystemType1").AddItem ("GAS COOKING")
    frmSystem.Controls("dc_SystemType1").AddItem ("HOT TUB")
    frmSystem.Controls("dc_SystemType1").AddItem ("HUMIDIFIER")
    frmSystem.Controls("dc_SystemType1").AddItem ("LASER PRINTER")
    frmSystem.Controls("dc_SystemType1").AddItem ("MICROWAVE")
    frmSystem.Controls("dc_SystemType1").AddItem ("MISCELLANEOUS")
    frmSystem.Controls("dc_SystemType1").AddItem ("POOL PUMP")
    frmSystem.Controls("dc_SystemType1").AddItem ("PRINTER")
    frmSystem.Controls("dc_SystemType1").AddItem ("STEREO")
    frmSystem.Controls("dc_SystemType1").AddItem ("SUMP PUMP")
    frmSystem.Controls("dc_SystemType1").AddItem ("WATERBED")
    frmSystem.Controls("dc_SystemType1").AddItem ("WELL PUMP")
    
    'QUANTITY
    Call addlabel(lblQuantity, "dc_lblQuantity1", "Quantity", toTop1, toLeft, lblWidth, lblHeight)
    Call addtextbox(Quantity, "dc_Quantity1", toTop1, toLeft1, txtWidth, txtHeight)
End Sub
            
Private Sub showcoolingoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    ' COOLING TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Cooling Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("CENTRAL AC")
    frmSystem.Controls("dc_SystemType1").AddItem ("HEAT PUMP-AIR SOURCE")
    frmSystem.Controls("dc_SystemType1").AddItem ("HEAT PUMP-WATER SOURCE")
    frmSystem.Controls("dc_SystemType1").AddItem ("SPLIT SYSTEM")
    frmSystem.Controls("dc_SystemType1").AddItem ("WINDOW AC")
    
'    loctextboxlen = 50
'    cboWidth = 70
    
    ' FUEL SOURCE
    Call addlabel(lblFuelSource, "dc_lblFuelSource1", "Fuel Source*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(FuelSource, "dc_FuelSource1", toTop1, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_FuelSource1").AddItem ("ELECTRIC")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop2, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemSize, "dc_SystemSize1", toTop2, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "dc_lblSizeUnit1", "System Size Unit*", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(SizeUnit, "dc_SizeUnit1", toTop2, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_SizeUnit1").AddItem ("BTU")
    frmSystem.Controls("dc_SizeUnit1").AddItem ("MBTU")
    frmSystem.Controls("dc_SizeUnit1").AddItem ("MMBTU")
    frmSystem.Controls("dc_SizeUnit1").AddItem ("TON")

    'SYSTEM Age
    Call addlabel(lblSystemAge, "dc_lblSystemAge1", "System Age*", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemAge, "dc_SystemAge1", toTop3, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM Efficiency Rating
    Call addlabel(lblEffRating, "dc_lblEffRating1", "Efficiency Rating*", toTop4, toLeft - 10, lblWidth, lblHeight)
    Call addtextbox(EffRating, "dc_EffRating1", toTop4, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM Efficiency Rating Type
    Call addlabel(lblEffRatingType, "dc_lblEffRatingType1", "Rating Type*", toTop4, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(EffRatingType, "dc_EffRatingType1", toTop4, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_EffRatingType1").AddItem ("EER")
    frmSystem.Controls("dc_EffRatingType1").AddItem ("SEER")
    frmSystem.Controls("dc_EffRatingType1").AddItem ("COP")
    
    'TOTAL PERCENTAGE OF SPACE COOLED
    Call addlabel(lblPercentageCooled, "dc_lblPercentageCooled1", "% of space cooled*", toTop5, toLeft - 13, lblWidth, lblHeight)
    Call addtextbox(PercentageCooled, "dc_PercentageCooled1", toTop5, toLeft1, txtWidth, txtHeight)
    
    'FREQUENCY OF SYSTEM USE

    Call addlabel(lblFrequencyUse, "dc_lblFrequencyUse1", "Frequency of use*", toTop5, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(FrequencyUse, "dc_FrequencyUse1", toTop5, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_FrequencyUse1").AddItem ("0%")
    frmSystem.Controls("dc_FrequencyUse1").AddItem ("10-30%")
    frmSystem.Controls("dc_FrequencyUse1").AddItem ("31-70%")
    frmSystem.Controls("dc_FrequencyUse1").AddItem ("71-100%")
    
    'TOTAL UNITS USED
    Call addlabel(lblTotalUnits, "dc_lblTotalUnits1", "Total units used", toTop6, toLeft, lblWidth, lblHeight)
    Call addtextbox(TotalUnits, "dc_TotalUnits1", toTop6, toLeft1, txtWidth, txtHeight)
        
    'QUANTITY
    Call addlabel(lblQuantity, "dc_lblQuantity1", "Quantity", toTop6, toLeft2, lblWidth, lblHeight)
    Call addtextbox(Quantity, "dc_Quantity1", toTop6, toLeft3, txtWidth, txtHeight)
        
End Sub

Private Sub showheatingoptions()
    'SYSTEM NOT APPLICABLE VALUE
    Call addlabel(lblSystemApplicable, "dc_lblSystemApplicable1", "System Applicable", toTop0, toLeft - 15, lblWidth * 2, lblHeight)
    Call addcomboBox(SystemApplicable, "dc_SystemApplicable1", toTop0, toLeft1, cboWidth, cboHeight)
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("N/A")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("X")
    frmSystem.Controls("dc_SystemApplicable1").AddItem ("BLANK")
    
    
    ' HEATING TYPE
    Call addlabel(lblSystemType, "dc_lblSystemType1", "Heating Type*", toTop, toLeft, lblWidth, lblHeight)
    Call addcomboBox(SystemType, "dc_SystemType1", toTop, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_SystemType1").AddItem ("GAS FURNACE")
    frmSystem.Controls("dc_SystemType1").AddItem ("HEAT PUMP-AIR SOURCE")
    frmSystem.Controls("dc_SystemType1").AddItem ("HEAT PUMP-GROUND SOURCE")
    frmSystem.Controls("dc_SystemType1").AddItem ("HEAT PUMP-DUAL FUEL")
    frmSystem.Controls("dc_SystemType1").AddItem ("RESISTANCE ELECTRIC HEAT")
    frmSystem.Controls("dc_SystemType1").AddItem ("HOT WATER BOILER")
    frmSystem.Controls("dc_SystemType1").AddItem ("FORCED AIR")
    frmSystem.Controls("dc_SystemType1").AddItem ("STEAM")
    frmSystem.Controls("dc_SystemType1").AddItem ("WOOD/COAL STOVE")
    
    ' FUEL SOURCE
    Call addlabel(lblFuelSource, "dc_lblFuelSource1", "Fuel Source*", toTop1, toLeft, lblWidth, lblHeight)
    Call addcomboBox(FuelSource, "dc_FuelSource1", toTop1, toLeft1, cboWidth * 2, cboHeight)
    frmSystem.Controls("dc_FuelSource1").AddItem ("ELECTRIC")
    frmSystem.Controls("dc_FuelSource1").AddItem ("GAS")
    frmSystem.Controls("dc_FuelSource1").AddItem ("PROPANE")
    frmSystem.Controls("dc_FuelSource1").AddItem ("CENTRAL STEAM")
    frmSystem.Controls("dc_FuelSource1").AddItem ("COAL")
    frmSystem.Controls("dc_FuelSource1").AddItem ("SOLAR")
    frmSystem.Controls("dc_FuelSource1").AddItem ("WOOD")
    frmSystem.Controls("dc_FuelSource1").AddItem ("OIL")
    frmSystem.Controls("dc_FuelSource1").AddItem ("OTHER")
    
    'SYSTEM SIZE
    Call addlabel(lblSystemSize, "dc_lblSystemSize1", "System Size*", toTop2, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemSize, "dc_SystemSize1", toTop2, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM SIZE UNIT
    Call addlabel(lblSizeUnit, "dc_lblSizeUnit1", "System Size Unit*", toTop2, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(SizeUnit, "dc_SizeUnit1", toTop2, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_SizeUnit1").AddItem ("MBTU")
    frmSystem.Controls("dc_SizeUnit1").AddItem ("MMBTU")
    frmSystem.Controls("dc_SizeUnit1").AddItem ("TON")

    'SYSTEM Age
    Call addlabel(lblSystemAge, "dc_lblSystemAge1", "System Age", toTop3, toLeft, lblWidth, lblHeight)
    Call addtextbox(SystemAge, "dc_SystemAge1", toTop3, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM Efficiency Rating
    Call addlabel(lblEffRating, "dc_lblEffRating1", "Efficiency Rating", toTop4, toLeft - 10, lblWidth, lblHeight)
    Call addtextbox(EffRating, "dc_EffRating1", toTop4, toLeft1, txtWidth, txtHeight)
    
    'SYSTEM Efficiency Rating Type
    Call addlabel(lblEffRatingType, "dc_lblEffRatingType1", "Rating Type*", toTop4, toLeft2, lblWidth, lblHeight)
    Call addcomboBox(EffRatingType, "dc_EffRatingType1", toTop4, toLeft3, cboWidth, cboHeight)
    frmSystem.Controls("dc_EffRatingType1").AddItem ("AFUE")
    frmSystem.Controls("dc_EffRatingType1").AddItem ("HSPF")
    frmSystem.Controls("dc_EffRatingType1").AddItem ("COP")
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

'    If cboSystem.ListIndex < 0 Then
'        prompt = "Heating system"
'    End If
    
    stv = frmSystem.Controls("dc_SystemType1").Value
    If stv = "GAS FURNACE" Or stv = "HEAT PUMP-AIR SOURCE" Or stv = "HEAT PUMP-GROUND SOURCE" _
        Or stv = "HEAT PUMP-DUAL FUEL" Or stv = "RESISTANCE ELECTRIC HEAT" Or stv = "HOT WATER BOILER" _
        Or stv = "FORCED AIR" Or stv = "STEAM" Or stv = "WOOD/COAL STOVE" Then
    Else
        errorstring ("System Type")
    End If
    
    fs = frmSystem.Controls("dc_FuelSource1").Value
    If fs = "ELECTRIC" Or fs = "GAS" Or fs = "PROPANE" Or fs = "CENTRAL STEAM" Or fs = "COAL" Or fs = "SOLAR" _
        Or fs = "WOOD" Or fs = "OIL" Or fs = "OTHER" Then
    Else
        errorstring ("Fuel Source")
    End If
    
    If Not IsNumeric(frmSystem.Controls("dc_SystemSize1").Value) Then
        errorstring ("System Size")
    End If

    su = frmSystem.Controls("dc_SizeUnit1").Value
    If su = "MBTU" Or su = "MMBTU" Or su = "TON" Then
    Else
        errorstring ("Size Unit")
    End If
    
    If IsNumeric(frmSystem.Controls("dc_SystemAge1").Value) Or frmSystem.Controls("dc_SystemAge1").Value = "" Then
    Else
        errorstring ("System Age")
    End If
    
    If Not IsNumeric(frmSystem.Controls("dc_EffRating1").Value) Then
        errorstring ("Efficiency Rating")
    End If
    
    et = frmSystem.Controls("dc_EffRatingType1").Value
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
        iReply = MsgBox("Please select the system to load", vbOKOnly, "Please select a system in the system list!")
        Exit Sub
    End If
        
    auditlastrow = Worksheets(AuditSheetName).Range("E" & Rows.Count).End(xlUp).Row
    auditcurrentrow = lstSelectedSystems.ListIndex + 2
    strSystem = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name).Value
    cboSystem.Text = strSystem
    strCurrentSystemName = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type).Value
    Select Case strSystem
        Case "HEATING"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_FuelSource1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Fuel_Source)
            frmSystem.Controls("dc_SystemSize1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size)
            frmSystem.Controls("dc_SizeUnit1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure)
            frmSystem.Controls("dc_SystemAge1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age)
            frmSystem.Controls("dc_EffRating1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Efficiency_Rating)
            frmSystem.Controls("dc_EffRatingType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Efficiency_Rating_Type)
        Case "COOLING"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_FuelSource1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Fuel_Source)
            frmSystem.Controls("dc_SystemSize1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size)
            frmSystem.Controls("dc_SizeUnit1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure)
            frmSystem.Controls("dc_SystemAge1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age)
            frmSystem.Controls("dc_EffRating1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Efficiency_Rating)
            frmSystem.Controls("dc_EffRatingType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Efficiency_Rating_Type)
            frmSystem.Controls("dc_PercentageCooled1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Percent_of_space_heated_or_cooled)
            frmSystem.Controls("dc_FrequencyUse1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Frequency_of_system_use)
            frmSystem.Controls("dc_TotalUnits1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Total_units_used)
            frmSystem.Controls("dc_Quantity1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity)
        Case "HVAC DISTRIBUTION"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_SystemSize1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size)
            frmSystem.Controls("dc_InsIndicator1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator)
            frmSystem.Controls("dc_InsType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type)
            frmSystem.Controls("dc_SystemLocation1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Location)
            frmSystem.Controls("dc_SystemLength1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Length)
            frmSystem.Controls("dc_FlexCondition1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Condition_of_flex_duct)
        Case "WATER HEATER"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_FuelSource1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Fuel_Source)
            frmSystem.Controls("dc_SystemSize1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size)
            frmSystem.Controls("dc_SizeUnit1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure)
            frmSystem.Controls("dc_SystemAge1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age)
            frmSystem.Controls("dc_InsIndicator1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator)
            frmSystem.Controls("dc_InsType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type)
            frmSystem.Controls("dc_TankRValue1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value)
            frmSystem.Controls("dc_PercentageLoad1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Percent_of_Load)
            frmSystem.Controls("dc_TemperatureSetting1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Current_temperature_setting)
            frmSystem.Controls("dc_EnergyFactor1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Energy_Factor)

        Case "THERMOSTAT"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_PercentageLoad1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Percent_of_Load)
            frmSystem.Controls("dc_AverageCoolingTemperature1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Average_cooling_temperature)
            frmSystem.Controls("dc_AverageHeatingTemperature1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Average_heating_temperature)
            frmSystem.Controls("dc_DaytimeSetback1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Day_setback_indicator)
            frmSystem.Controls("dc_EveningSetback1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Evening_setback_indicator)
            frmSystem.Controls("dc_NightSetback1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Night_setback_indicator)
            frmSystem.Controls("dc_HeatingDayTemperature1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Heating_day_temperature_setting)
            frmSystem.Controls("dc_HeatingEveningTemperature1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Heating_evening_temperature_setting)
            frmSystem.Controls("dc_HeatingNightTemperature1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Heating_night_temperature_setting)
            frmSystem.Controls("dc_CoolingDayTemperature1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Cooling_day_temperature_setting)
            frmSystem.Controls("dc_CoolingEveningTemperature1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Cooling_evening_temperature_setting)
            frmSystem.Controls("dc_CoolingNightTemperature1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Cooling_night_temperature_setting)
            frmSystem.Controls("dc_ACCtrlPresent1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.AC_load_control_present_indicator)

        Case "WINDOW"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_Quantity1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity)
            frmSystem.Controls("dc_WindowDoorCondition1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Condition_of_window_or_door)
            frmSystem.Controls("dc_SurfaceArea1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Square_Footage)
            frmSystem.Controls("dc_WindowUVCoated1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Window_UV_coated_indicator)
            frmSystem.Controls("dc_NumberOfGlazing1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Number_of_window_glazings)

        Case "DOOR"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_Quantity1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity)
            frmSystem.Controls("dc_WindowDoorCondition1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Condition_of_window_or_door)

        Case "LIGHTING"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_Quantity1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity)
            frmSystem.Controls("dc_SystemLocation1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Location)
            frmSystem.Controls("dc_TotalWeeklyHours1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Weekly_operating_hours)
            frmSystem.Controls("dc_BulbWattage1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Existing_bulb_wattage)

        Case "WALL"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_InsIndicator1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator)
            frmSystem.Controls("dc_InsType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type)
            frmSystem.Controls("dc_TankRValue1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value)
            frmSystem.Controls("dc_SystemLength1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Length)
            frmSystem.Controls("dc_SystemHeight1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.height)

        Case "ATTIC"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_SurfaceArea1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Total_window_surface_area)
            frmSystem.Controls("dc_InsIndicator1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator)
            frmSystem.Controls("dc_InsType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type)
            frmSystem.Controls("dc_TankRValue1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value)
            frmSystem.Controls("dc_SystemLength1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Length)
            frmSystem.Controls("dc_SystemHeight1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.height)
            frmSystem.Controls("dc_VentIndicator1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Vent_required_indicator)
            frmSystem.Controls("dc_AccessType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Access_Type)
            frmSystem.Controls("dc_SystemDepth1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Depth)

        Case "BASEMENT"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_SurfaceArea1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Square_Footage)
            frmSystem.Controls("dc_PerimeterFootage1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Perimeter_Footage)
            frmSystem.Controls("dc_InsIndicator1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator)
            frmSystem.Controls("dc_InsType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type)
            frmSystem.Controls("dc_TankRValue1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value)
            frmSystem.Controls("dc_BasementAC1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Basement_air_conditioned_indicator)
            frmSystem.Controls("dc_RJInsRecommended1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Rim_joist_recommended_indicator)
        

        Case "BASEMENT WALL"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_TankRValue1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value)
            frmSystem.Controls("dc_InsIndicator1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator)
            frmSystem.Controls("dc_InsType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type)

        Case "REFRIGERATOR"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_SystemSize1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size)
            frmSystem.Controls("dc_SizeUnit1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure)
            frmSystem.Controls("dc_SystemAge1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age)
            frmSystem.Controls("dc_DefrostType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Defrost_Type)
            frmSystem.Controls("dc_SystemMake1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_make_manufacturer)
            frmSystem.Controls("dc_SystemMeteredUsage1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Metered_Usage)
            
        Case "FREEZER"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_SystemSize1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size)
            frmSystem.Controls("dc_SizeUnit1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure)
            frmSystem.Controls("dc_SystemAge1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age)
            frmSystem.Controls("dc_DefrostType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Defrost_Type)
        
        Case "APPLIANCE"
            frmSystem.Controls("dc_SystemApplicable1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable)
            frmSystem.Controls("dc_SystemType1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type)
            frmSystem.Controls("dc_Quantity1").Value = Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity)
    
    End Select
    
    bSystemLoad = True
    oldSystemName = cboSystem.Text
End Sub

Private Sub cmdNew_Click()
    bSystemLoad = False
    cboSystem.Text = ""
    strCurrentSystemName = ""
    'auditcurrentrow = Worksheets(AuditSheetName).Range("E" & Rows.Count).End(xlUp).Row + 1
    Call updatelistbox
End Sub

Private Sub cmdOK_Click()
    Dim flag As Boolean
    If lstSelectedSystems.ListIndex <> -1 Then
        auditcurrentrow = lstSelectedSystems.ListIndex + 2
    Else
        auditlastrow = Worksheets(AuditSheetName).Range("E" & Rows.Count).End(xlUp).Row
        auditcurrentrow = auditlastrow + 1
    End If
    Select Case cboSystem
        Case "HEATING"
            If heatingvalidation = True Then
                Call saveheatingsystem
            End If
        Case "COOLING"
            Call savecoolingsystem
        Case "HVAC DISTRIBUTION"
            Call savehvacdistribution
        Case "WATER HEATER"
            Call savewh
        Case "THERMOSTAT"
            Call savethermo
        Case "WINDOW"
            Call savewindow
        Case "DOOR"
            Call savedoor
        Case "LIGHTING"
            Call savelighting
        Case "WALL"
            Call savewall
        Case "ATTIC"
            Call saveattic
        Case "BASEMENT"
            Call savebasement
        Case "BASEMENT WALL"
            Call savebasementwall
        Case "REFRIGERATOR"
            Call saverefrigerator
        Case "FREEZER"
            Call savefreezer
        Case "APPLIANCE"
            Call saveappliance
        Case Else
            MsgBox "The system type is invalid."
    End Select
    Call updatelistbox
End Sub

Private Sub savewh()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addwh
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writewh
        Else
            Call addwh
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addwh()
        If iWH < 3 Then
            iWH = iWH + 1
            strCurrentSystemName = "WATER HEATER-" + CStr(iWH)
            Call writewh
        Else
            MsgBox ("You can only enter at most 3 water heaters!")
        End If
End Sub
Private Sub writewh()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "WATER HEATER"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Fuel_Source) = frmSystem.Controls("dc_FuelSource1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size) = frmSystem.Controls("dc_SystemSize1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure) = frmSystem.Controls("dc_SizeUnit1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age) = frmSystem.Controls("dc_SystemAge1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator) = frmSystem.Controls("dc_InsIndicator1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type) = frmSystem.Controls("dc_InsType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value) = frmSystem.Controls("dc_TankRValue1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Percent_of_Load) = frmSystem.Controls("dc_PercentageLoad1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Current_temperature_setting) = frmSystem.Controls("dc_TemperatureSetting1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Energy_Factor) = frmSystem.Controls("dc_EnergyFactor1").Value
End Sub
Private Sub savethermo()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addthermo
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writethermo
        Else
            Call addthermo
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addthermo()
        If iThermostat < 3 Then
            iThermostat = iThermostat + 1
            strCurrentSystemName = "THERMOSTAT-" + CStr(iThermostat)
            Call writethermo
        Else
            MsgBox ("You can only enter at most 3 thermostats!")
        End If
End Sub
Private Sub writethermo()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "THERMOSTAT"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Percent_of_Load) = frmSystem.Controls("dc_PercentageLoad1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Average_cooling_temperature) = frmSystem.Controls("dc_AverageCoolingTemperature1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Average_heating_temperature) = frmSystem.Controls("dc_AverageHeatingTemperature1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Day_setback_indicator) = frmSystem.Controls("dc_DaytimeSetback1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Evening_setback_indicator) = frmSystem.Controls("dc_EveningSetback1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Night_setback_indicator) = frmSystem.Controls("dc_NightSetback1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Heating_day_temperature_setting) = frmSystem.Controls("dc_HeatingDayTemperature1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Heating_evening_temperature_setting) = frmSystem.Controls("dc_HeatingEveningTemperature1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Heating_night_temperature_setting) = frmSystem.Controls("dc_HeatingNightTemperature1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Cooling_day_temperature_setting) = frmSystem.Controls("dc_CoolingDayTemperature1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Cooling_evening_temperature_setting) = frmSystem.Controls("dc_CoolingEveningTemperature1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Cooling_night_temperature_setting) = frmSystem.Controls("dc_CoolingNightTemperature1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.AC_load_control_present_indicator) = frmSystem.Controls("dc_ACCtrlPresent1").Value
End Sub

Private Sub savewindow()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addwindow
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writewindow
        Else
            Call addwindow
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addwindow()
        If iWindow < 5 Then
            iWindow = iWindow + 1
            strCurrentSystemName = "WINDOW-" + CStr(iWindow)
            Call writewindow
        Else
            MsgBox ("You can only enter at most 5 windows!")
        End If
End Sub
Private Sub writewindow()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "WINDOW"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity) = frmSystem.Controls("dc_Quantity1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Condition_of_window_or_door) = frmSystem.Controls("dc_WindowDoorCondition1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Square_Footage) = frmSystem.Controls("dc_SurfaceArea1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Window_UV_coated_indicator) = frmSystem.Controls("dc_WindowUVCoated1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Number_of_window_glazings) = frmSystem.Controls("dc_NumberOfGlazing1").Value
End Sub
Private Sub savedoor()
    If lstSelectedSystems.ListIndex = -1 Then
        Call adddoor
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writedoor
        Else
            Call adddoor
        End If
    End If
    Call updatelistbox
End Sub
Private Sub adddoor()
        If iDoor < 5 Then
            iDoor = iDoor + 1
            strCurrentSystemName = "DOOR-" + CStr(iDoor)
            Call writedoor
        Else
            MsgBox ("You can only enter at most 5 doors!")
        End If
End Sub
Private Sub writedoor()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "DOOR"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity) = frmSystem.Controls("dc_Quantity1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Condition_of_window_or_door) = frmSystem.Controls("dc_WindowDoorCondition1").Value
End Sub
Private Sub savelighting()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addlighting
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writelighting
        Else
            Call addlighting
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addlighting()
        If iLighting < 4 Then
            iLighting = iLighting + 1
            strCurrentSystemName = "LIGHTING-" + CStr(iLighting)
            Call writelighting
        Else
            MsgBox ("You can only enter at most 4 lighting systems!")
        End If
End Sub
Private Sub writelighting()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "LIGHTING"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity) = frmSystem.Controls("dc_Quantity1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Location) = frmSystem.Controls("dc_SystemLocation1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Weekly_operating_hours) = frmSystem.Controls("dc_TotalWeeklyHours1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Existing_bulb_wattage) = frmSystem.Controls("dc_BulbWattage1").Value
End Sub
Private Sub savewall()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addwall
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writewall
        Else
            Call addwall
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addwall()
        If iWall < 4 Then
            iWall = iWall + 1
            strCurrentSystemName = "WALL-" + CStr(iWall)
            Call writewall
        Else
            MsgBox ("You can only enter at most 4 walls!")
        End If
End Sub
Private Sub writewall()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "WALL"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator) = frmSystem.Controls("dc_InsIndicator1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type) = frmSystem.Controls("dc_InsType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value) = frmSystem.Controls("dc_TankRValue1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Length) = frmSystem.Controls("dc_SystemLength1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.height) = frmSystem.Controls("dc_SystemHeight1").Value
End Sub
Private Sub saveattic()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addattic
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writeattic
        Else
            Call addattic
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addattic()
        If iAttic < 4 Then
            iAttic = iAttic + 1
            strCurrentSystemName = "ATTIC-" + CStr(iAttic)
            Call writeattic
        Else
            MsgBox ("You can only enter at most 4 attics!")
        End If
End Sub
Private Sub writeattic()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "ATTIC"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Total_window_surface_area) = frmSystem.Controls("dc_SurfaceArea1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator) = frmSystem.Controls("dc_InsIndicator1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type) = frmSystem.Controls("dc_InsType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value) = frmSystem.Controls("dc_TankRValue1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Length) = frmSystem.Controls("dc_SystemLength1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.height) = frmSystem.Controls("dc_SystemHeight1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Vent_required_indicator) = frmSystem.Controls("dc_VentIndicator1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Access_Type) = frmSystem.Controls("dc_AccessType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Depth) = frmSystem.Controls("dc_SystemDepth1").Value
End Sub
Private Sub savebasement()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addbasement
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writebasement
        Else
            Call addbasement
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addbasement()
        If iBasement < 3 Then
            iBasement = iBasement + 1
            strCurrentSystemName = "BASEMENT-" + CStr(iBasement)
            Call writebasement
        Else
            MsgBox ("You can only enter at most 3 basements!")
        End If
End Sub
Private Sub writebasement()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "BASEMENT"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Square_Footage) = frmSystem.Controls("dc_SurfaceArea1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Perimeter_Footage) = frmSystem.Controls("dc_PerimeterFootage1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator) = frmSystem.Controls("dc_InsIndicator1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type) = frmSystem.Controls("dc_InsType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value) = frmSystem.Controls("dc_TankRValue1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Basement_air_conditioned_indicator) = frmSystem.Controls("dc_BasementAC1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Rim_joist_recommended_indicator) = frmSystem.Controls("dc_RJInsRecommended1").Value
End Sub
Private Sub savebasementwall()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addbasementwall
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writebasementwall
        Else
            Call addbasementwall
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addbasementwall()
        If iBW < 3 Then
            iBW = iBW + 1
            strCurrentSystemName = "BASEMENT WALL-" + CStr(iBW)
            Call writebasementwall
        Else
            MsgBox ("You can only enter at most 3 basementwalls!")
        End If
End Sub
Private Sub writebasementwall()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "BASEMENT WALL"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.R_Value) = frmSystem.Controls("dc_TankRValue1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator) = frmSystem.Controls("dc_InsIndicator1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type) = frmSystem.Controls("dc_InsType1").Value
End Sub
Private Sub saverefrigerator()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addrefrigerator
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writerefrigerator
        Else
            Call addrefrigerator
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addrefrigerator()
        If iRefrigerator < 3 Then
            iRefrigerator = iRefrigerator + 1
            strCurrentSystemName = "REFRIGERATOR-" + CStr(iRefrigerator)
            Call writerefrigerator
        Else
            MsgBox ("You can only enter at most 3 refrigerators!")
        End If
End Sub
Private Sub writerefrigerator()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "REFRIGERATOR"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size) = frmSystem.Controls("dc_SystemSize1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure) = frmSystem.Controls("dc_SizeUnit1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age) = frmSystem.Controls("dc_SystemAge1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Defrost_Type) = frmSystem.Controls("dc_DefrostType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_make_manufacturer) = frmSystem.Controls("dc_SystemMake1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Metered_Usage) = frmSystem.Controls("dc_SystemMeteredUsage1").Value
End Sub
Private Sub savefreezer()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addfreezer
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writefreezer
        Else
            Call addfreezer
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addfreezer()
        If iFreezer < 3 Then
            iFreezer = iFreezer + 1
            strCurrentSystemName = "FREEZER-" + CStr(iFreezer)
            Call writefreezer
        Else
            MsgBox ("You can only enter at most 3 freezers!")
        End If
End Sub
Private Sub writefreezer()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "FREEZER"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size) = frmSystem.Controls("dc_SystemSize1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure) = frmSystem.Controls("dc_SizeUnit1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age) = frmSystem.Controls("dc_SystemAge1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Defrost_Type) = frmSystem.Controls("dc_DefrostType1").Value
End Sub
Private Sub saveappliance()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addappliance
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writeappliance
        Else
            Call addappliance
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addappliance()
        If iAppliance < 27 Then
            iAppliance = iAppliance + 1
            strCurrentSystemName = "APPLIANCE-" + CStr(iAppliance)
            Call writeappliance
        Else
            MsgBox ("You can only enter at most 27 appliance!")
        End If
End Sub
Private Sub writeappliance()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "APPLIANCE"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity) = frmSystem.Controls("dc_Quantity1").Value
End Sub
Private Sub savehvacdistribution()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addhvac
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writehvac
        Else
            Call addhvac
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addhvac()
        If iHVAC < 6 Then
            iHVAC = iHVAC + 1
            strCurrentSystemName = "HVAC DISTRIBUTION-" + CStr(iHVAC)
            Call writehvac
        Else
            MsgBox ("You can only enter at most 6 HVAC distribution systems!")
        End If
End Sub
Private Sub writehvac()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "HVAC DISTRIBUTION"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size) = frmSystem.Controls("dc_SystemSize1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_exist_indicator) = frmSystem.Controls("dc_InsIndicator1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Insulation_Type) = frmSystem.Controls("dc_InsType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Location) = frmSystem.Controls("dc_SystemLocation1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Length) = frmSystem.Controls("dc_SystemLength1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Condition_of_flex_duct) = frmSystem.Controls("dc_FlexCondition1").Value
End Sub
Private Sub saveheatingsystem()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addheating
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writeheating
        Else
            Call addheating
        End If
    End If
    Call updatelistbox
End Sub
Private Sub addheating()
        If iHeating < 6 Then
            iHeating = iHeating + 1
            strCurrentSystemName = "HEATING-" + CStr(iHeating)
            Call writeheating
        Else
            MsgBox ("You can only enter at most 6 heating systems!")
        End If
End Sub
Private Sub writeheating()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "HEATING"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Fuel_Source) = frmSystem.Controls("dc_FuelSource1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size) = frmSystem.Controls("dc_SystemSize1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure) = frmSystem.Controls("dc_SizeUnit1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age) = frmSystem.Controls("dc_SystemAge1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Efficiency_Rating) = frmSystem.Controls("dc_EffRating1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Efficiency_Rating_Type) = frmSystem.Controls("dc_EffRatingType1").Value
End Sub
Private Sub addcooling()
        If iCooling < 6 Then
            iCooling = iCooling + 1
            strCurrentSystemName = "COOLING-" + CStr(iCooling)
            Call writecooling
        Else
            MsgBox ("You can only enter at most 6 cooling systems!")
        End If
End Sub
Private Sub writecooling()
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Record_Type) = strCurrentSystemName
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Name) = "COOLING"
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Not_Applicable) = frmSystem.Controls("dc_SystemApplicable1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Type) = frmSystem.Controls("dc_SystemType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Fuel_Source) = frmSystem.Controls("dc_FuelSource1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Size) = frmSystem.Controls("dc_SystemSize1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Size_Unit_of_Measure) = frmSystem.Controls("dc_SizeUnit1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.System_Age) = frmSystem.Controls("dc_SystemAge1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Efficiency_Rating) = frmSystem.Controls("dc_EffRating1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Efficiency_Rating_Type) = frmSystem.Controls("dc_EffRatingType1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Percent_of_space_heated_or_cooled) = frmSystem.Controls("dc_PercentageCooled1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Frequency_of_system_use) = frmSystem.Controls("dc_FrequencyUse1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Total_units_used) = frmSystem.Controls("dc_TotalUnits1").Value
        Worksheets(AuditSheetName).Cells(auditcurrentrow, LGEContextual.Quantity) = frmSystem.Controls("dc_Quantity1").Value
End Sub
Private Sub savecoolingsystem()
    If lstSelectedSystems.ListIndex = -1 Then
        Call addcooling
    Else
        pos = InStr(1, lstSelectedSystems.Text, "-")
        lastSystemType = Mid(lstSelectedSystems.Text, 1, pos - 1)
        If lastSystemType = cboSystem.Text Then
            Call writecooling
        Else
            Call addcooling
        End If
    End If
    Call updatelistbox
End Sub

Private Sub cmdRemove_Click()
    currentrow = lstSelectedSystems.ListIndex
    Select Case Worksheets(AuditSheetName).Cells(currentrow + 2, 5).Value
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
    strCurrentSystemName = ""
    lstSelectedSystems.RemoveItem (currentrow)
    Worksheets(AuditSheetName).Rows(currentrow + 2).Delete
    Call updatelistbox
End Sub

Private Sub cmdRemoveAll_Click()
    auditlastrow = Worksheets(AuditSheetName).Range("E" & Rows.Count).End(xlUp).Row
    cboSystem.Text = ""
    strCurrentSystemName = ""
    Worksheets(AuditSheetName).Range("A2:AZ" & auditlastrow).Clear
    lstSelectedSystems.Clear
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
    iAppliance = 0
    Call updatelistbox
End Sub

Private Sub cmdRename_Click()
    Dim strSystem As String
    If lstSelectedSystems.ListIndex = -1 Then
        iReply = MsgBox("Please select the system to rename", vbOKOnly, "Select a system!")
        Exit Sub
    End If
        
    currentrow = lstSelectedSystems.ListIndex
    strSystem = Worksheets(AuditSheetName).Cells(currentrow + 2, 5).Value
    
    Dim message, title, defaultValue As String
    Dim myValue As String

    message = "Enter the system name"
    title = "System Name"
    defaultValue = "my favoriate system"
    myValue = InputBox(message, title, defaultValue)
    If myValue = "" Then myValue = defaultValue

    strCurrentSystemName = strSystem + "-" + myValue
    Worksheets(AuditSheetName).Cells(currentrow + 2, 1).Value = strCurrentSystemName
    
    auditlastrow = Worksheets(AuditSheetName).Range("E" & Rows.Count).End(xlUp).Row
    lstSelectedSystems.Clear
    If auditlastrow > 1 Then
        For i = 2 To auditlastrow
            lstSelectedSystems.AddItem (Worksheets(AuditSheetName).Cells(i, 1))
        Next i
    End If
    Call updatelistbox
End Sub

Private Sub lstSelectedSystems_Change()
    If lstSelectedSystems.ListIndex <> -1 Then
        cmdRemove.Enabled = True
    End If
    auditcurrentrow = CInt(lstSelectedSystems.ListIndex + 2)
End Sub

Private Sub UserForm_Activate()
    vertInterval = 25
    toTop0 = 60 ' not applicable
    toTop = 85
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
    toLeft2 = 180
    toLeft3 = 250
    
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
    
    Call updatelistbox
'    auditlastrow = Worksheets(AuditSheetName).Range("E" & Rows.Count).End(xlUp).Row
'
'    If auditlastrow > 1 Then
'        For i = 2 To auditlastrow
'            lstSelectedSystems.AddItem (Worksheets(AuditSheetName).Cells(i, 1))
'            Select Case Worksheets(AuditSheetName).Cells(i, 5)
'                Case "HEATING"
'                    iHeating = iHeating + 1
'                Case "COOLING"
'                    iCooling = iCooling + 1
'                Case "HVAC DISTRIBUTION"
'                    iHVAC = iHVAC + 1
'                Case "WATER HEATER"
'                    iWH = iWH + 1
'                Case "THERMOSTAT"
'                    iThermostat = iThermostat + 1
'                Case "WINDOW"
'                    iWindow = iWindow + 1
'                Case "DOOR"
'                    iDoor = iDoor + 1
'                Case "LIGHTING"
'                    iLighting = iLighting + 1
'                Case "WALL"
'                    iWall = iWall + 1
'                Case "ATTIC"
'                    iAttic = iAttic + 1
'                Case "BASEMENT"
'                    iBasement = iBasement + 1
'                Case "BASEMENT WALL"
'                    iBW = iBW + 1
'                Case "REFRIGERATOR"
'                    iRefrigerator = iRefrigerator + 1
'                Case "FREEZER"
'                    iFreezer = iFreezer + 1
'                Case "APPLIANCE"
'                    iAppliance = iAppliance + 1
'            End Select
'        Next i
'    End If

    If lstSelectedSystems.ListIndex = -1 Then
        cmdRemove.Enabled = False
    End If
    
    txtEnrollmentID.Text = currentEnrollment
    txtPremiseID.Text = premiseid
    txtAccountNumber.Text = accountnumber
    txtEnrollmentID.Enabled = False
    txtPremiseID.Enabled = False
    txtAccountNumber.Enabled = False
    
    'Application.Visible = False
    
'    applianceStartCol = NexantEnrollments.APPLIANCE_AQUARIUM_quantity
'    applianceNum = NexantEnrollments.ATTIC_1_access_type - NexantEnrollments.APPLIANCE_AQUARIUM_quantity
'    applianceLimit = 1
'    atticStartCol = NexantEnrollments.ATTIC_1_access_type
'    atticNum = NexantEnrollments.ATTIC_1_vent - NexantEnrollments.ATTIC_1_access_type + 1
'    atticLimit = 4
'    basementStartCol = NexantEnrollments.BASEMENT_1_air_conditioned
'    basementNum = NexantEnrollments.BASEMENT_1_type - NexantEnrollments.BASEMENT_1_air_conditioned + 1
'    basementLimit = 3
'    basementwallStartCol = NexantEnrollments.BASEMENT_WALL_1_insulation
'    basementwallNum = NexantEnrollments.BASEMENT_WALL_1_type - NexantEnrollments.BASEMENT_WALL_1_insulation + 1
'    basementwallLimit = 3
'    coolingStartCol = NexantEnrollments.COOLING_1_age
'    coolingNum = NexantEnrollments.COOLING_1_usage_frequency - NexantEnrollments.COOLING_1_age + 1
'    coolingLimit = 6
'    doorStartCol = NexantEnrollments.DOOR_1_condition
'    doorNum = NexantEnrollments.DOOR_1_type - NexantEnrollments.DOOR_1_condition + 1
'    doorLimit = 5
'    freezerStartCol = NexantEnrollments.FREEZER_1_age
'    freezerNum = NexantEnrollments.FREEZER_1_type - NexantEnrollments.FREEZER_1_age + 1
'    freezerLimit = 3
'    heatingStartCol = NexantEnrollments.HEATING_1_age
'    heatingNum = NexantEnrollments.HEATING_1_type - NexantEnrollments.HEATING_1_age + 1
'    heatingLimit = 6
'    hvacdistStartCol = NexantEnrollments.HVAC_DIST_1_flex_duct_condition
'    hvacdistNum = NexantEnrollments.HVAC_DIST_1_type - NexantEnrollments.HVAC_DIST_1_flex_duct_condition + 1
'    hvacdistLimit = 6
'    lightingStartCol = NexantEnrollments.LIGHTING_1_not_applicable
'    lightingNum = NexantEnrollments.LIGHTING_1_weekly_hrs - NexantEnrollments.LIGHTING_1_not_applicable + 1
'    lightingLimit = 4
'    refrigStartCol = NexantEnrollments.REFRIGERATOR_1_age
'    refrigNum = NexantEnrollments.REFRIGERATOR_1_type - NexantEnrollments.REFRIGERATOR_1_age + 1
'    refrigLimit = 3
'    thermostatStartCol = NexantEnrollments.THERMOSTAT_1_ac_load_control
'    thermostatNum = NexantEnrollments.THERMOSTAT_1_type - NexantEnrollments.THERMOSTAT_1_ac_load_control + 1
'    thermostatLimit = 3
'    wallStartCol = NexantEnrollments.WALL_1_height
'    wallNum = NexantEnrollments.WALL_1_type - NexantEnrollments.WALL_1_height + 1
'    wallLimit = 4
'    waterheaterStartCol = NexantEnrollments.WATER_HEATER_1_age
'    waterheaterNum = NexantEnrollments.WATER_HEATER_1_type - NexantEnrollments.WATER_HEATER_1_age + 1
'    waterheaterLimit = 3
'
'    sysnum = Array(applianceNum, atticNum, basementNum, basementwallNum, coolingNum, doorNum, freezerNum, heatingNum, hvacdistNum, lightingNum, refrigNum, thermostatNum, wallNum, waterheaterNum)
'    syslimit = Array(applianceLimit, atticLimit, basementLimit, basementwallLimit, coolingLimit, doorLimit, freezerLimit, heatingLimit, hvacdistLimit, lightingLimit, refrigLimit, thermostatLimit, wallLimit, waterheaterLimit)
'    Call updatesystem(currentrow)
End Sub


Private Sub updatelistbox()
    auditlastrow = Worksheets(AuditSheetName).Range("E" & Rows.Count).End(xlUp).Row
    auditcurrentrow = auditlastrow + 1
'    If lstSelectedSystems.ListIndex <> -1 Then
'        auditcurrentrow = lstSelectedSystems.ListIndex + 2
'    End If
    
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
    iAppliance = 0
    
    lstSelectedSystems.Clear
    
    If auditlastrow > 1 Then
        For i = 2 To auditlastrow
            lstSelectedSystems.AddItem (Worksheets(AuditSheetName).Cells(i, 1))
            Select Case Worksheets(AuditSheetName).Cells(i, 5)
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
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
End Sub
