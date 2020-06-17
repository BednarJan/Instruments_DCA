Public Interface IPWAN
    Inherits IDevice


#Region "Enum definitions"
    Enum Wiring
        P1W3 = 1
        P3W3 = 2
        P3W4 = 3
        V3A3 = 4
        P1W2x3 = 5
        P1W3AndP1W2 = 6
        P3W3AndP1W2 = 7
        P3W3M2 = 8
        P3W3M3 = 9
    End Enum

    Enum THDNorm
        IEC = 1
        CSA = 2
    End Enum

    Enum Items
        Item1 = 1
        Item2 = 2
        Item3 = 3
        Item4 = 4
    End Enum

    Enum Elements
        Element1 = 1
        Element2 = 2
        Element3 = 3
        Sigma = 4
    End Enum

    Enum RectifierMode
        RMS = 0
        VMEAN = 1
        modedc = 2
        modeac = 3
        fnd = 4
    End Enum

    Enum PA_Function
        voltage = 1
        Current = 2
        Active_power = 3
        Apparent_power = 4
        Reactive_power = 5
        PF = 6
        Deg = 7
        Freq_Voltage = 8
        Freq_Current = 9
        Voltage_Peak_Plus = 10
        Voltage_Peak_Minus = 11
        Current_Peak_Plus = 12
        Current_Peak_Minus = 13
        Active_power_Peak_Plus = 14
        Active_power_Peak_Minus = 15

        PF_Total = 16
        THD_Voltage = 17
        THD_Current = 18
        IntegratedActivePower = 19
        IntegratedCurrent = 20
    End Enum

    Enum PA_Attributes
        U = 1
        I = 2
        P = 3
        S = 4
        Q = 5
        PF = 6
        PHI = 7
        FU = 8
        FI = 9
        THDU = 10
        THDI = 11
        Items = 16
    End Enum


    Enum PA_Display
        a = 1
        b = 2
        c = 3
        d = 4
    End Enum

#End Region

#Region "Properties"
    ReadOnly Property VoltageMax As Single
    ReadOnly Property CurrentMax As Single
    ReadOnly Property InputElements As UInteger
#End Region

#Region "Methods (Sub & Functions)"
    Sub SetWiring(iWir As Wiring)
    Sub SetRectifierMode(iMode As RectifierMode, Optional elm As Elements = Elements.Sigma)
    Function GetVrms(Optional elm As Elements = Elements.Element1) As Single
    Function GetVPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single
    Function GetVPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single
    Function GetIrms(Optional elm As Elements = Elements.Element1) As Single
    Function GetIPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single
    Function GetIPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single

    Function GetPactive(Optional elm As Elements = Elements.Element1) As Single
    Function GetPapparent(Optional elm As Elements = Elements.Element1) As Single
    Function GetPreact(Optional elm As Elements = Elements.Element1) As Single

    Function GetUTHD(Optional elm As Elements = Elements.Element1) As Single
    Function GetITHD(Optional elm As Elements = Elements.Element1) As Single
    Function GetPF(Optional elm As Elements = Elements.Element1) As Single

    Function GetFreqU(Optional elm As Elements = Elements.Element1) As Single
    Function GetFreqI(Optional elm As Elements = Elements.Element1) As Single

    Function GetHarmonicsU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single()
    Function GetHarmonicsI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single()

    Sub ClearNumericItems()
    Sub SetNumericItemsCount(nCount As Integer)
    Sub SetNumericItem(nFn As IPWAN.PA_Function, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional itm As Integer = 1, Optional ordHarm As Integer = 0)
    Sub QueryNumericItems(ByRef oVals() As Single)

    Sub SetDisplayItem(nFn As IPWAN.PA_Function, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional disp As IPWAN.PA_Display = IPWAN.PA_Display.a)

    Sub SetInputVoltageRange(Optional nRangeInVolts As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)
    Sub SetInputCurrentRange(Optional nRangeInAmps As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)

    Sub PresetCurrentTransformer(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)
    Sub PresetVoltDivider(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)
    Sub PresetCurrentShunt(shuntRes As Single, sRange As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)

    Sub StartIntegration()
    Sub StopIntegration()

    Function GetIntegratedPower(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single
    Function GetIntegratedCurrent(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single


    'Function GetPF(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    'Function GetVrms(Optional ByVal WaitOnPWAN As Boolean = False) As Single
    'Function GetVthd(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    'Function GetIrms(Optional ByVal WaitOnPWAN As Boolean = False) As Single
    'Function GetIthd(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    'Function GetPreal(Optional ByVal WaitOnPWAN As Boolean = False) As Single
    'Function GetPapp(Optional ByVal WaitOnPWAN As Boolean = False) As Single
    'Function GetPreact(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    'Function GetFreq(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    'Sub GetHarmonics_IEC61000_3_2(EquipClass As CPWAN_Helper.EEquipmentClass, ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, Optional ByVal WaitOnPWAN As Boolean = False)
    'Sub GetHarmonics_IEC61000_3_12(ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, ByRef THD As Single, ByRef PWHD As Single, Optional ByVal WaitOnPWAN As Boolean = False)
    'Function IntegratePower(durationInSec As Single) As Double
    'Sub Init_Integration(durationInSec As Single)
    'Sub Get_Integrations_Data(ByRef measData() As Double)
    'Sub Set_Volt_Range(sRange As String)
    'Sub Set_Curr_Range(sRange As String)

#End Region
End Interface