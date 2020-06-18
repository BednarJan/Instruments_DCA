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

    'Enum RectifierMode
    '    RMS = 0
    '    VMEAN = 1
    '    modedc = 2
    '    modeac = 3
    '    fnd = 4
    'End Enum

    Enum RectifierMode
        ACDC = 0
        MN = 1
        AC = 2
        DC = 3
        FND = 4
        RMS = 5
    End Enum



    Enum PA_Function
        Voltage = 1
        Current = 2
        ActivePower = 3
        ApparentPower = 4
        ReactivPower = 5
        PF = 6
        PhaseDiff = 7
        FrequencyU = 8
        FrequencyI = 9
        VoltPeakPlus = 10
        VoltPeakMinus = 11
        CurrentPeakPlus = 12
        CurrentPeakMinus = 13
        PowerPeakPlus = 14
        PowerPeakMinus = 15

        PFtot = 16
        THDvolt = 17
        THDCurr = 18
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
    Function QueryNumericItems() As Single()

    Sub SetDisplayItem(nFn As IPWAN.PA_Function, disp As IPWAN.PA_Display, nRectMode As IPWAN.RectifierMode, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)

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