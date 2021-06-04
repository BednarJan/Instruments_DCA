Public Interface IPWAN
    Inherits IDevice


#Region "Enum definitions"
    Enum Wiring
        oneP2Wx3 = 1
        oneP3W_1P2W = 2
        threP3W_P12W = 3
        threeP3W2M = 4
        threeV3A = 5
        threeP3W3M = 6
        threeP4W = 7
        oneP3W = 8
        threeP3W = 9
        oneP2W = 10

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
        ACDC = 0
        MN = 1
        AC = 2
        DC = 3
        FND = 4
        RMS = 5
    End Enum

    Enum PA_Function
        Voltage = 1
        VoltPeakPlus = 2
        VoltPeakMinus = 3
        Current = 4
        CurrentPeakPlus = 5
        CurrentPeakMinus = 6
        ActivePower = 7
        PowerPeakPlus = 8
        PowerPeakMinus = 9
        ApparentPower = 10
        ReactivPower = 11
        PF = 12
        FrequencyU = 13
        FrequencyI = 14
        'THDvolt = 16
        'THDCurr = 17
        'IntegratedActivePower = 18
        'IntegratedCurrent = 19
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

    Function GetHarmonicsU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Double()
    Function GetHarmonicsI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Double()

    Function GetFunctionIndex(nFn As IPWAN.PA_Function, nElm As Integer) As Integer


    Sub ClearNumericItems()
    Sub SetNumericItem(_NNItem As CNumericNormalItem)

    Function CreateNumericNormalItemsList() As Integer

    Sub SetTHDNorm(nTHDNorm As IPWAN.THDNorm)
    Sub SetInputMode(iMode As IPWAN.RectifierMode)

    Function QueryNumericItems() As Double()

    Sub SetDisplayItem(nFn As IPWAN.PA_Function, disp As IPWAN.PA_Display, nRectMode As IPWAN.RectifierMode, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)

    Sub SetInputVoltageRange(Optional nRangeInVolts As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)
    Sub SetInputCurrentRange(Optional nRangeInAmps As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)

    Sub PresetCurrentProbe(sRatioInMamps As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)
    Sub PresetCurrentTransformer(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)
    Sub PresetVoltDivider(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)
    Sub PresetCurrentShunt(shuntRes As Single, sRange As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)

    Sub PresetIntegrationPower(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)
    Sub PresetIntegrationCurrent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1)

    Sub StartIntegration()
    Sub StopIntegration()

    Function GetIntegratedPower(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single
    Function GetIntegratedCurrent(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single

#End Region
End Interface