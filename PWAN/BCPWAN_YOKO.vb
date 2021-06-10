'Base clase  BCPWAN common for all Power Analysers,  inherited fro the BCDevice
'17.06.2020, JaBe

Imports Ivi.Visa
Public Class BCPWAN_YOKO
    Inherits BCPWAN
    Implements IPWAN

    Const HarmonicStartPos As Integer = 0

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        Visa.SendString("*RST;*CLS")
        Visa.SendString(":NUMERIC:FORMAT ASCII")
        Visa.SendString(":SYSTEM:COMMUNICATE:COMMAND WT300")
        Visa.SendString(":HARMONICS:THD TOTAL")
        Visa.SendString(":RATE 500MS")
        Visa.SendString(":INPUT:VOLTAGE:AUTO ON")
        Visa.SendString(":INPUT:CURRENT:AUTO ON")

        Visa.SendString(":DISPLAY:NORMAL:ITEM1 U")
        Visa.SendString(":DISPLAY:NORMAL:ITEM2 I")
        Visa.SendString(":DISPLAY:NORMAL:ITEM3 P")
        Visa.SendString(":DISPLAY:NORMAL:ITEM4 PF")

        ClearNumericItems()

        Dim nItemsCount As Integer = CreateNumericNormalItemsList()
        SetNumericItemsCount(nItemsCount)

    End Sub
#End Region

#Region "Interface Methodes IPWAN"
    Public Overrides Sub SetWiring(iWir As IPWAN.Wiring) Implements IPWAN.SetWiring
        Dim cmdStr As String = ":INPUT:WIRING"
        Select Case iWir
            Case IPWAN.Wiring.oneP2Wx3, IPWAN.Wiring.oneP2W
                cmdStr &= " P1W2"
            Case IPWAN.Wiring.oneP3W, IPWAN.Wiring.oneP3W_1P2W
                cmdStr &= " P1W3"
            Case IPWAN.Wiring.threeP3W
                cmdStr &= " P3W3"
            Case IPWAN.Wiring.threeP4W
                cmdStr &= " P3W4"
            Case IPWAN.Wiring.threeV3A
                cmdStr &= " V3A3"
        End Select

        Visa.SendString(cmdStr)

    End Sub

    Public Overrides Sub SetTHDNorm(nTHDNorm As IPWAN.THDNorm) Implements IPWAN.SetTHDNorm

        Select Case nTHDNorm
            Case IPWAN.THDNorm.CSA
                Visa.SendString(":HARMONICS:THD TOT")         'CSA norm
            Case IPWAN.THDNorm.IEC
                Visa.SendString(":HARMONICS:THD FUND")         'CSA norm
        End Select

    End Sub

    '''SetInputMode(iMode As IPWAN.RectifierMode, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional fn As IPWAN.PA_Function = IPWAN.PA_Function.Voltage) Implements IPWAN.SetInputMode
    '''Set rectifier mode for all input elements and measurements. Unable set it separatelly for the input channel
    '''
    Public Overrides Sub SetInputMode(iMode As IPWAN.RectifierMode) Implements IPWAN.SetInputMode

        Dim strMode As String = GetRectifierMode(iMode)
        Visa.SendString(":INPUT:MODE " & strMode)

    End Sub

    Public Overrides Function QueryNumericItems() As Double() Implements IPWAN.QueryNumericItems

        If Not IsEmptyHarmonicsList Then
            HarmonicsList.Clear()
        End If

        If IsEmptyNumericNormaltemsList Then
            CreateNumericNormalItemsList()
        End If

        'Dim dummy() As Double = QueryValueList(":NUMERIC:NORMAL:VALUE?;*WAI?")
        'cHelper.Delay(0.5)
        Return QueryValueList(":NUMERIC:NORMAL:VALUE?;*WAI?")

    End Function

    Public Overrides Sub ClearNumericItems() Implements IPWAN.ClearNumericItems

        Visa.SendString(":NUMERIC:NORMAL:CLEAR ALL")
        Visa.SendString(":NUMERIC:NORMAL:NUMBER 0")

    End Sub

    Public Overrides Sub SetNumericItem(_NNItem As CNumericNormalItem) Implements IPWAN.SetNumericItem
        Dim cmdStr As String

        cmdStr = ":NUMERIC:NORMAL:ITEM"
        cmdStr &= _NNItem.Itm.ToString & " "
        cmdStr &= _NNItem.Fn & ","
        cmdStr &= _NNItem.Elm

        If _NNItem.OrderHarm <> String.Empty Then

            cmdStr &= "," & _NNItem.OrderHarm
        End If

        Visa.SendString(cmdStr)

    End Sub

    Overrides Function CreateNumericNormalItemsList() As Integer Implements IPWAN.CreateNumericNormalItemsList

        Return MyBase.CreateNumericNormalItemsList()

    End Function


    Public Overrides Function GetVrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVrms

        Return QueryNumericItem(IPWAN.PA_Function.Voltage, elm)

    End Function

    Public Overrides Function GetVPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakPlus

        Return QueryNumericItem(IPWAN.PA_Function.VoltPeakPlus, elm)

    End Function

    Public Overrides Function GetVPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakMinus

        Return QueryNumericItem(IPWAN.PA_Function.VoltPeakMinus, elm)

    End Function

    Public Overrides Function GetIrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIrms

        Return QueryNumericItem(IPWAN.PA_Function.Current, elm)

    End Function

    Public Overrides Function GetIPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakPlus

        Return QueryNumericItem(IPWAN.PA_Function.CurrentPeakPlus, elm)

    End Function

    Public Overrides Function GetIPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakMinus

        Return QueryNumericItem(IPWAN.PA_Function.CurrentPeakMinus, elm)

    End Function

    Public Overrides Function GetPactive(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPactive

        Return QueryNumericItem(IPWAN.PA_Function.ActivePower, elm)

    End Function

    Public Overrides Function GetPapparent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPapparent

        Return QueryNumericItem(IPWAN.PA_Function.ApparentPower, elm)

    End Function

    Public Overrides Function GetPreact(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPreact

        Return QueryNumericItem(IPWAN.PA_Function.ReactivPower, elm)

    End Function

    Public Overrides Function GetUTHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetUTHD

        Dim harm() As Double = GetHarmonicsU(elm)

        'very last item is always THD
        Return harm(harm.Length - 1)

    End Function

    Public Overrides Function GetITHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetITHD

        Dim harm() As Double = GetHarmonicsI(elm)

        'very last item is always THD
        Return harm(harm.Length - 1)

    End Function

    Public Overrides Function GetFreqU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqU

        Return QueryNumericItem(IPWAN.PA_Function.FrequencyU, elm)

    End Function

    Public Overrides Function GetFreqI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqI

        Return QueryNumericItem(IPWAN.PA_Function.FrequencyI, elm)

    End Function

    Public Overrides Function GetPF(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPF

        Return QueryNumericItem(IPWAN.PA_Function.PF, elm)

    End Function

    Public Overrides Function GetHarmonicsU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Double() Implements IPWAN.GetHarmonicsU

        If IsEmptyHarmonicsList Then

            PresetHarmonics("UK", elm)

        End If

        Return QueryHarmonicItems()

    End Function

    Public Overrides Function GetHarmonicsI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Double() Implements IPWAN.GetHarmonicsI

        If Not IsEmptyNumericNormaltemsList OrElse IsEmptyHarmonicsList Then

            PresetHarmonics("IK", elm)

        End If

        Return QueryHarmonicItems()

    End Function


    Public Overrides Sub SetInputVoltageRange(Optional nRangeInVolts As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetInputVoltageRange
        If nRangeInVolts = 0 Then
            Visa.SendString(":INPUT:VOLTAGE:AUTO ON")
        Else
            Visa.SendString(":INPUT:VOLTAGE:RANGE " & FormatNumber(nRangeInVolts, 2) & " V")
        End If
    End Sub

    Public Overrides Sub SetInputCurrentRange(Optional nRangeInAmps As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetInputCurrentRange
        If nRangeInAmps = 0 Then
            Visa.SendString(":INPUT:CURRENT:AUTO ON")
        Else
            Visa.SendString(":INPUT:CURRENT:RANGE " & FormatNumber(nRangeInAmps, 2) & " A")
        End If
    End Sub


    Public Overrides Sub SetDisplayItem(nFn As IPWAN.PA_Function, disp As IPWAN.PA_Display, nRectMode As IPWAN.RectifierMode, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetDisplayItem

        Dim sELM As String = GetElement(elm)
        Dim sFN As String = GetFunction(nFn)

        Visa.SendString(":DISPLAY:NORMAL:ITEM" & disp & " " & sFN & "," & sELM)

    End Sub

    Public Overrides Sub PresetCurrentProbe(sRatioInMamps As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentProbe

        Visa.SendString(":INPUT:SCALING:STATE OFF")
        Visa.SendString(":INPUT:RCONFIG ON")
        Visa.SendString(":INPUT:SCALING:CT:ELEMENT" & elm & " " & FormatNumber(sRatioInMamps / 1000))
        Visa.SendString(":INPUT:SCALING:STATE ON")

    End Sub


    Public Overrides Sub PresetCurrentTransformer(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentTransformer

        Visa.SendString(":INPUT:SCALING:STATE OFF")
        Visa.SendString(":INPUT:RCONFIG ON")
        Visa.SendString(":INPUT:SCALING:CT:ELEMENT" & elm & " " & FormatNumber(sRatio))
        Visa.SendString(":INPUT:SCALING:STATE ON")

    End Sub

    Public Overrides Sub PresetVoltDivider(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetVoltDivider

        Visa.SendString(":INPUT:SCALING:STATE OFF")
        Visa.SendString(":INPUT:SCALING:VT:ELEMENT" & elm & " " & FormatNumber(sRatio))
        Visa.SendString(":INPUT:SCALING:STATE ON")

    End Sub

    Public Overrides Sub PresetCurrentShunt(resMiliOhms As Single, sRange As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentShunt

        Visa.SendString(":INPUT:CURRENT:RANGE EXTERNAL," & FormatNumber(sRange) & "V")
        Visa.SendString(":INPUT:CURRENT:SRATIO:ELEMENT" & elm & " " & FormatNumber(resMiliOhms))

    End Sub

    Public Overrides Sub PresetIntegrationPower(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetIntegrationPower

        Call Visa.SendString("DISPLAY" & CStr(IPWAN.PA_Display.a) & ":FUNCtion " & "Time")
        Call Visa.SendString("DISPLAY" & CStr(IPWAN.PA_Display.b) & ":FUNCtion " & "W")
        Call Visa.SendString("DISPLAY" & CStr(IPWAN.PA_Display.c) & ":FUNCtion " & "Wh")

        Call Visa.SendString("INTEGrate:RESet")
        Call Visa.SendString("INTEGrate:MODE NORMAL")
        Call Visa.SendString("INTEGRATE:TYPE STANdard")

    End Sub

    Public Overrides Sub PresetIntegrationCurrent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetIntegrationCurrent

        Call Visa.SendString("DISPLAY" & CStr(IPWAN.PA_Display.a) & ":FUNCtion " & "Time")
        Call Visa.SendString("DISPLAY" & CStr(IPWAN.PA_Display.b) & ":FUNCtion " & "A")
        Call Visa.SendString("DISPLAY" & CStr(IPWAN.PA_Display.c) & ":FUNCtion " & "Ah")

        Call Visa.SendString("INTEGrate:RESet")
        Call Visa.SendString("INTEGrate:MODE NORMAL")
        Call Visa.SendString("INTEGRATE:TYPE STANdard")

    End Sub

    Public Overrides Sub StartIntegration() Implements IPWAN.StartIntegration

        Visa.SendString("INTEGrate:RESet")
        Visa.SendString("INTEGrate:STARt")

    End Sub

    Public Overrides Sub StopIntegration() Implements IPWAN.StopIntegration

        Visa.SendString("INTEGrate:STop")

    End Sub

    Public Overrides Function GetIntegratedPower(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIntegratedPower
        Dim dRet As Double = Double.NaN
        Dim measData() As Double = GetIntegrationsData()

        If durationInSec > 0 Then
            dRet = (measData(1) * 3600) / durationInSec
        End If

        Return dRet

    End Function

    Public Overrides Function GetIntegratedCurrent(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIntegratedCurrent
        Dim dRet As Double = Double.NaN
        Dim measData() As Double = GetIntegrationsData()

        If durationInSec > 0 Then
            dRet = (measData(1) * 3600) / durationInSec
        End If

        Return dRet

    End Function


#End Region

#Region "Help functions"

    Overrides Function GetFunctionIndex(nFn As IPWAN.PA_Function, nElm As Integer) As Integer

        Return 3 * (CInt(nFn) - 1) + nElm - 1

    End Function

    Overrides Function GetFunction(nFn As IPWAN.PA_Function) As String

        Return MyBase.GetFunction(nFn)

    End Function

    Public Overrides Sub SetNumericItemsCount(nCount As Integer)

        Call Visa.SendString(":NUMERIC:NORMAL:NUMBER " & nCount)

    End Sub

    Public Overrides Function GetNumericItemsCount() As Integer

        Dim ErrorMsg As String = vbNullString

        Call Visa.SendString(":NUMERIC:NORMAL:NUMBER?")
        Dim nCount As Integer = Visa.ReceiveValue(ErrorMsg)

        Return CInt(nCount)

    End Function


    Private Sub PresetHarmonics(Fn As String)

        PresetHarmonics(Fn, IPWAN.Elements.Element1)

    End Sub

    Private Sub PresetHarmonics(Fn As String, Elm As IPWAN.Elements)

        Dim _NNItem As CNumericNormalItem

        Call ClearNumericItems()
        NumericNormaltemsList.Clear()

        HarmonicsList = New List(Of CNumericNormalItem)

        Call SetNumericItemsCount(HarmCount + 3)


        Dim itm As Integer = 1

        _NNItem = New CNumericNormalItem(Fn, HarmonicStartPos + itm, Elm, "TOTAL")
        HarmonicsList.Add(_NNItem)
        itm += 1

        _NNItem = New CNumericNormalItem(Fn, HarmonicStartPos + itm, Elm, "DC")
        HarmonicsList.Add(_NNItem)
        itm += 1

        Do While itm < HarmCount + 3

            _NNItem = New CNumericNormalItem(Fn, HarmonicStartPos + itm, Elm, (itm - HarmonicStartPos - 2).ToString)
            HarmonicsList.Add(_NNItem)

            itm += 1
        Loop

        'THD item as suffix

        Select Case Fn
            Case "UK"
                _NNItem = New CNumericNormalItem("UHDFk", HarmonicStartPos + itm, Elm, itm.ToString)
            Case "IK"
                _NNItem = New CNumericNormalItem("IHDFk", HarmonicStartPos + itm, Elm, itm.ToString)
        End Select
        HarmonicsList.Add(_NNItem)

        For Each _NNItem In HarmonicsList
            Call SetNumericItem(_NNItem)
            cHelper.Delay(0.2)
        Next

    End Sub

    Overrides Function QueryHarmonicItems()

        Return QueryValueList(":NUMERIC:NORMAL:VALUE?")

    End Function

    Overrides Function GetElement(nElm As Integer) As String
        Dim sELM As String
        If nElm <> IPWAN.Elements.Element1 Then nElm = IPWAN.Elements.Element1

        If nElm = IPWAN.Elements.Sigma Then
            sELM = "SIGMA"
        Else
            sELM = nElm
        End If
        Return sELM
    End Function

    Overrides Function GetIntegrationsData() As Double()

        Visa.SendString("MEASure:NORMal:Item:PRESet INTEGrate")

        Return QueryValueList("Measure: Value?")

    End Function

    Private Function QueryValueList(ByVal cmdStr As String) As Double()

        Visa.SendString(cmdStr)
        Return Visa.ReceiveValueList(",")

    End Function

    Private Function QueryNumericItem(fn As IPWAN.PA_Function, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Decimal

        If IsEmptyNumericNormaltemsList Then

            CreateNumericNormalItemsList()

        End If

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(fn, elm))

    End Function


    Overrides Function CreateFunctionList() As SortedList(Of String, String)

        'this is the fix item order within the preset pattern 3 

        Dim fsl As New SortedList(Of String, String) From {
        {IPWAN.PA_Function.Voltage.ToString, "U"},
        {IPWAN.PA_Function.Current.ToString, "I"},
        {IPWAN.PA_Function.ActivePower.ToString, "P"},
        {IPWAN.PA_Function.ApparentPower.ToString, "S"},
        {IPWAN.PA_Function.ReactivPower.ToString, "Q"},
        {IPWAN.PA_Function.PF.ToString, "LAMBDA"},
        {IPWAN.PA_Function.FrequencyU.ToString, "FU"},
        {IPWAN.PA_Function.FrequencyI.ToString, "FI"},
        {IPWAN.PA_Function.VoltPeakPlus.ToString, "UPP"},
        {IPWAN.PA_Function.VoltPeakMinus.ToString, "UMP"},
        {IPWAN.PA_Function.CurrentPeakPlus.ToString, "IPP"},
        {IPWAN.PA_Function.CurrentPeakMinus.ToString, "IMP"},
        {IPWAN.PA_Function.PowerPeakPlus.ToString, "PPP"},
        {IPWAN.PA_Function.PowerPeakMinus.ToString, "PMP"}
        }

        Return fsl

    End Function


#End Region



End Class
