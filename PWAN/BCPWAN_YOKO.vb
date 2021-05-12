'Base clase  BCPWAN common for all Power Analysers,  inherited fro the BCDevice
'17.06.2020, JaBe

Imports Ivi.Visa
Public Class BCPWAN_YOKO
    Inherits BCPWAN
    Implements IPWAN

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        Visa.SendString("*RST;*CLS" & Chr(10))
        Visa.SendString(":NUMERIC:FORMAT ASCII" & Chr(10))
        Visa.SendString(":SYSTEM:COMMUNICATE:COMMAND WT300" & Chr(10))
        Visa.SendString(":HARMONICS:THD TOTAL" & Chr(10))
        Visa.SendString(":RATE 500MS" & Chr(10))
        Visa.SendString(":INPUT:VOLTAGE:AUTO ON" & Chr(10))
        Visa.SendString(":INPUT:CURRENT:AUTO ON" & Chr(10))

        Visa.SendString(":DISPLAY:NORMAL:ITEM1 U" & Chr(10))
        Visa.SendString(":DISPLAY:NORMAL:ITEM2 I" & Chr(10))
        Visa.SendString(":DISPLAY:NORMAL:ITEM3 P" & Chr(10))
        Visa.SendString(":DISPLAY:NORMAL:ITEM4 LAMB" & Chr(10))

        Dim nItemsCount As Integer = CreateNumericItemsList()
        SetNumericItemsCount(nItemsCount)

    End Sub
#End Region

#Region "Interface Methodes IPWAN"
    Public Overrides Sub SetWiring(iWir As IPWAN.Wiring) Implements IPWAN.SetWiring
        Dim cmdStr As String = ":INPUT:WIRING"
        Select Case iWir
            Case IPWAN.Wiring.P1W3
                cmdStr &= " P1W3"
            Case IPWAN.Wiring.P3W3
                cmdStr &= " P3W3"
            Case IPWAN.Wiring.P3W4
                cmdStr &= " P3W4"
            Case IPWAN.Wiring.V3A3
                cmdStr &= " V3A3"
            Case Else
                cmdStr &= " P1W3"
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

        Return QueryValueList(":NUMERIC:NORMAL:VALUE?")

    End Function

    Public Overrides Sub ClearNumericItems() Implements IPWAN.ClearNumericItems

        Visa.SendString(":NUMERIC:NORMAL:CLEAR ALL")

    End Sub

    Public Overrides Sub SetNumericItem(nFn As IPWAN.PA_Function, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional itm As Integer = 1, Optional ordHarm As Integer = 0) Implements IPWAN.SetNumericItem
        Dim cmdStr As String
        Dim sELM As String
        Dim sFN As String

        sELM = GetElement(elm)
        sFN = GetFunction(nFn)

        cmdStr = ":NUMERIC:NORMAL:ITEM" & itm.ToString & " " & sFN & "," & sELM
        If ordHarm > 0 Then
            cmdStr = cmdStr & "," & ordHarm
        End If

        Visa.SendString(cmdStr)

    End Sub

    Overrides Function CreateNumericItemsList() As Integer Implements IPWAN.CreateNumericItemsList

        Return MyBase.CreateNumericItemsList

    End Function


    Public Overrides Function GetVrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVrms

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.Voltage, elm))

    End Function

    Public Overrides Function GetVPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakPlus

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.VoltPeakPlus, elm))

    End Function

    Public Overrides Function GetVPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakMinus

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.VoltPeakMinus, elm))

    End Function

    Public Overrides Function GetIrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIrms

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.Current, elm))

    End Function

    Public Overrides Function GetIPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakPlus

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.CurrentPeakPlus, elm))

    End Function

    Public Overrides Function GetIPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakMinus

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.CurrentPeakMinus, elm))

    End Function

    Public Overrides Function GetPactive(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPactive

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.ActivePower, elm))

    End Function

    Public Overrides Function GetPapparent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPapparent

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.ApparentPower, elm))

    End Function

    Public Overrides Function GetPreact(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPreact

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.ReactivPower, elm))

    End Function

    Public Overrides Function GetUTHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetUTHD

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.THDvolt, elm))

    End Function

    Public Overrides Function GetITHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetITHD

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.THDCurr, elm))

    End Function

    Public Overrides Function GetFreqU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqU

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.FrequencyU, elm))

    End Function

    Public Overrides Function GetFreqI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqI

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.FrequencyI, elm))

    End Function

    Public Overrides Function GetPF(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPF

        Dim oVals() As Double = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.PF, elm))

    End Function

    Public Overrides Function GetHarmonicsU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single() Implements IPWAN.GetHarmonicsU

        Dim harm() As Single = getHarmonics(elm, IPWAN.PA_Function.Voltage)

        Return harm

    End Function

    Public Overrides Function GetHarmonicsI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single() Implements IPWAN.GetHarmonicsI

        Dim harm() As Single = getHarmonics(elm, IPWAN.PA_Function.Current)

        Return harm

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

    Public Overrides Sub PresetCurrentShunt(shuntRes As Single, sRange As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentShunt

        Visa.SendString(":INPUT:CURRENT:RANGE EXTERNAL," & FormatNumber(sRange) & "V")
        Visa.SendString(":INPUT:CURRENT:SRATIO:ELEMENT" & elm & " " & FormatNumber(shuntRes))

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

        Return MyBase.GetFunctionIndex(nFn, nElm)

    End Function


    Overrides Function GetFunction(nFn As IPWAN.PA_Function) As String

        Return MyBase.GetFunction(nFn)

    End Function

    Public Overrides Sub SetNumericItemsCount(nCount As Integer)

        Call Visa.SendString(":NUMERIC:NORMAL:NUMBER " & nCount)

    End Sub

    Overrides Function getHarmonics(elm As Integer, fn As IPWAN.PA_Function)

        Call ClearNumericItems()
        Call SetNumericItemsCount(_HarmCount)

        For i As Integer = 1 To _HarmCount
            Call SetNumericItem(fn, elm, i, i)
        Next i

        Call cHelper.Delay(1)
        Dim harm() As Double = QueryNumericItems()


        Return harm
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
        cHelper.Delay(1)
        Return Visa.ReceiveValueList()

    End Function

    Overrides Function CreateFunctionList() As SortedList

        Return MyBase.CreateFunctionList

    End Function



#End Region



End Class
