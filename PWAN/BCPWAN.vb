'Base clase  BCPWAN common for all Power Analysers,  inherited fro the BCDevice
'17.06.2020, JaBe

Imports Ivi.Visa
Public Class BCPWAN
    Inherits BCDevice
    Implements IPWAN

#Region "Shorthand Properties"
    Public Property VoltageMax As Single Implements IPWAN.VoltageMax
    Public Property CurrentMax As Single Implements IPWAN.CurrentMax
    Public Property InputElements As UInteger Implements IPWAN.InputElements
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Function IDN() As String Implements IDevice.IDN
        Return MyBase.IDN()
    End Function

    Public Overrides Sub RST() Implements IDevice.RST
        MyBase.RST()
    End Sub

    Public Overrides Sub CLS() Implements IDevice.CLS
        MyBase.CLS()
    End Sub

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        Dim ItemDef() As String = {"U", "I", "P", "S", "Q", "LAMB", "PHI", "FU", "FI", "UTHD", "ITHD", "NONE", "NONE", "NONE", "NONE", "NONE"}


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

        For elm As Integer = 1 To InputElements
            For itm As Integer = 1 To IPWAN.PA_Attributes.Items

                Visa.SendString(":NUMERIC:NORMAL:ITEM" & itm + (elm - 1) * IPWAN.PA_Attributes.Items & " " & ItemDef(itm - 1) & ", " & GetElement(elm))

            Next itm
        Next elm

    End Sub
#End Region

#Region "Interface Methodes IPWAN"
    Public Overridable Sub SetWiring(iWir As IPWAN.Wiring) Implements IPWAN.SetWiring
        Dim wirType As String = vbNullString

        Select Case iWir
            Case IPWAN.Wiring.P1W2x3
                wirType = "1"
            Case IPWAN.Wiring.P1W3, IPWAN.Wiring.P1W3AndP1W2
                wirType = "2"
            Case IPWAN.Wiring.P3W3, IPWAN.Wiring.P3W3AndP1W2
                wirType = "3"
            Case IPWAN.Wiring.P3W3M2
                wirType = "4"
            Case IPWAN.Wiring.V3A3
                wirType = "5"
            Case IPWAN.Wiring.P3W3M3
                wirType = "6"
            Case IPWAN.Wiring.P3W4
                wirType = "7"
        End Select

        Visa.SendString(":WIRING TYPE" & wirType)

    End Sub

    Public Overridable Sub SetRectifierMode(iMode As IPWAN.RectifierMode, Optional elm As IPWAN.Elements = IPWAN.Elements.Sigma) Implements IPWAN.SetRectifierMode
        Dim strMode As String
        Select Case iMode
            Case IPWAN.RectifierMode.RMS
                strMode = "RMS"
            Case IPWAN.RectifierMode.VMEAN
                strMode = "VMEAN"
            Case IPWAN.RectifierMode.modedc
                strMode = "DC"
            Case Else
                strMode = "RMS"
        End Select

        Visa.SendString(":INPUT:MODE " & strMode)

    End Sub

    Public Overridable Function GetVrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVrms
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.voltage, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetVPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakPlus
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Voltage_Peak_Plus, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetVPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakMinus
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Voltage_Peak_Minus, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function


    Public Overridable Function GetIrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIrms
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Current, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetIPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakPlus
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Current_Peak_Plus, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetIPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakMinus
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Current_Peak_Minus, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetPactive(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPactive
        Dim oVals(1) As Single

        SetNumericItem(IPWAN.PA_Function.Active_power, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetPapparent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPapparent
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Apparent_power, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetPreact(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPreact
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Reactive_power, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetUUTHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetUTHD
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.THD_Voltage, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetITHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetITHD
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.THD_Current, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetFreqU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqU
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Freq_Voltage, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetFreqI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqI
        Dim oVals(1) As Single
        SetNumericItem(IPWAN.PA_Function.Freq_Current, elm)
        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetPF(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPF
        Dim oVals(1) As Single

        If elm = IPWAN.Elements.Sigma Then
            SetNumericItem(IPWAN.PA_Function.PF_Total, elm)
        Else
            SetNumericItem(IPWAN.PA_Function.PF, elm)
        End If

        cHelper.Delay(1)
        Call QueryNumericItems(oVals)
        Return oVals(0)
    End Function

    Public Overridable Function GetHarmonicsU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single() Implements IPWAN.GetHarmonicsU

        Dim harm() As Single = getHarmonics(elm, IPWAN.PA_Function.voltage)

        Return harm

    End Function

    Public Overridable Function GetHarmonicsI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single() Implements IPWAN.GetHarmonicsI

        Dim harm() As Single = getHarmonics(elm, IPWAN.PA_Function.Current)

        Return harm

    End Function

    Public Overridable Sub SetDisplayItem(nFn As IPWAN.PA_Function, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional disp As IPWAN.PA_Display = IPWAN.PA_Display.a) Implements IPWAN.SetDisplayItem

        Dim sELM As String = GetElement(elm)
        Dim sFN As String = GetFunction(nFn)

        Visa.SendString(":DISPLAY:NORMAL:ITEM" & disp & " " & sFN & "," & sELM)

    End Sub


    Public Overridable Sub SetNumericItem(nFn As IPWAN.PA_Function, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional itm As Integer = 1, Optional ordHarm As Integer = 0) Implements IPWAN.SetNumericItem
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

    Public Overridable Sub SetNumericItemsCount(nCount As Integer) Implements IPWAN.SetNumericItemsCount

        Visa.SendString(":NUMERIC:NORMAL:NUMBER " & nCount.ToString)

    End Sub


    Public Overridable Sub ClearNumericItems() Implements IPWAN.ClearNumericItems
        Visa.SendString(":NUMERIC:NORMAL:CLEAR ALL")
    End Sub

    Public Overridable Sub QueryNumericItems(ByRef oVals() As Single) Implements IPWAN.QueryNumericItems

        oVals = QueryValueList(":NUMERIC:NORMAL:VALUE?")

    End Sub

    Public Overridable Sub SetInputVoltageRange(Optional nRangeInVolts As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetInputVoltageRange
        If nRangeInVolts = 0 Then
            Visa.SendString(":INPUT:VOLTAGE:AUTO ON")
        Else
            Visa.SendString(":INPUT:VOLTAGE:RANGE " & CStr(nRangeInVolts) & " V")
        End If
    End Sub

    Public Overridable Sub SetInputCurrentRange(Optional nRangeInAmps As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetInputCurrentRange
        If nRangeInAmps = 0 Then
            Visa.SendString(":INPUT:CURRENT:AUTO ON")
        Else
            Visa.SendString(":INPUT:CURRENT:RANGE " & CStr(nRangeInAmps) & " A")
        End If
    End Sub

    Public Overridable Sub PresetCurrentTransformer(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentTransformer

        Visa.SendString(":INPUT:SCALING:STATE OFF")
        Visa.SendString(":INPUT:RCONFIG ON")
        Visa.SendString(":INPUT:SCALING:CT:ELEMENT" & elm & " " & FormatNumber(sRatio))
        Visa.SendString(":INPUT:SCALING:STATE ON")

    End Sub

    Public Overridable Sub PresetVoltDivider(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetVoltDivider

        Visa.SendString(":INPUT:SCALING:STATE OFF")
        Visa.SendString(":INPUT:SCALING:VT:ELEMENT" & elm & " " & FormatNumber(sRatio))
        Visa.SendString(":INPUT:SCALING:STATE ON")

    End Sub

    Public Overridable Sub PresetCurrentShunt(shuntRes As Single, sRange As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentShunt

        Visa.SendString(":INPUT:CURRENT:RANGE EXTERNAL," & FormatNumber(sRange) & "V")
        Visa.SendString(":INPUT:CURRENT:SRATIO:ELEMENT" & elm & " " & FormatNumber(shuntRes))

    End Sub

    Public Overridable Sub StartIntegration() Implements IPWAN.StartIntegration

        Visa.SendString("INTEGrate:RESet")
        Visa.SendString("INTEGrate:STARt")

    End Sub

    Public Overridable Sub StopIntegration() Implements IPWAN.StopIntegration

        Visa.SendString("INTEGrate:STop")

    End Sub

    Public Overridable Function GetIntegratedPower(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIntegratedPower
        Dim dRet As Double
        Dim measData() As Single = GetIntegrationsData()

        If durationInSec > 0 Then
            dRet = (measData(1) * 3600) / durationInSec
        End If

        Return dRet

    End Function

    Public Overridable Function GetIntegratedCurrent(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIntegratedCurrent
        Dim dRet As Double
        Dim measData() As Single = GetIntegrationsData()

        If durationInSec > 0 Then
            dRet = (measData(1) * 3600) / durationInSec
        End If

        Return dRet

    End Function


#End Region

#Region "Help functions"
    Overridable Function GetFunction(nFn As IPWAN.PA_Function) As String
        Dim sRet As String

        Select Case nFn
            Case IPWAN.PA_Function.voltage
                sRet = "U"
            Case IPWAN.PA_Function.Voltage_Peak_Plus
                sRet = "UPPeak"
            Case IPWAN.PA_Function.Voltage_Peak_Minus
                sRet = "UMPeak"
            Case IPWAN.PA_Function.Current
                sRet = "I"
            Case IPWAN.PA_Function.Current_Peak_Plus
                sRet = "IPPeak"
            Case IPWAN.PA_Function.Current_Peak_Minus
                sRet = "IMPeak"
            Case IPWAN.PA_Function.Active_power
                sRet = "P"
            Case IPWAN.PA_Function.Active_power_Peak_Plus
                sRet = "PPPeak"
            Case IPWAN.PA_Function.Active_power_Peak_Minus
                sRet = "MPPeak"
            Case IPWAN.PA_Function.Apparent_power
                sRet = "S"
            Case IPWAN.PA_Function.Reactive_power
                sRet = "Q"
            Case IPWAN.PA_Function.PF
                sRet = "PF"
            Case IPWAN.PA_Function.PF_Total
                sRet = "LAMBDA"
            Case IPWAN.PA_Function.THD_Voltage
                sRet = "UTHD"
            Case IPWAN.PA_Function.THD_Current
                sRet = "ITHD"
            Case IPWAN.PA_Function.Freq_Voltage
                sRet = "FU"
            Case IPWAN.PA_Function.Freq_Current
                sRet = "FI"
            Case Else
                sRet = "U"
        End Select

        Return sRet
    End Function


    Overridable Function getHarmonics(elm As Integer, fn As IPWAN.PA_Function, Optional nHarmCount As Integer = 50)

        Dim harm() As Single

        Call ClearNumericItems()
        Call SetNumericItemsCount(nHarmCount)

        For i As Integer = 1 To nHarmCount
            Call SetNumericItem(fn, elm, i, i)
        Next i

        Call cHelper.Delay(1)

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        Call QueryNumericItems(harm)
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

        Return harm
    End Function


    Overridable Function GetElement(nElm As Integer) As String
        Dim sELM As String
        If nElm <> IPWAN.Elements.Element1 Then nElm = IPWAN.Elements.Element1

        sELM = ""
        If nElm = IPWAN.Elements.Sigma Then
            sELM = "SIGMA"
        Else
            sELM = nElm
        End If
        Return sELM
    End Function


    Overridable Function GetIntegrationsData() As Single()

        Visa.SendString("MEASure:NORMal:Item:PRESet INTEGrate")

        Return QueryValueList("Measure: Value?")

    End Function


    Private Function QueryValueList(ByVal cmdStr As String) As Single()
        Dim oVals() As Single

        Visa.SendString(cmdStr)
        Dim buffer As String = Visa.ReceiveString()

        Dim strVals() As String = Split(buffer, ",")

        For i As Integer = 0 To UBound(strVals)
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
            oVals(i) = Single.MinValue
#Enable Warning BC42104 ' Variable is used before it has been assigned a value
            If IsNumeric(strVals(i)) Then
                oVals(i) = CSng(strVals(i))
            End If
        Next i

        Return oVals

    End Function



#End Region



End Class
