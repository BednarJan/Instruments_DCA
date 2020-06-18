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

    Public ReadOnly Property FunctionList As SortedList
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

        _FunctionList = CreateFunctionList()

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

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.Voltage, elm))

    End Function

    Public Overridable Function GetVPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakPlus

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.VoltPeakPlus, elm))

    End Function

    Public Overridable Function GetVPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakMinus

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.VoltPeakMinus, elm))

    End Function

    Public Overridable Function GetIrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIrms

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.Current, elm))

    End Function

    Public Overridable Function GetIPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakPlus

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.CurrentPeakPlus, elm))

    End Function

    Public Overridable Function GetIPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakMinus

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.CurrentPeakMinus, elm))

    End Function

    Public Overridable Function GetPactive(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPactive

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.ActivePower, elm))

    End Function

    Public Overridable Function GetPapparent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPapparent

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.ApparentPower, elm))

    End Function

    Public Overridable Function GetPreact(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPreact

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.ReactivPower, elm))

    End Function

    Public Overridable Function GetUTHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetUTHD

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.THDvolt, elm))

    End Function

    Public Overridable Function GetITHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetITHD

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.THDCurr, elm))

    End Function

    Public Overridable Function GetFreqU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqU

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.FrequencyU, elm))

    End Function

    Public Overridable Function GetFreqI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqI

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.FrequencyI, elm))

    End Function

    Public Overridable Function GetPF(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPF

        Dim oVals() As Single = QueryNumericItems()
        Return oVals(GetFunctionIndex(IPWAN.PA_Function.PF, elm))

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

    Public Overridable Function QueryNumericItems() As Single() Implements IPWAN.QueryNumericItems

        Dim oVals() As Single = QueryValueList(":NUMERIC:NORMAL:VALUE?")

        Return oVals

    End Function

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
        Dim sRet As String = vbNullString
        Dim myKeys As ICollection = FunctionList.Keys

        For i As Integer = 0 To FunctionList.Count - 1
            If FunctionList.ContainsKey(nFn.ToString) Then
                sRet = FunctionList.GetByIndex(i).ToString
            End If
        Next

        Return sRet
    End Function

    Overridable Function GetFunctionIndex(nFn As IPWAN.PA_Function, nElm As Integer) As Integer
        Dim nRet As Integer

        For i As Integer = 0 To FunctionList.Count - 1

            Dim fn As KeyValuePair(Of String, String) = FunctionList(i)
            If fn.Key = nFn.ToString Then
                nRet = (nElm - 1) * FunctionList.Count + i + 1
            End If

        Next

        Return nRet

    End Function

    Overridable Function getHarmonics(elm As Integer, fn As IPWAN.PA_Function, Optional nHarmCount As Integer = 50)

        Call ClearNumericItems()
        Call SetNumericItemsCount(nHarmCount)

        For i As Integer = 1 To nHarmCount
            Call SetNumericItem(fn, elm, i, i)
        Next i

        Call cHelper.Delay(1)
        Dim harm() As Single = QueryNumericItems()

        Return harm
    End Function

    Overridable Function GetElement(nElm As Integer) As String
        Dim sELM As String
        If nElm <> IPWAN.Elements.Element1 Then nElm = IPWAN.Elements.Element1

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

    Overridable Function CreateFunctionList() As SortedList

        Visa.SendString(":NUMERIC:NORMAL:PRESET 3" & Chr(10))

        'Fix item order within the preset pattern 3 

        Dim fsl As New SortedList

        fsl.Add(IPWAN.PA_Function.Voltage.ToString, "U")
        fsl.Add(IPWAN.PA_Function.Current.ToString, "I")
        fsl.Add(IPWAN.PA_Function.ActivePower.ToString, "P")
        fsl.Add(IPWAN.PA_Function.ApparentPower.ToString, "S")
        fsl.Add(IPWAN.PA_Function.ReactivPower.ToString, "Q")
        fsl.Add(IPWAN.PA_Function.PF.ToString, "LAMBda")
        fsl.Add(IPWAN.PA_Function.PhaseDiff.ToString, "PHI")
        fsl.Add(IPWAN.PA_Function.FrequencyU.ToString, "FU")
        fsl.Add(IPWAN.PA_Function.FrequencyI.ToString, "FI")
        fsl.Add(IPWAN.PA_Function.VoltPeakPlus, "UPP")
        fsl.Add(IPWAN.PA_Function.VoltPeakMinus.ToString, "UMP")
        fsl.Add(IPWAN.PA_Function.CurrentPeakPlus.ToString, "IPP")
        fsl.Add(IPWAN.PA_Function.CurrentPeakMinus.ToString, "IMP")
        fsl.Add(IPWAN.PA_Function.PowerPeakPlus.ToString, "PPP")
        fsl.Add(IPWAN.PA_Function.PowerPeakMinus.ToString, "PMP")

        Return fsl

    End Function

    Overridable Sub PresetPattern()

        Visa.SendString(":NUMERIC:NORMAL:PRESET 3" & Chr(10))

    End Sub






#End Region



End Class
