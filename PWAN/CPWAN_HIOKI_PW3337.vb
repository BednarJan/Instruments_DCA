﻿'Class CPWAN_HIOKI_PW3337
'02.05.2019, J. Bednar 
'17.06.2020, J. Bednar
'Compatible Instruments:
'- HIOKI PW3337
Imports Ivi.Visa

Public Class CPWAN_HIOKI_PW3337
    Inherits BCPWAN
    Implements IPWAN


#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 1000
        CurrentMax = 65
        InputElements = 3
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"


    Public Overrides Sub Initialize() Implements IDevice.Initialize
        Dim ItemDef() As String = {"U", "I", "P", "S", "Q", "LAMB", "PHI", "FU", "FI", "UTHD", "ITHD", "NONE", "NONE", "NONE", "NONE", "NONE"}

        Visa.SendString("*RST;*CLS" & Chr(10))
        Visa.SendString("HEADER OFF" & Chr(10))

        SetDisplayItem(IPWAN.PA_Function.Voltage, IPWAN.PA_Display.a, IPWAN.RectifierMode.AC, "1")
        SetDisplayItem(IPWAN.PA_Function.Voltage, IPWAN.PA_Display.b, IPWAN.RectifierMode.AC, "1")
        SetDisplayItem(IPWAN.PA_Function.Voltage, IPWAN.PA_Display.c, IPWAN.RectifierMode.AC, "1")
        SetDisplayItem(IPWAN.PA_Function.Voltage, IPWAN.PA_Display.d, IPWAN.RectifierMode.AC, "0")

        Dim nItemsCount As Integer = CreateNumericItemsList()

    End Sub
#End Region

#Region "Interface Methodes IPWAN"

    Public Overrides Function QueryNumericItems() As Double() Implements IPWAN.QueryNumericItems

        Return QueryValueList(":MEASURE?")

    End Function

    Public Overrides Sub ClearNumericItems() Implements IPWAN.ClearNumericItems

        Call Visa.SendString(":MEAS:ITEM:ALLC")

    End Sub

    Public Overrides Sub SetNumericItem(nFn As IPWAN.PA_Function, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional itm As Integer = 1, Optional ordHarm As Integer = 0) Implements IPWAN.SetNumericItem
        Dim cmdStr As String
        Dim sELM As String
        Dim sFN As String

        sELM = GetElement(elm)
        sFN = GetFunction(nFn)

        cmdStr = ":MEAS:ITEM:" & itm.ToString & " " & sFN & "," & sELM
        If ordHarm > 0 Then
            cmdStr = cmdStr & "," & ordHarm
        End If

        Visa.SendString(cmdStr)

    End Sub

    Overrides Function CreateNumericItemsList() As Integer Implements IPWAN.CreateNumericItemsList

        Return MyBase.CreateNumericItemsList

    End Function



    Public Overrides Sub SetDisplayItem(nFn As IPWAN.PA_Function, disp As IPWAN.PA_Display, nRectMode As IPWAN.RectifierMode, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetDisplayItem

        Dim cmdStr As String = ":DISPlay:NORMal:" & disp.ToString
        cmdStr &= " "
        cmdStr &= GetFunction(nFn)
        cmdStr &= GetRectifierMode(nRectMode)
        cmdStr &= "1"
        Visa.SendString(cmdStr)

    End Sub

    Public Overrides Sub PresetCurrentTransformer(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentTransformer

        MyBase.PresetCurrentTransformer(sRatio, elm)

        'Visa.SendString(":INPUT:SCALING:STATE OFF")
        'Visa.SendString(":INPUT:RCONFIG ON")
        'Visa.SendString(":INPUT:SCALING:CT:ELEMENT" & elm & " " & FormatNumber(sRatio))
        'Visa.SendString(":INPUT:SCALING:STATE ON")

    End Sub

    Public Overrides Sub PresetVoltDivider(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetVoltDivider

        MyBase.PresetVoltDivider(sRatio, elm)
        'Visa.SendString(":INPUT:SCALING:STATE OFF")
        'Visa.SendString(":INPUT:SCALING:VT:ELEMENT" & elm & " " & FormatNumber(sRatio))
        'Visa.SendString(":INPUT:SCALING:STATE ON")

    End Sub

    Public Overrides Sub PresetCurrentShunt(shuntRes As Single, sRange As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentShunt

        MyBase.PresetCurrentShunt(shuntRes, sRange, elm)

        'Visa.SendString(":INPUT:CURRENT:RANGE EXTERNAL," & FormatNumber(sRange) & "V")
        'Visa.SendString(":INPUT:CURRENT:SRATIO:ELEMENT" & elm & " " & FormatNumber(shuntRes))

    End Sub

    Public Overrides Sub PresetIntegrationPower(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetIntegrationPower

        Call SetDisplayItem(IPWAN.PA_Function.Voltage, elm, IPWAN.PA_Display.a)
        Call SetDisplayItem(IPWAN.PA_Function.Current, elm, IPWAN.PA_Display.b)
        Call SetDisplayItem(IPWAN.PA_Function.ActivePower, elm, IPWAN.PA_Display.c)
        Call SetDisplayItem(IPWAN.PA_Function.IntegratedActivePower, elm, IPWAN.PA_Display.d)
        Call ClearNumericItems()
        Call SetNumericItem(IPWAN.PA_Function.IntegratedActivePower, elm)

    End Sub

    Public Overrides Sub PresetIntegrationCurrent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetIntegrationCurrent

        Call SetDisplayItem(IPWAN.PA_Function.Voltage, elm, IPWAN.PA_Display.a)
        Call SetDisplayItem(IPWAN.PA_Function.Current, elm, IPWAN.PA_Display.b)
        Call SetDisplayItem(IPWAN.PA_Function.ActivePower, elm, IPWAN.PA_Display.c)
        Call SetDisplayItem(IPWAN.PA_Function.IntegratedCurrent, elm, IPWAN.PA_Display.d)
        Call ClearNumericItems()
        Call SetNumericItem(IPWAN.PA_Function.IntegratedCurrent, elm)

    End Sub





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

    Public Overrides Sub SetWiring(iWir As IPWAN.Wiring) Implements IPWAN.SetWiring
        Dim wirType As String = vbNullString

        Select Case iWir
            Case IPWAN.Wiring.P1W3
                wirType = "2"
            Case IPWAN.Wiring.P3W3
                wirType = "3"
            Case IPWAN.Wiring.V3A3
                wirType = "5"
            Case IPWAN.Wiring.P3W4
                wirType = "7"
        End Select

        Visa.SendString(":WIRING TYPE" & wirType)

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
    '''Set rectifier mode for the specified input element and measurement
    '''
    Public Overrides Sub SetInputMode(iMode As IPWAN.RectifierMode) Implements IPWAN.SetInputMode

        Throw New NotImplementedException

    End Sub

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

        Call PresetHarmonics(IPWAN.PA_Function.Voltage, elm)

        Return getHarmonics(elm, IPWAN.PA_Function.Voltage)

    End Function

    Public Overrides Function GetHarmonicsI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single() Implements IPWAN.GetHarmonicsI

        Call PresetHarmonics(IPWAN.PA_Function.Current, elm)

        Return getHarmonics(elm, IPWAN.PA_Function.Current)

    End Function


#End Region

#Region "Public Special Functions HIOKI PW3337"



    Overrides Function CreateFunctionList() As SortedList

        'Fix item order analogical the Yokogawa preset pattern 3 

        Dim fsl As New SortedList
        fsl.Add(IPWAN.PA_Function.Voltage.ToString, "U")
        fsl.Add(IPWAN.PA_Function.Current.ToString, "I")
        fsl.Add(IPWAN.PA_Function.ActivePower.ToString, "P")
        fsl.Add(IPWAN.PA_Function.ApparentPower.ToString, "S")
        fsl.Add(IPWAN.PA_Function.ReactivPower.ToString, "Q")
        fsl.Add(IPWAN.PA_Function.PF.ToString, "PF")
        fsl.Add(IPWAN.PA_Function.PhaseDiff.ToString, "DEG")
        fsl.Add(IPWAN.PA_Function.FrequencyU.ToString, "FREQU")
        fsl.Add(IPWAN.PA_Function.FrequencyI.ToString, "FREQI")
        fsl.Add(IPWAN.PA_Function.VoltPeakPlus.ToString, "U_MAX")
        fsl.Add(IPWAN.PA_Function.VoltPeakMinus.ToString, "U_MIN")
        fsl.Add(IPWAN.PA_Function.CurrentPeakPlus.ToString, "I_MAX")
        fsl.Add(IPWAN.PA_Function.CurrentPeakMinus.ToString, "I_MIN")
        fsl.Add(IPWAN.PA_Function.PowerPeakPlus.ToString, "P_MAX")
        fsl.Add(IPWAN.PA_Function.PowerPeakMinus.ToString, "P_MIN")
        Return fsl

    End Function

    Overrides Sub PresetPattern()
        Dim fsl As SortedList = CreateFunctionList()

        For elm As Integer = 1 To InputElements

            For itm As Integer = 0 To fsl.Count - 1

                Dim fn As KeyValuePair(Of String, String) = fsl(itm)

                SetNumericItem(fn.Value, elm, (elm - 1) * fsl.Count + itm)

            Next
        Next
    End Sub

    Public Overrides Sub SetNumericItemsCount(nCount As Integer)

        Throw New NotImplementedException

    End Sub

#End Region

#Region "PPrivate  Functions HIOKI PW3337"
    'Private Function QueryNumericItem(nFn As IPWAN.PA_Function, nElm As IPWAN.Elements) As Single


    '    Visa.SendString(":MEASURE? " & GetFunction(nFn) & nElm.ToString)

    '    Dim Buffer As String = Visa.ReceiveString()

    '    If IsNumeric(Buffer) Then
    '        Return CSng(Buffer)
    '    Else
    '        Return Single.MinValue
    '    End If

    'End Function

    Private Function QueryValueList(ByVal cmdStr As String) As Double()

        Visa.SendString(cmdStr)
        cHelper.Delay(1)
        Return Visa.ReceiveValueList()

    End Function


    Private Sub PresetHarmonics(nFn As IPWAN.PA_Function)

        PresetHarmonics(nFn, IPWAN.Elements.Element1)

    End Sub

    Private Sub PresetHarmonics(nFn As IPWAN.PA_Function, nElm As IPWAN.Elements)

        Call Visa.SendString(":MEASURE:HARMONIC:ITEM:ALLC")
        Call Visa.SendString(":MEASURE:HARMONIC:ITEM:LIST 1,0,0,0,0,0")
        Call Visa.SendString(":MEASURE:HARMONIC:ITEM:ORDER 0," & CStr(_HarmCount) & ",ALL")

        Select Case nFn
            Case IPWAN.PA_Function.Current
                Call Visa.SendString(":MEASURE:HARMONIC:ITEM:I:CH" & nElm & " 1")
                Call Visa.SendString(":MEASURE:HARMONIC:ITEM:U:CH" & nElm & " 0")
            Case IPWAN.PA_Function.Voltage
                Call Visa.SendString(":MEASURE:HARMONIC:ITEM:I:CH" & nElm & " 0")
                Call Visa.SendString(":MEASURE:HARMONIC:ITEM:U:CH" & nElm & " 1")
            Case Else
                Call Visa.SendString(":MEASURE:HARMONIC:ITEM:I:CH" & nElm & " 0")
                Call Visa.SendString(":MEASURE:HARMONIC:ITEM:U:CH" & nElm & " 1")
        End Select

    End Sub



#End Region

End Class
