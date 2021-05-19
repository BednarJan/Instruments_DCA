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

    Public ReadOnly Property FunctionList As SortedList(Of String, String)

    Public ReadOnly Property _HarmCount As Integer = 50

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

        MyBase.Initialize()

    End Sub
#End Region

#Region "Interface Methodes IPWAN"
    Public Overridable Sub SetWiring(iWir As IPWAN.Wiring) Implements IPWAN.SetWiring

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub SetTHDNorm(nTHDNorm As IPWAN.THDNorm) Implements IPWAN.SetTHDNorm

        Throw New NotImplementedException

    End Sub

    '''SetInputMode(iMode As IPWAN.RectifierMode, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional fn As IPWAN.PA_Function = IPWAN.PA_Function.Voltage) Implements IPWAN.SetInputMode
    '''Set rectifier mode for all input elements and measurements. Unable set it separatelly for the input channel
    '''
    Public Overridable Sub SetInputMode(iMode As IPWAN.RectifierMode) Implements IPWAN.SetInputMode

        Throw New NotImplementedException

    End Sub

    Public Overridable Function QueryNumericItems() As Double() Implements IPWAN.QueryNumericItems

        Throw New NotImplementedException

    End Function


    Public Overridable Sub ClearNumericItems() Implements IPWAN.ClearNumericItems

        Throw New NotImplementedException

    End Sub


    Overridable Function CreateNumericItemsList() As Integer Implements IPWAN.CreateNumericItemsList

        Dim sfl As SortedList(Of String, String) = CreateFunctionList()
        Dim nCount As Integer = 0

        For elm As Integer = 1 To 3

            For Each kvp As KeyValuePair(Of String, String) In sfl

                SetNumericItem(kvp.Value, elm, sfl.IndexOfKey(kvp.Key))

                nCount += 1

            Next

        Next

        Return nCount

    End Function

    Public Overridable Sub SetInputVoltageRange(Optional nRangeInVolts As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetInputVoltageRange

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub SetInputCurrentRange(Optional nRangeInAmps As Single = 0, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetInputCurrentRange

        Throw New NotImplementedException

    End Sub


    Public Overridable Function GetVrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVrms

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetVPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakPlus

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetVPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetVPeakMinus

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetIrms(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIrms

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetIPeakPlus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakPlus

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetIPeakMinus(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIPeakMinus

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetPactive(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPactive

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetPapparent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPapparent

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetPreact(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPreact

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetUTHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetUTHD

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetITHD(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetITHD

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetFreqU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqU

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetFreqI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetFreqI

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetPF(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetPF

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetHarmonicsU(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single() Implements IPWAN.GetHarmonicsU

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetHarmonicsI(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single() Implements IPWAN.GetHarmonicsI

        Throw New NotImplementedException

    End Function

    Public Overridable Sub SetDisplayItem(nFn As IPWAN.PA_Function, disp As IPWAN.PA_Display, nRectMode As IPWAN.RectifierMode, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.SetDisplayItem

        Throw New NotImplementedException

    End Sub


    Public Overridable Sub SetNumericItem(nFn As IPWAN.PA_Function, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional itm As Integer = 1, Optional ordHarm As Integer = 0) Implements IPWAN.SetNumericItem

        Dim sFN As String

        sFN = GetFunction(nFn)

        Call SetNumericItem(sFN, CInt(elm), itm, ordHarm)

    End Sub

    Public Overridable Sub SetNumericItem(nFn As String, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1, Optional itm As Integer = 1, Optional ordHarm As Integer = 0) Implements IPWAN.SetNumericItem

        Dim sFN As String

        sFN = GetFunction(nFn)

        SetNumericItem(sFN, CInt(elm), itm, ordHarm)

    End Sub


    Public Overridable Sub SetNumericItem(nFn As String, Optional elm As Integer = 1, Optional itm As Integer = 1, Optional ordHarm As Integer = 0) Implements IPWAN.SetNumericItem

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub PresetCurrentTransformer(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentTransformer

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub PresetVoltDivider(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetVoltDivider

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub PresetCurrentShunt(shuntRes As Single, sRange As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentShunt

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub PresetIntegrationPower(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetIntegrationPower

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub PresetIntegrationCurrent(Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetIntegrationCurrent

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub StartIntegration() Implements IPWAN.StartIntegration

        Throw New NotImplementedException

    End Sub

    Public Overridable Sub StopIntegration() Implements IPWAN.StopIntegration

        Throw New NotImplementedException

    End Sub

    Public Overridable Function GetIntegratedPower(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIntegratedPower

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetIntegratedCurrent(durationInSec As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) As Single Implements IPWAN.GetIntegratedCurrent

        Throw New NotImplementedException

    End Function

    Public Overridable Function GetFunctionIndex(nFn As IPWAN.PA_Function, nElm As Integer) As Integer Implements IPWAN.GetFunctionIndex

        Dim nRet As Integer = Integer.MinValue

        If FunctionList.ContainsKey(nFn.ToString) Then

            nRet = FunctionList.IndexOfKey(nFn.ToString)

        End If

        Return nRet

    End Function


#End Region

#Region "Help functions"

    Overridable Function getHarmonics(elm As Integer, fn As IPWAN.PA_Function)

        Visa.SendString(":MEAS:HARM:VALUE?")
        cHelper.Delay(1)
        Return Visa.ReceiveValueList()

    End Function


    Overridable Function GetElement(nElm As Integer) As String

        Throw New NotImplementedException

    End Function

    Overridable Function GetIntegrationsData() As Double()

        Throw New NotImplementedException

    End Function


    Private Function QueryValueList(ByVal cmdStr As String) As Single()

        Throw New NotImplementedException

    End Function

    Overridable Function CreateFunctionList() As SortedList(Of String, String)


        'this is the fix item order within the preset pattern 3 

        Dim fsl As New SortedList(Of String, String)

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

    Public Overridable Sub SetNumericItemsCount(nCount As Integer)

        Throw New NotImplementedException

    End Sub


    Overridable Function GetFunction(nFn As IPWAN.PA_Function) As String

        Dim sRet As String = String.Empty

        If FunctionList.ContainsKey(nFn.ToString) Then

            sRet = FunctionList.Item(nFn.ToString).ToString

        End If


        Return sRet
    End Function


    Overridable Sub PresetPattern()

        Visa.SendString(":NUMERIC:NORMAL:PRESET 3" & Chr(10))

    End Sub

    Function GetRectifierMode(iMode As IPWAN.RectifierMode) As String

        Return iMode.ToString

    End Function


#End Region


End Class
