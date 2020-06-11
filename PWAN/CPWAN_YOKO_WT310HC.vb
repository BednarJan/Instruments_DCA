'Class CPWAN_YOKO_WT310HC
'06.01.2019, A. Zahler
'Compatible Instruments:
'- Yokogawa WT310HC
'- Yokogawa WT310EH
Imports Ivi.Visa

Public Class CPWAN_YOKO_WT310HC
    Implements IDevice
    Implements IPWAN

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMaxRms As Single = 600 Implements IPWAN.VoltageMax
    Public ReadOnly Property CurrentMaxRms As Single = 40 Implements IPWAN.CurrentMax
    Public ReadOnly Property InputElements As UInteger = 1 Implements IPWAN.InputElements
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Visa = New CVisaDevice(Session, ErrorLogger)
        _ErrorLogger = ErrorLogger
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"
    Public Function IDN() As String Implements IDevice.IDN
        Dim ErrorMessages(1) As String

        _Visa.SendString("*IDN?", ErrorMessages(0))
        Return _Visa.ReceiveString(ErrorMessages(1))

    End Function

    Public Sub RST() Implements IDevice.RST
        Dim ErrorMessage As String = ""

        _Visa.SendString("*RST", ErrorMessage)

    End Sub

    Public Sub CLS() Implements IDevice.CLS
        Dim ErrorMessage As String = ""

        _Visa.SendString("*CLS", ErrorMessage)
    End Sub

    Public Sub Initialize() Implements IDevice.Initialize
        Dim Item As Integer, ItemDef() As String = {"U", "I", "P", "S", "Q", "LAMB", "PHI", "FU", "FI", "UTHD", "ITHD", "NONE", "NONE", "NONE", "NONE", "NONE"}

        _Visa.SendString("*RST;*CLS" & Chr(10))
        _Visa.SendString(":SYSTEM:COMMUNICATE:COMMAND WT300" & Chr(10))
        _Visa.SendString(":HARMONICS:THD TOTAL" & Chr(10))
        _Visa.SendString(":NUMERIC:FORMAT ASCII" & Chr(10))
        _Visa.SendString(":RATE 500MS" & Chr(10))
        _Visa.SendString(":INPUT:VOLTAGE:AUTO ON" & Chr(10))
        _Visa.SendString(":INPUT:CURRENT:AUTO ON" & Chr(10))
        _Visa.SendString(":INPUT:VOLTAGE:CONFIG 600,300,150" & Chr(10)) 'select valid voltage ranges for Auto Range (speeds up range search)
        _Visa.SendString(":DISPLAY:NORMAL:ITEM1 U" & Chr(10))
        _Visa.SendString(":DISPLAY:NORMAL:ITEM2 I" & Chr(10))
        _Visa.SendString(":DISPLAY:NORMAL:ITEM3 P" & Chr(10))
        _Visa.SendString(":DISPLAY:NORMAL:ITEM4 LAMB" & Chr(10))

        For Item = 1 To CPWAN_Helper.EPWAN_Attributes.Items
            _Visa.SendString(":NUMERIC:ITEM" & Item & " " & ItemDef(Item - 1) & Chr(10))
        Next Item
    End Sub
#End Region

#Region "Interface Methodes IPWAN"
    '***** Function from DCA *****
    'Public Sub Init_Integration(durationInSec As Single) Implements IPWAN.Init_Integration
    '    Throw New NotImplementedException()
    'End Sub

    '***** Function from DCA *****
    'Public Sub Get_Integrations_Data(ByRef measData() As Double) Implements IPWAN.Get_Integrations_Data
    '    Throw New NotImplementedException()
    'End Sub

    '***** Function from DCA *****
    'Public Function IntegratePower(durationInSec As Single) As Double Implements IPWAN.IntegratePower
    '    Throw New NotImplementedException()
    'End Function

    '***** Function from DCA *****
    'Public Sub Set_Volt_Range(sRange As String) Implements IPWAN.Set_Volt_Range
    '    Throw New NotImplementedException()
    'End Sub

    '***** Function from DCA *****
    'Public Sub Set_Curr_Range(sRange As String) Implements IPWAN.Set_Cur_Range
    '    Throw New NotImplementedException()
    'End Sub

    Public Function GetPF(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPF
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.PF & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetVrms(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetVrms
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.U & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetVthd(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetVthd
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.THDU & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetIrms(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetIrms
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.I & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetIthd(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetIthd
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.THDI & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPreal(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPreal
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.P & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPapp(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPapp
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.S & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPreact(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPreact
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.Q & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetFreq(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetFreq
        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.FU & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Sub GetHarmonics_IEC61000_3_2(EquipClass As CPWAN_Helper.EEquipmentClass, ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, Optional WaitOnPWAN As Boolean = False) Implements IPWAN.GetHarmonics_IEC61000_3_2
        Dim HarmOrder As Integer, HarmString As String(), Preal As Single

        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":HARMONICS:PLLSOURCE U" & Chr(10))
        _Visa.SendString(":HARMONICS:ORDER 1,40" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:PRESET 2" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:NUMBER 1" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:VALUE?" & CPWAN_Helper.EPWAN_Attributes.I & Chr(10))

        'PWAN delivers string with 40 comma separated values. Split converts into 40 substrings
        HarmString = Split(_Visa.ReceiveString, ",")

        For HarmOrder = 1 To 40
            HarmMeas(HarmOrder - 1) = Val(HarmString(HarmOrder - 1)) 'Converts string to numeric
        Next HarmOrder

        'Read power and calculate limits
        _Visa.SendString(":NUMERIC:VALUE? " & CPWAN_Helper.EPWAN_Attributes.P & Chr(10))
        Preal = _Visa.ReceiveValue

        CPWAN_Helper.GetHarmLimit_IEC61000_3_2(EquipClass, Preal, HarmLimit)

    End Sub

    Public Sub GetHarmonics_IEC61000_3_12(ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, ByRef THD As Single, ByRef PWHD As Single, Optional WaitOnPWAN As Boolean = False) Implements IPWAN.GetHarmonics_IEC61000_3_12
        Dim HarmMeasCurr(39) As Single, HarmString As String(), Sum As Single

        If WaitOnPWAN Then Call AutoRangeDelay()

        _Visa.SendString(":HARMONICS:PLLSOURCE U" & Chr(10))
        _Visa.SendString(":HARMONICS:ORDER 1,40" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:PRESET 2" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:NUMBER 1" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:VALUE?" & CPWAN_Helper.EPWAN_Attributes.I & Chr(10))

        'PWAN delivers string with 40 comma separated values. Split converts into 40 substrings
        HarmString = Split(_Visa.ReceiveString, ",")

        For HarmOrder = 1 To 40
            HarmMeasCurr(HarmOrder - 1) = Val(HarmString(HarmOrder - 1)) 'Converts string to numeric
            HarmMeas(HarmOrder - 1) = HarmMeasCurr(HarmOrder - 1) / HarmMeasCurr(0) * 100 'Calculate perecentage of fundamental current
        Next HarmOrder

        'Calculate THD (2 to 40 order)
        Sum = 0

        For HarmOrder = 2 To 40
            Sum += (HarmMeas(HarmOrder - 1) / HarmMeas(0)) ^ 2
        Next HarmOrder

        THD = 100 * Math.Sqrt(Sum)

        'Calculate PWHD (14 to 40 order)
        Sum = 0

        For HarmOrder = 14 To 40
            Sum += HarmOrder * (HarmMeas(HarmOrder - 1) / HarmMeas(0)) ^ 2
        Next HarmOrder

        PWHD = 100 * Math.Sqrt(Sum)

        CPWAN_Helper.GetHarmLimit_IEC61000_3_12(HarmLimit)

    End Sub

#End Region

#Region "Public Special Functions WT310HC"
    'Required before energy or flicker measurement over a defined time period
    Sub IntegrationReset()
        _Visa.SendString(":INTEGRATE:RESET")
    End Sub

    Private Sub AutoRangeDelay()
        Dim TimerEnd As Double, TimerCounter As Double
        Dim SumCheckRange As Integer

        'Check during one second if Auto Range is "busy"
TimerStart:
        SumCheckRange = 0
        TimerEnd = Timer + 1
        Do While TimerCounter < TimerEnd
            SumCheckRange = SumCheckRange + CheckRange()
            TimerCounter = Timer
        Loop

        If SumCheckRange > 0 Then GoTo TimerStart
    End Sub

    Private Function CheckRange() As Integer
        _Visa.SendString(":INPUT:CRANGE?" & Chr(10))
        Return _Visa.ReceiveValue

    End Function
#End Region

End Class
