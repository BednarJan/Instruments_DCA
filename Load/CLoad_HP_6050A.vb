'Class CLoad_HP_6050A
'10.04.2019, A. Zahler
'Compatible Instruments:
'- HP 6050A ULD02001 (3-channel)
'  - CH1: ULD03005 60507B 150V/60A/500W
'  - CH2: ULD03006 60507B 150V/60A/500W
'  - CH3: ULD03007 60507B 150V/60A/500W

'- HP 6050A ULD02002 (5-channel)
'  - CH1: ULD03001 60501B 60V/30A/150W
'  - CH2: ULD03002 60501B 60V/30A/150W
'  - CH3: ULD03024 60501B 60V/30A/150W
'  - CH4: ULD03003 60503B 240V/10A/250W
'  - CH5: ULD03025 60501B 60V/30A/150W

'- HP 6050A ULD02003 (5-channel):
'  - CH1: ULD03026 60501B 60V/30A/150W
'  - CH2: ULD03027 60501B 60V/30A/150W
'  - CH3: ULD03028 60501B 60V/30A/150W
'  - CH4: ULD03004 60501B 240V/10A/250W
'  - CH5: ULD03029 60501B 60V/30A/150W
Imports Ivi.Visa

Public Class CLoad_HP_6050A
    Implements IDevice
    Implements ILoad

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMax As Single = 150 Implements ILoad.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 60 Implements ILoad.CurrentMax
    Public ReadOnly Property PowerMax As Single = 500 Implements ILoad.PowerMax
    Public Property NumOfChannel As Byte = 1
    Public Property ChannelNo As Byte = 1
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
        Dim Channel As Integer = 0

        _Visa.SendString("*CLS;*RST")

        For Channel = 1 To NumOfChannel
            _Visa.SendString("CHANNEL:LOAD " & Channel)
            _Visa.SendString("MODE:CURR")
            _Visa.SendString("CURR:RANGE MAX")
            _Visa.SendString("INP:STAT ON")
        Next Channel
    End Sub
#End Region

#Region "Interface ILoad Methodes"
    Public Sub SetCC(Current As Single) Implements ILoad.SetCC
        Dim Channel As Integer = 0


        If Current > _CurrentMax Then Current = _CurrentMax

        If _ChannelNo > 0 Then 'Set single channel
            If Current = 0 Then
                _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
                _Visa.SendString("TRAN OFF")
                _Visa.SendString("INPUT OFF")
                _Visa.SendString("MODE:CURR")
                _Visa.SendString("CURR:TRIG " & Current)
                _Visa.SendString("TRIG:IMM")
            Else
                _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
                _Visa.SendString("TRAN OFF")
                _Visa.SendString("INPUT ON")
                _Visa.SendString("MODE:CURR")
                _Visa.SendString("CURR:TRIG " & Current)
                _Visa.SendString("TRIG:IMM")
            End If
        Else 'Synchronize all channels
            If Current = 0 Then
                For Channel = 1 To _NumOfChannel
                    _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
                    _Visa.SendString("TRAN OFF")
                    _Visa.SendString("INPUT OFF")
                    _Visa.SendString("MODE:CURR")
                    _Visa.SendString("CURR:TRIG " & Current)
                Next Channel
                _Visa.SendString("TRIG:IMM")
            Else
                For Channel = 1 To _NumOfChannel
                    _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
                    _Visa.SendString("TRAN OFF")
                    _Visa.SendString("INPUT ON")
                    _Visa.SendString("MODE:CURR")
                    _Visa.SendString("CURR:TRIG " & Current / _NumOfChannel)
                Next Channel
                _Visa.SendString("TRIG:IMM")
            End If
        End If
    End Sub

    Public Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR
        Dim Channel As Integer = 0

        If _ChannelNo > 0 Then 'Set single channel
            _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
            _Visa.SendString("TRAN OFF")
            _Visa.SendString("INPUT ON")
            _Visa.SendString("MODE:RES")
            _Visa.SendString("RES:RANGE " & Resistance)
            _Visa.SendString("RES:TRIG " & Resistance)
            _Visa.SendString("TRIG:IMM")
        Else 'Synchronize all channels
            For Channel = 1 To _NumOfChannel
                _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
                _Visa.SendString("TRAN OFF")
                _Visa.SendString("INPUT ON")
                _Visa.SendString("MODE:RES")
                _Visa.SendString("RES:RANGE " & Resistance)
                _Visa.SendString("RES:TRIG " & Resistance * _NumOfChannel)
            Next Channel
            _Visa.SendString("TRIG:IMM")
        End If
    End Sub

    Public Sub SetCV(Voltage As Single) Implements ILoad.SetCV
        'Applicable for single channel only
        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString("TRAN OFF")
        _Visa.SendString("MODE:VOLT")
        _Visa.SendString("VOLT " & Voltage)
        _Visa.SendString("INPUT ON")
    End Sub

    Public Sub SetCP(Voltage As Single) Implements ILoad.SetCP
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        'Applicable for single channel only
        If CC1 > _CurrentMax Then CC1 = _CurrentMax
        If CC2 > _CurrentMax Then CC2 = _CurrentMax

        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString("TRAN OFF")
        _Visa.SendString("INPUT OFF")
        _Visa.SendString("MODE:CURR")
        _Visa.SendString("CURR " & CC1)
        _Visa.SendString("CURR:TLEV " & CC2)
        _Visa.SendString("CURR:SLEW " & 1000000 * SlewRate)
        _Visa.SendString("TRAN:MODE CONT")
        _Visa.SendString("TRAN:FREQ " & Freq)
        _Visa.SendString("TRAN:DCYC 50")
        _Visa.SendString("TRAN ON")
        _Visa.SendString("INPUT ON")
    End Sub

    'Public Sub SetCRDyn(CR1 As Single, CR2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCRDyn
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetCVDyn(CV1 As Single, CV2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCVDyn
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetCPDyn(CP1 As Single, CP2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCPDyn
    '    Throw New NotImplementedException()
    'End Sub

    Public Sub SetCCPulse(ByVal CCInitial As Single, ByVal CurrPulse As Single, Duration_ms As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Dim Channel As Integer = 0
        Dim Width_s As Single

        If CurrPulse > _CurrentMax Then CurrPulse = _CurrentMax
        If Duration_ms < 0.05 Then Duration_ms = 0.05
        If Duration_ms > 4000 Then Duration_ms = 4000
        width_s = Duration_ms / 1000 'Load expects pulse width in s

        If _ChannelNo > 0 Then 'Set single channel
            _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
            _Visa.SendString("INPUT ON")
            _Visa.SendString("TRAN ON")
            _Visa.SendString("TRAN:MODE CONT")
            _Visa.SendString("CURR " & CCInitial)
            _Visa.SendString("CURR:TLEV " & CurrPulse)
            _Visa.SendString("CURR:SLEW " & 1000000 * SlewRate)
            _Visa.SendString("TRAN:TWID " & Width_s)
            _Visa.SendString("TRIG:IMM")
            cHelper.Delay(width_s + 0.1)
            _Visa.SendString("TRAN OFF")
        Else 'Synchronize all channels
            For Channel = 1 To _NumOfChannel
                _Visa.SendString("CHANNEL:LOAD " & Channel)
                _Visa.SendString("INPUT ON")
                _Visa.SendString("TRAN ON")
                _Visa.SendString("TRAN:MODE CONT")
                _Visa.SendString("CURR " & CCInitial / NumOfChannel)
                _Visa.SendString("CURR:TLEV " & CurrPulse / NumOfChannel)
                _Visa.SendString("CURR:SLEW " & 1000000 * SlewRate)
                _Visa.SendString("TRAN:TWID " & Width_s)
            Next Channel
            _Visa.SendString("TRIG:IMM")
            cHelper.Delay(Width_s + 0.1)
            _Visa.SendString("TRAN OFF")
        End If
    End Sub

    Public Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString("INPUT OFF")
    End Sub

    Public Sub SetOutputON() Implements ILoad.SetOutputON
        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString("INPUT ON")
    End Sub

    Public Sub SetShortON() Implements ILoad.SetShortON
        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString(":INPUT:SHOR ON")
    End Sub

    Public Sub SetShortOFF() Implements ILoad.SetShortOFF
        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString(":INPUT:SHOR OFF")
    End Sub

    Public Sub SetModeCC() Implements ILoad.SetModeCC
        '        _Visa.SendString("MODE CCH")
    End Sub

    Public Sub SetModeCR() Implements ILoad.SetModeCR
        '        _Visa.SendString("MODE CRH")
    End Sub

    Public Sub SetModeCV() Implements ILoad.SetModeCV
        '        _Visa.SendString("MODE CVH")
    End Sub

    Public Sub SetModeCP() Implements ILoad.SetModeCP
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        'Note: Slew rate rise and fall cannot be separately set
        Dim Channel As Integer = 0

        If _ChannelNo > 0 Then 'Set single channel
            _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
            _Visa.SendString("CURR:SLEW " & 1000000 * SlewRateRise)
        Else 'Synchronize all channels
            For Channel = 1 To _NumOfChannel
                _Visa.SendString("CHANNEL:LOAD " & Channel)
                _Visa.SendString("CURR:SLEW " & 1000000 * SlewRateRise)
            Next Channel
        End If
    End Sub

    Public Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        'Note: Slew rate rise and fall cannot be separately set
        'Note: Same command for slew rate CC and CR
        Dim Channel As Integer = 0

        If _ChannelNo > 0 Then 'Set single channel
            _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
            _Visa.SendString("CURR:SLEW " & 1000000 * SlewRateRise)
        Else 'Synchronize all channels
            For Channel = 1 To _NumOfChannel
                _Visa.SendString("CHANNEL:LOAD " & Channel)
                _Visa.SendString("CURR:SLEW " & 1000000 * SlewRateRise)
            Next Channel
        End If
    End Sub

    'Public Sub SetRangeAuto() Implements ILoad.SetRangeAuto
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetRangeHigh() Implements ILoad.SetRangeHigh
    '        _Visa.SendString("CURR:RANGE MAX")
    'End Sub

    'Public Sub SetRangeLow() Implements ILoad.SetRangeLow
    '        _Visa.SendString("CURR:RANGE MIN")
    'End Sub

#End Region

#Region "Public Special Functions HP 6050A"
    Private Sub ClearProt()
        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString("INP:PROT:CLE")
    End Sub

    Private Function GetStatus() As Single
        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString("STAT:QUES:COND?")
        Return _Visa.ReceiveString
    End Function

#End Region

End Class
