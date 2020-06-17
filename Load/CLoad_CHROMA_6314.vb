'Class CLoad_CHROMA_6314
'11.04.2019, A. Zahler
'Single channel control implemented only (no parallel operation of modules)
'Compatible Instruments:
'- Chroma 6314 ULD02004 (4-channel) Tessy1:
'  - CH1: ULD03011 Chroma 63103 150V/60A/500W
'  - CH3: ULD03010 Chroma 63103 150V/60A/500W
'  - CH5: ULD03009 Chroma 63103 150V/60A/500W
'  - CH7: ULD03008 Chroma 63103 150V/60A/500W

'- Chroma 6314 ULD02005 (1-channel) Stock Archive:
'  - CH1: ULD03012 Chroma 63112 80V/240A/1200W

'- Chroma 6314 ULD02006 (1-channel) Tessy4:
'  - CH1: ULD03014 Chroma 63112 80V/240A/1200W

'- Chroma 6314 ULD02007 (1-channel) Stock Archive:
'  - CH1: ULD03013 Chroma 63112 80V/240A/1200W

'- Chroma 6314 ULD02008 (1-channel) Vibration:
'  - CH1: ULD03015 Chroma 63112 80V/240A/1200W

'- Chroma 6314 ULD02017 (1-channel) Vibration:
'  - CH1: ULD03017 Chroma 63112 80V/240A/1200W

'- Chroma 6314 ULD02018 (1-channel) Stock Archive:
'  - CH1: ULD03018 Chroma 63112 80V/240A/1200W

'- Chroma 6314 ULD02015 (4-channel) Spare:
'  - CH1: ULD03021 Chroma 63102 80V/20A/100W
'  - CH3: ULD03022 Chroma 63102 80V/20A/100W
'  - CH5: ULD03019 Chroma 63103 80V/60A/300W
'  - CH7: ULD03020 Chroma 63103 80V/60A/300W
Imports Ivi.Visa

Public Class CLoad_CHROMA_6314
    Implements IDevice
    Implements ILoad

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMax As Single = 80 Implements ILoad.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 240 Implements ILoad.CurrentMax
    Public ReadOnly Property PowerMax As Single = 1200 Implements ILoad.PowerMax
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

    Public Sub SndString(cmdStr As String) Implements IDevice.SendString
        _Visa.SendString(cmdStr)
    End Sub

    Public Function RecieveString() As String Implements IDevice.ReceiveString
        Return _Visa.ReceiveString
    End Function

    Public Sub Initialize() Implements IDevice.Initialize
        Dim Channel As Integer = 0

        _Visa.SendString("*CLS;*RST")

        For Channel = 1 To 8 'Note: 2nd Module = channel 3!
            _Visa.SendString("CHAN " & Channel)
            _Visa.SendString("LOAD OFF")
            _Visa.SendString("MODE CRH")
            _Visa.SendString("RES:L1 100")
            _Visa.SendString("RES:L2 100")
            _Visa.SendString("RES:RISE 1")
            _Visa.SendString("RES:FALL 1")
            _Visa.SendString("MODE CV")
            _Visa.SendString("VOLT:CURR MAX")
            _Visa.SendString("MODE CCH")
            _Visa.SendString("CURR:STAT:L1 0")
            _Visa.SendString("CURR:STAT:L2 0")
            _Visa.SendString("CURR:STAT:RISE 1")
            _Visa.SendString("CURR:STAT:FALL 1")
            _Visa.SendString("LOAD ON")
        Next Channel
    End Sub
#End Region

#Region "Interface ILoad Methodes"
    Public Sub SetCC(Current As Single) Implements ILoad.SetCC
        If Current > _CurrentMax Then Current = _CurrentMax

        If Current = 0 Then
            _Visa.SendString("CHAN " & _ChannelNo)
            _Visa.SendString("MODE CCH")
            _Visa.SendString("CURR:STAT:L1 " & Current)
            _Visa.SendString("LOAD OFF")
        Else
            _Visa.SendString("CHAN " & _ChannelNo)
            _Visa.SendString("MODE CCH")
            _Visa.SendString("CURR:STAT:L1 " & Current)
            _Visa.SendString("LOAD ON")
        End If
    End Sub

    Public Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("MODE CRL")
        _Visa.SendString("RES:L1 " & Resistance)
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetCV(Voltage As Single) Implements ILoad.SetCV
        'Applicable for single channel only
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("MODE CV")
        _Visa.SendString("VOLT:L1 " & Voltage)
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetCP(Voltage As Single) Implements ILoad.SetCP
        Throw New NotImplementedException() 'Not supported from instrument
    End Sub

    Public Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        If CC1 > _CurrentMax Then CC1 = _CurrentMax
        If CC2 > _CurrentMax Then CC2 = _CurrentMax

        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("LOAD OFF")
        _Visa.SendString("MODE CCDH")
        _Visa.SendString("CURR:DYN:L1 " & CC1)
        _Visa.SendString("CURR:DYN:L2 " & CC2)
        _Visa.SendString("CURR:DYN:T1 " & (1 / Freq) / 2)
        _Visa.SendString("CURR:DYN:T2 " & (1 / Freq) / 2)
        _Visa.SendString("CURR:DYN:RISE " & SlewRate)
        _Visa.SendString("CURR:DYN:FALL " & SlewRate)
        _Visa.SendString("LOAD ON")
    End Sub

    'Public Sub SetCRDyn(CR1 As Single, CR2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCRDyn
    '    Throw New NotImplementedException() 'Not supported from instrument
    'End Sub

    'Public Sub SetCVDyn(CV1 As Single, CV2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCVDyn
    '    Throw New NotImplementedException() 'Not supported from instrument
    'End Sub

    'Public Sub SetCPDyn(CP1 As Single, CP2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCPDyn
    '    Throw New NotImplementedException() 'Not supported from instrument
    'End Sub

    Public Sub SetCCPulse(ByVal CCInitial As Single, ByVal CurrPulse As Single, Duration_ms As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Dim Channel As Integer = 0
        Dim Width_s As Single

        If CurrPulse > _CurrentMax Then CurrPulse = _CurrentMax
        If Duration_ms < 0.05 Then Duration_ms = 0.05
        If Duration_ms > 4000 Then Duration_ms = 4000
        Width_s = Duration_ms / 1000 'Load expects pulse width in s

        _Visa.SendString("CHANNEL:LOAD " & _ChannelNo)
        _Visa.SendString("INPUT ON")
        _Visa.SendString("TRAN ON")
        _Visa.SendString("TRAN:MODE CONT")
        _Visa.SendString("CURR " & CCInitial)
        _Visa.SendString("CURR:TLEV " & CurrPulse)
        _Visa.SendString("CURR:SLEW " & 1000000 * SlewRate)
        _Visa.SendString("TRAN:TWID " & Width_s)
        _Visa.SendString("TRIG:IMM")
        cHelper.Delay(Width_s + 0.1)
        _Visa.SendString("TRAN OFF")
    End Sub

    Public Sub SetOutputON() Implements ILoad.SetOutputON
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("LOAD OFF")
    End Sub

    Public Sub SetShortON() Implements ILoad.SetShortON
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("LOAD ON")
        _Visa.SendString("LOAD:SHOR ON")
    End Sub

    Public Sub SetShortOFF() Implements ILoad.SetShortOFF
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("LOAD:SHOR OFF")
    End Sub

    Public Sub SetModeCC() Implements ILoad.SetModeCC
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("MODE CCH")
    End Sub

    Public Sub SetModeCR() Implements ILoad.SetModeCR
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("MODE CRL")
    End Sub

    Public Sub SetModeCV() Implements ILoad.SetModeCV
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("MODE CV")
    End Sub

    Public Sub SetModeCP() Implements ILoad.SetModeCP
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("CURR:DYN:RISE " & SlewRateRise)
        _Visa.SendString("CURR:DYN:FALL " & SlewRateFall)
    End Sub

    Public Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("RES:RISE " & SlewRateRise)
        _Visa.SendString("RES:FALL " & SlewRateFall)
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

#Region "Public Special Functions Chroma 6314"
    Private Sub ClearProt()
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("LOAD:PROT:CLE")
    End Sub

    Private Function GetStatus() As Single
        _Visa.SendString("CHAN " & _ChannelNo)
        _Visa.SendString("STAT:QUES:COND?")
        Return _Visa.ReceiveString
    End Function
#End Region

End Class
