'Class CLoad_CHROMA_63210E_600_700
'16.03.2019, A. Zahler
'Compatible Instruments:
'- Chroma 63210E-600-700 (single channel)
Imports Ivi.Visa

Public Class CLoad_CHROMA_63210E_600_700
    Implements IDevice
    Implements ILoad

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMax As Single = 600 Implements ILoad.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 700 Implements ILoad.CurrentMax
    Public ReadOnly Property PowerMax As Single = 10000 Implements ILoad.PowerMax
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
        _Visa.SendString("*RST;*CLS")
        _Visa.SendString("LOAD OFF")
        _Visa.SendString("MODE CRH")
        _Visa.SendString("RES:L1 6")
        _Visa.SendString("RES:L2 6")
        _Visa.SendString("RES:RISE 1")
        _Visa.SendString("RES:FALL 1")
        _Visa.SendString("MODE CVH")
        _Visa.SendString("VOLT:MODE SLOW")
        _Visa.SendString("VOLT:STAT:L1 600")
        _Visa.SendString("VOLT:STAT:L2 600")
        _Visa.SendString("VOLT:STAT:ILIM MAX")
        _Visa.SendString("MODE CCH")
        _Visa.SendString("CURR:STAT:L1 0")
        _Visa.SendString("CURR:STAT:L2 0")
        _Visa.SendString("CURR:STAT:RISE 1")
        _Visa.SendString("CURR:STAT:FALL 1")
        _Visa.SendString("CONG:VOLT:ON 10")
        _Visa.SendString("LOAD ON")
    End Sub
#End Region

#Region "Interface ILoad Methodes"
    Public Sub SetCC(Current As Single) Implements ILoad.SetCC

        If Current > _CurrentMax Then Current = _CurrentMax

        If Current = 0 Then
            _Visa.SendString("LOAD OFF")
            _Visa.SendString("CURR:STAT:L1 0")
        Else
            _Visa.SendString("MODE CCH")
            _Visa.SendString("CURR:STAT:L1 " & Current)
            _Visa.SendString("LOAD ON")
        End If
    End Sub

    Public Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR
        _Visa.SendString("MODE CR")
        _Visa.SendString("RES:STAT:L1 " & Resistance)
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetCV(Voltage As Single) Implements ILoad.SetCV

        If Voltage > _VoltageMax Then Voltage = _VoltageMax

        _Visa.SendString("MODE CVH")
        _Visa.SendString("VOLT:STAT:ILIM MAX")
        _Visa.SendString("VOLT:STAT:L1 " & Voltage)
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetCP(Power As Single) Implements ILoad.SetCP

        If Power > _PowerMax Then Power = _PowerMax

        _Visa.SendString("MODE CPH")
        _Visa.SendString("POW:STAT:L1 " & Power)
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn

        If CC1 > _CurrentMax Then CC1 = _CurrentMax
        If CC2 > _CurrentMax Then CC2 = _CurrentMax

        _Visa.SendString("LOAD OFF")
        _Visa.SendString("MODE CCDH")
        _Visa.SendString("CURR:DYN:L1 " & CC1)
        '_Visa.SendString("MODE CCDH") 'Required??
        _Visa.SendString("CURR:DYN:L2 " & CC2)
        _Visa.SendString("CURR:DYN:T1 " & ((1 / Freq) / 2))
        _Visa.SendString("CURR:DYN:T2 " & ((1 / Freq) / 2))
        _Visa.SendString("CURR:DYN:RISE " & SlewRate)
        _Visa.SendString("CURR:DYN:FALL " & SlewRate)
        _Visa.SendString("LOAD ON")
    End Sub

    'Public Sub SetCRDyn(CR1 As Single, CR2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCRDyn
    'NOT supported from instrument
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetCVDyn(CV1 As Single, CV2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCVDyn
    'NOT supported from instrument
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetCPDyn(CP1 As Single, CP2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCPDyn
    'NOT supported from instrument
    '    Throw New NotImplementedException()
    'End Sub

    Public Sub SetCCPulse(ByVal CCInitial As Single, ByVal CCPulse As Single, Duration_s As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        _Visa.SendString("LOAD OFF")
        _Visa.SendString("MODE CCDH")
        _Visa.SendString("CURR:DYN:L1 " & CCInitial)
        _Visa.SendString("CURR:DYN:L2 " & CCPulse)
        _Visa.SendString("CURR:DYN:T1 1")
        _Visa.SendString("CURR:DYN:T2 " & Duration_s * 1000 & "ms")
        _Visa.SendString("CURR:DYN:RISE " & SlewRate)
        _Visa.SendString("CURR:DYN:FALL " & SlewRate)
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetOutputON() Implements ILoad.SetOutputON
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        _Visa.SendString("LOAD OFF")
    End Sub

    Public Sub SetShortON() Implements ILoad.SetShortON
        _Visa.SendString("LOAD ON")
        _Visa.SendString("LOAD:SHOR ON")
    End Sub

    Public Sub SetShortOFF() Implements ILoad.SetShortOFF
        _Visa.SendString("LOAD:SHOR OFF")
    End Sub

    Public Sub SetModeCC() Implements ILoad.SetModeCC
        _Visa.SendString("MODE CCH")
    End Sub

    Public Sub SetModeCR() Implements ILoad.SetModeCR
        _Visa.SendString("MODE CRH")
    End Sub

    Public Sub SetModeCV() Implements ILoad.SetModeCV
        _Visa.SendString("MODE CVH")
    End Sub

    Public Sub SetModeCP() Implements ILoad.SetModeCP
        _Visa.SendString("MODE CPH")
    End Sub

    Public Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        _Visa.SendString("CURR:STAT:RISE " & SlewRateRise)
        _Visa.SendString("CURR:STAT:FALL " & SlewRateFall)
    End Sub

    Public Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        _Visa.SendString("RES:RISE " & SlewRateRise)
        _Visa.SendString("RES:FALL " & SlewRateFall)
    End Sub

    'Public Sub SetRangeAuto() Implements ILoad.SetRangeAuto
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetRangeHigh() Implements ILoad.SetRangeHigh
    '        _Visa.SendString("CONF:VOLT:RANG: H")
    'End Sub

    'Public Sub SetRangeLow() Implements ILoad.SetRangeLow
    '        _Visa.SendString("CONF:VOLT:RANG: L")
    'End Sub
#End Region

#Region "Public Special Functions CHROMA 63210E-600-700"
    Private Sub ClearProt()
        _Visa.SendString("LOAD:PROT:CLE")
    End Sub

    Private Function GetStatus() As Single
        _Visa.SendString("STAT:QUES:COND?")
        Return _Visa.ReceiveString
    End Function
#End Region

End Class
