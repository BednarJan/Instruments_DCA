'Class CLoad_ZENTRO_EL3000
'06.04.2019, A. Zahler
'Compatible Instruments:
'- Zentro EL3000 (single channel)
Imports Ivi.Visa

Public Class CLoad_ZENTRO_EL3000
    Implements IDevice
    Implements ILoad

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMax As Single = 160 Implements ILoad.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 100 Implements ILoad.CurrentMax
    Public ReadOnly Property PowerMax As Single = 3000 Implements ILoad.PowerMax
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
        _Visa.SendString("*RST;*CLS" & vbLf)
        _Visa.SendString("IMODE" & vbLf)
        _Visa.SendString("CURR1 " & Format(0, "00.00") & vbLf)
        _Visa.SendString("LOAD ON" & vbLf)
    End Sub
#End Region

#Region "Interface ILoad Methodes"
    Public Sub SetCC(Current As Single) Implements ILoad.SetCC

        If Current > _CurrentMax Then Current = _CurrentMax

        If Current = 0 Then
            _Visa.SendString("LOAD OFF" & vbLf)
            _Visa.SendString("CURR1 " & Format(0, "00.00") & vbLf)
        Else
            _Visa.SendString("IMODE" & vbLf)
            _Visa.SendString("CURR1 " & Format(Current, "00.00") & vbLf)
            _Visa.SendString("LOAD ON" & vbLf)
        End If
    End Sub

    Public Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR
        Dim Conductance As Single

        Conductance = 1 / Resistance
        If Conductance > 80 Then Conductance = 80

        _Visa.SendString("GMODE" & vbLf)
        _Visa.SendString("COND1 " & Format(Conductance, "00.0000") & vbLf)
        _Visa.SendString("LOAD ON" & vbLf)
    End Sub

    Public Sub SetCV(Voltage As Single) Implements ILoad.SetCV

        If Voltage > _VoltageMax Then Voltage = _VoltageMax

        _Visa.SendString("UMODE" & vbLf)
        _Visa.SendString("VOLT1 " & Format(Voltage, "00.00") & vbLf)
        _Visa.SendString("LOAD ON" & vbLf)
    End Sub

    Public Sub SetCP(Power As Single) Implements ILoad.SetCP

        If Power > _PowerMax Then Power = _PowerMax

        _Visa.SendString("PMODE" & vbLf)
        _Visa.SendString("POW1 " & Format(Power, "00.00") & vbLf)
        _Visa.SendString("LOAD ON" & vbLf)
    End Sub

    Public Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn

        If CC1 > _CurrentMax Then CC1 = _CurrentMax
        If CC2 > _CurrentMax Then CC2 = _CurrentMax

        Throw New NotImplementedException()
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

    Public Sub SetCCPulse(ByVal CCInitial As Single, ByVal CCPulse As Single, Duration_s As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Throw New NotImplementedException()
    End Sub

    'Public Sub SetCRDyn(CR1 As Single, CR2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCRDyn
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetCVDyn(CV1 As Single, CV2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCVDyn
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetDyn(CP1 As Single, CP2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCPDyn
    '    Throw New NotImplementedException()
    'End Sub

    Public Sub SetOutputON() Implements ILoad.SetOutputON
        _Visa.SendString("LOAD ON" & vbLf)
    End Sub

    Public Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        _Visa.SendString("LOAD OFF" & vbLf)
    End Sub

    Public Sub SetShortON() Implements ILoad.SetShortON
        Throw New NotImplementedException() 'Not supported from this load
    End Sub

    Public Sub SetShortOFF() Implements ILoad.SetShortOFF
        Throw New NotImplementedException() 'Not supported from this load
    End Sub

    Public Sub SetModeCC() Implements ILoad.SetModeCC
        'This command sets LOAD OFF
        _Visa.SendString("IMODE" & vbLf)
    End Sub

    Public Sub SetModeCR() Implements ILoad.SetModeCR
        'This command sets LOAD OFF
        _Visa.SendString("GMODE" & vbLf)
    End Sub

    Public Sub SetModeCV() Implements ILoad.SetModeCV
        'This command sets LOAD OFF
        _Visa.SendString("UMODE" & vbLf)
    End Sub

    Public Sub SetModeCP() Implements ILoad.SetModeCP
        'This command sets LOAD OFF
        _Visa.SendString("PMODE" & vbLf)
    End Sub

    Public Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        Throw New NotImplementedException() 'Not supported from this load
    End Sub

    Public Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        Throw New NotImplementedException() 'Not supported from this load
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

#Region "Public Special Functions ZENTRO EL3000"
    Private Sub ClearProt()
        Throw New NotImplementedException()
    End Sub

    Private Function GetStatus() As Single
        _Visa.SendString("STATUS" & vbLf)
        Return _Visa.ReceiveString()
    End Function
#End Region

End Class
