'Class CLoad_PRODIGIT_3311
'10.04.2019, A. Zahler
'Compatible Instruments:
'- Prodigit mainframe 3302 and load module 3311 (single channel)
Imports Ivi.Visa

Public Class CLoad_PRODIGIT_3311
    Implements IDevice
    Implements ILoad

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As Integer

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMax As Single = 60 Implements ILoad.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 30 Implements ILoad.CurrentMax
    Public ReadOnly Property PowerMax As Single = 300 Implements ILoad.PowerMax
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Visa = New CVisaDeviceNI(Session, ErrorLogger)
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
        _Visa.SendString("CLER")
        _Visa.SendString("LOAD OFF")
        _Visa.SendString("SHOR OFF")
        _Visa.SendString("DYN OFF")
        _Visa.SendString("MODE CC")
        _Visa.SendString("RANG HIGH")
        _Visa.SendString("CURR:HIGH 0")
        _Visa.SendString("LEV HIGH")
        _Visa.SendString("LOAD ON")
    End Sub
#End Region

#Region "interface methodes"
    Public Sub SetCC(Current As Single) Implements ILoad.SetCC
        If Current = 0 Then
            _Visa.SendString("LOAD OFF")
            _Visa.SendString("DYN OFF")
        Else
            If Current <= 6 Then
                _Visa.SendString("DYN OFF")
                _Visa.SendString("MODE CC")
                _Visa.SendString("RANG LOW")
                _Visa.SendString("LEV HIGH")
                _Visa.SendString("CURR:LOW 0")
                _Visa.SendString("CURR:HIGH " & Current)
                _Visa.SendString("LOAD ON")
            Else
                _Visa.SendString("DYN OFF")
                _Visa.SendString("MODE CC")
                _Visa.SendString("RANG HIGH")
                _Visa.SendString("LEV HIGH")
                _Visa.SendString("CURR:LOW 0")
                _Visa.SendString("CURR:HIGH " & Current)
                _Visa.SendString("LOAD ON")
            End If
        End If
    End Sub

    Public Sub SetCR(Val As Single) Implements ILoad.SetCR
        Throw New NotImplementedException()
    End Sub

    Public Sub SetCV(Val As Single) Implements ILoad.SetCV
        Throw New NotImplementedException()
    End Sub

    Public Sub SetCP(Val As Single) Implements ILoad.SetCP
        Throw New NotImplementedException()
    End Sub

    Public Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        _Visa.SendString("LOAD OFF")
        _Visa.SendString("MODE CC")
        _Visa.SendString("RANG HIGH")
        _Visa.SendString("DYN ON")
        _Visa.SendString("CURR:HIGH " & CC2)
        _Visa.SendString("CURR:LOW " & CC1)
        _Visa.SendString("RISE: " & SlewRate)
        _Visa.SendString("FALL: " & SlewRate)
        _Visa.SendString("PERD:HIGH " & ((1 / Freq) / 2) * 1000)
        _Visa.SendString("PERD:LOW " & ((1 / Freq) / 2) * 1000)
        _Visa.SendString("LOAD ON")
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

    Public Sub SetCCPulse(ByVal Curr As Single, ByVal CurrPulse As Single, Duration_ms As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Throw New NotImplementedException()
    End Sub

    Public Sub SetOutputON() Implements ILoad.SetOutputON
        _Visa.SendString("LOAD ON")
    End Sub

    Public Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        _Visa.SendString("LOAD OFF")
    End Sub

    Public Sub SetShortON() Implements ILoad.SetShortON
        _Visa.SendString("LOAD ON")
        _Visa.SendString("SHOR ON")
    End Sub

    Public Sub SetShortOFF() Implements ILoad.SetShortOFF
        _Visa.SendString("SHOR OFF")
    End Sub

    Public Sub SetModeCC() Implements ILoad.SetModeCC
        _Visa.SendString("MODE CC")
    End Sub

    Public Sub SetModeCR() Implements ILoad.SetModeCR
        _Visa.SendString("MODE CR")
    End Sub

    Public Sub SetModeCV() Implements ILoad.SetModeCV
        _Visa.SendString("MODE CV")
    End Sub

    Public Sub SetModeCP() Implements ILoad.SetModeCP
        Throw New NotImplementedException() 'Not supported from this load
    End Sub

    'Public Sub SetRangeAuto() Implements ILoad.SetRangeAuto
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetRangeHigh() Implements ILoad.SetRangeHigh
    '    Throw New NotImplementedException()
    'End Sub

    'Public Sub SetRangeLow() Implements ILoad.SetRangeLow
    '    Throw New NotImplementedException()
    'End Sub

    Public Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        _Visa.SendString("RISE " & SlewRateRise)
        _Visa.SendString("FALL " & SlewRateFall)
    End Sub

    Public Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        _Visa.SendString("RISE " & SlewRateRise)
        _Visa.SendString("FALL " & SlewRateFall)
    End Sub
#End Region

#Region "Public Special Functions Prodigit 3311"
    Private Sub ClearProt()
        _Visa.SendString("CLER")
    End Sub

    Private Function GetStatus() As Single
        _Visa.SendString("PROT?")
        Return _Visa.ReceiveString
    End Function
#End Region

End Class
