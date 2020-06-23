'Class CLoad_HK_ZSACRV_9826
'23.06.2020 JaBe,  inherited from BCLoad 

Imports Ivi.Visa

Public Class CLoad_HK_ZSACRV_9826
    Inherits BCLoad
    Implements ILoad

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 260
        CurrentMax = 70
        PowerMax = 9800
        ChannelsCount = 1
        ChannelNo = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()

        Visa.SendString("*CLS;*RST")
        Visa.SendString("INP OFF")
        Visa.SendString("SYST:FUNC AC")

    End Sub

#End Region

#Region "Interface ILoad Methodes"
    Public Overrides Sub SetCC(Current As Single) Implements ILoad.SetCC

        If Current > CurrentMax Then Current = CurrentMax

        If Current = 0 Then
            Visa.SendString("INPUT OFF")
        Else
            Visa.SendString("INPUT ON")
        End If
        Visa.SendString("MODE:CURR")
        Visa.SendString("CURR " & Current)

    End Sub

    Public Overrides Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR

        Visa.SendString("MODE:RES")
        Visa.SendString("RES " & Resistance)

    End Sub

    Public Overrides Sub SetCV(Voltage As Single) Implements ILoad.SetCV

        If Voltage > VoltageMax Then Voltage = VoltageMax

        If Voltage = 0 Then
            Visa.SendString("INPUT OFF")
        Else
            Visa.SendString("INPUT ON")
        End If

        Visa.SendString("MODE:VOLT")
        Visa.SendString("VOLT " & Voltage)
    End Sub

    Public Overrides Sub SetCP(Pow As Single) Implements ILoad.SetCP

        If Pow > PowerMax Then Pow = PowerMax

        Visa.SendString("LOAD:MODE POW")
        Visa.SendString("POW:HIGH " & PowerMax)
        Visa.SendString("POW:LEV " & Pow)

    End Sub

    Public Overrides Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        'Applicable for single channel only
        If CC1 > CurrentMax Then CC1 = CurrentMax
        If CC2 > CurrentMax Then CC2 = CurrentMax

        Visa.SendString("CHANNEL:LOAD " & GetChannel(ChannelNo))
        Visa.SendString("TRAN OFF")
        Visa.SendString("INPUT OFF")
        Visa.SendString("MODE:CURR")
        Visa.SendString("CURR " & CC1)
        Visa.SendString("CURR:TLEV " & CC2)
        Visa.SendString("CURR:SLEW " & 1000000 * SlewRate)
        Visa.SendString("TRAN:MODE CONT")
        Visa.SendString("TRAN:FREQ " & Freq)
        Visa.SendString("TRAN:DCYC 50")
        Visa.SendString("TRAN ON")
        Visa.SendString("INPUT ON")
    End Sub

    Public Overrides Sub SetCCPulse(ByVal CCInitial As Single, ByVal CurrPulse As Single, Duration_ms As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Dim Channel As Integer = 0
        Dim Width_s As Single

        If CurrPulse > CurrentMax Then CurrPulse = CurrentMax
        If Duration_ms < 0.05 Then Duration_ms = 0.05
        If Duration_ms > 4000 Then Duration_ms = 4000
        Width_s = Duration_ms / 1000 'Load expects pulse width in s

        Visa.SendString("CHANNEL:LOAD " & GetChannel(ChannelNo))
        Visa.SendString("INPUT ON")
        Visa.SendString("TRAN ON")
        Visa.SendString("TRAN:MODE CONT")
        Visa.SendString("CURR " & CCInitial)
        Visa.SendString("CURR:TLEV " & CurrPulse)
        Visa.SendString("CURR:SLEW " & 1000000 * SlewRate)
        Visa.SendString("TRAN:TWID " & Width_s)
        Visa.SendString("TRIG:IMM")
        cHelper.Delay(Width_s + 0.1)
        Visa.SendString("TRAN OFF")
    End Sub

    Public Overrides Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        MyBase.SetOutputOFF()
    End Sub

    Public Overrides Sub SetOutputON() Implements ILoad.SetOutputON
        MyBase.SetOutputON()
    End Sub

    Public Overrides Sub SetShortON() Implements ILoad.SetShortON
        Visa.SendString("SCIR ON")
    End Sub

    Public Overrides Sub SetShortOFF() Implements ILoad.SetShortOFF
        Visa.SendString("SCIR OFF")
    End Sub

    Public Overrides Sub SetModeCC() Implements ILoad.SetModeCC
        Visa.SendString("MODE:CURR")
    End Sub

    Public Overrides Sub SetModeCR() Implements ILoad.SetModeCR
        Visa.SendString("MODE:RES")
    End Sub

    Public Overrides Sub SetModeCV() Implements ILoad.SetModeCV
        Visa.SendString("MODE:VOLT")
    End Sub

    Public Overrides Sub SetModeCP() Implements ILoad.SetModeCP
        Visa.SendString("MODE:POW")
    End Sub

    Public Overrides Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Overrides Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

#End Region

#Region "Public Special Functions HP 6050A"


#End Region

End Class
