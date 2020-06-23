'Class CLoad_HP_6050A
'10.04.2019, A. Zahler
'23.06.2020 JaBe,  inherited from BCLoad 

Imports Ivi.Visa

Public Class CLoad_HP_6050A
    Inherits BCLoad
    Implements ILoad

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 150
        CurrentMax = 60
        PowerMax = 500
        ChannelsCount = 3
        ChannelNo = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()

        Visa.SendString("*CLS;*RST")

        For Channel As Integer = 1 To ChannelsCount
            Visa.SendString("CHANNEL:LOAD " & Channel)
            Visa.SendString("MODE:CURR")
            Visa.SendString("CURR:RANGE MAX")
            Visa.SendString("INP:STAT OFF")
        Next Channel

    End Sub
#End Region

#Region "Interface ILoad Methodes"
    Public Overrides Sub SetCC(Current As Single) Implements ILoad.SetCC

        If Current > CurrentMax Then Current = CurrentMax
        Visa.SendString("CHANNEL:LOAD " & GetChannel(ChannelNo))
        Visa.SendString("TRAN OFF")
        If Current = 0 Then
            Visa.SendString("INPUT OFF")
        Else
            Visa.SendString("INPUT ON")
        End If
        Visa.SendString("MODE:CURR")
        Visa.SendString("CURR:TRIG " & Current)
        Visa.SendString("TRIG:IMM")
    End Sub

    Public Overrides Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR

        Visa.SendString("CHANNEL:LOAD " & GetChannel(ChannelNo))
        Visa.SendString("TRAN OFF")
        Visa.SendString("INPUT ON")
        Visa.SendString("MODE:RES")
        Visa.SendString("RES:RANGE " & Resistance)
        Visa.SendString("RES:TRIG " & Resistance)
        Visa.SendString("TRIG:IMM")

    End Sub

    Public Overrides Sub SetCV(Voltage As Single) Implements ILoad.SetCV
        'Applicable for single channel only
        Visa.SendString("CHANNEL:LOAD " & GetChannel(ChannelNo))
        Visa.SendString("TRAN OFF")
        Visa.SendString("MODE:VOLT")
        Visa.SendString("VOLT " & Voltage)
        Visa.SendString("INPUT ON")
    End Sub

    Public Overrides Sub SetCP(Voltage As Single) Implements ILoad.SetCP
        Throw New NotImplementedException() 'NOT supported from instrument
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
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("INPUT OFF")
    End Sub

    Public Overrides Sub SetOutputON() Implements ILoad.SetOutputON
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("INPUT ON")
    End Sub

    Public Overrides Sub SetShortON() Implements ILoad.SetShortON
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString(":INPUT:SHOR ON")
    End Sub

    Public Overrides Sub SetShortOFF() Implements ILoad.SetShortOFF
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString(":INPUT:SHOR OFF")
    End Sub

    Public Overrides Sub SetModeCC() Implements ILoad.SetModeCC
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Overrides Sub SetModeCR() Implements ILoad.SetModeCR
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Overrides Sub SetModeCV() Implements ILoad.SetModeCV
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Overrides Sub SetModeCP() Implements ILoad.SetModeCP
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Overrides Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        'Note: Slew rate rise and fall cannot be separately set

        Visa.SendString("CHAN " & GetChannel(ChannelNo))
        Visa.SendString("CURR:SLEW " & 1000000 * SlewRateRise)

    End Sub

    Public Overrides Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        'Note: Slew rate rise and fall cannot be separately set
        'Note: Same command for slew rate CC and CR
        Visa.SendString("CHAN " & GetChannel(ChannelNo))
        Visa.SendString("RES:SLEW " & 1000000 * SlewRateRise)

    End Sub

#End Region

#Region "Public Special Functions HP 6050A"


#End Region

End Class
