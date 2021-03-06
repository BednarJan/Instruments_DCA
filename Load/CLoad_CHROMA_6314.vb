﻿'Class CLoad_CHROMA_6314
'11.04.2019, A. Zahler
'Single channel control implemented only (no parallel operation of modules)
'23.06.2020 JaBe,  inherited from BCLoad 

Imports Ivi.Visa

Public Class CLoad_CHROMA_6314or
    Inherits BCLoad
    Implements ILoad

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)

        VoltageMax = 80
        CurrentMax = 240
        PowerMax = 1200
        ChannelsCount = 8
        ChannelNo = 1

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()

        Visa.SendString("*RST;*CLS")

        For Channel As Integer = 1 To ChannelsCount
            Visa.SendString("CHAN " & Channel)
            Visa.SendString("LOAD OFF")
            Visa.SendString("MODE CRH")
            Visa.SendString("RES:L1 100")
            Visa.SendString("RES:L2 100")
            Visa.SendString("RES:RISE 1")
            Visa.SendString("RES:FALL 1")
            Visa.SendString("MODE CV")
            Visa.SendString("VOLT:CURR MAX")
            Visa.SendString("MODE CCH")
            Visa.SendString("CURR:STAT:L1 0")
            Visa.SendString("CURR:STAT:L2 0")
            Visa.SendString("CURR:STAT:RISE 1")
            Visa.SendString("CURR:STAT:FALL 1")
            Visa.SendString("LOAD OFF")
        Next Channel
    End Sub

#End Region

#Region "Interface ILoad Methodes"
    Public Overrides Sub SetCC(Current As Single) Implements ILoad.SetCC
        If Current > CurrentMax Then Current = CurrentMax

        Visa.SendString("CHAN " & GetChannel(ChannelNo))
        Visa.SendString("MODE CCH")
        Visa.SendString("CURR:STAT:L1 " & Current)

        If Current = 0 Then
            Visa.SendString("LOAD OFF")
        Else
            Visa.SendString("LOAD ON")
        End If
    End Sub

    Public Overrides Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR
        Visa.SendString("CHAN " & GetChannel(ChannelNo))
        Visa.SendString("MODE CRL")
        Visa.SendString("RES:L1 " & Resistance)
        Visa.SendString("LOAD ON")
    End Sub

    Public Overrides Sub SetCV(Voltage As Single) Implements ILoad.SetCV
        'Applicable for single channel only
        Visa.SendString("CHAN " & GetChannel(ChannelNo))
        Visa.SendString("MODE CV")
        Visa.SendString("VOLT:L1 " & Voltage)
        Visa.SendString("LOAD ON")
    End Sub

    Public Overrides Sub SetCP(Val As Single) Implements ILoad.SetCP
        MyBase.SetCP(Val)
    End Sub


    Public Overrides Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        If CC1 > CurrentMax Then CC1 = CurrentMax
        If CC2 > CurrentMax Then CC2 = CurrentMax

        Visa.SendString("CHAN " & GetChannel(ChannelNo))
        Visa.SendString("LOAD OFF")
        Visa.SendString("MODE CCDH")
        Visa.SendString("CURR:DYN:L1 " & CC1)
        Visa.SendString("CURR:DYN:L2 " & CC2)
        Visa.SendString("CURR:DYN:T1 " & (1 / Freq) / 2)
        Visa.SendString("CURR:DYN:T2 " & (1 / Freq) / 2)
        Visa.SendString("CURR:DYN:RISE " & SlewRate)
        Visa.SendString("CURR:DYN:FALL " & SlewRate)
        Visa.SendString("LOAD ON")
    End Sub

    Public Overrides Sub SetCCPulse(ByVal CCInitial As Single, ByVal CurrPulse As Single, Duration_ms As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Dim Width_s As Single

        If CurrPulse > CurrentMax Then CurrPulse = CurrentMax
        If Duration_ms < 0.05 Then Duration_ms = 0.05
        If Duration_ms > 4000 Then Duration_ms = 4000
        Width_s = Duration_ms / 1000 'Load expects pulse width in s

        Visa.SendString("CHAN " & GetChannel(ChannelNo))
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

    Public Overrides Sub SetOutputON() Implements ILoad.SetOutputON
        MyBase.SetOutputON()
    End Sub

    Public Overrides Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        MyBase.SetOutputOFF()
    End Sub

    Public Overrides Sub SetShortON() Implements ILoad.SetShortON
        MyBase.SetShortON()
    End Sub

    Public Overrides Sub SetShortOFF() Implements ILoad.SetShortOFF
        MyBase.SetShortOFF()
    End Sub

    Public Overrides Sub SetModeCC() Implements ILoad.SetModeCC
        MyBase.SetModeCC()
    End Sub

    Public Overrides Sub SetModeCR() Implements ILoad.SetModeCR
        MyBase.SetModeCR()
    End Sub

    Public Overrides Sub SetModeCV() Implements ILoad.SetModeCV
        MyBase.SetModeCV()
    End Sub

    Public Overrides Sub SetModeCP() Implements ILoad.SetModeCP
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC

    End Sub

    Public Overrides Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR

    End Sub

#End Region

#Region "Public Special Functions "
    Overrides Function GetChannel(ByVal chan As Integer)

        If chan = 1 Then Return 1 Else Return 2 * chan - 1

    End Function

#End Region

End Class
