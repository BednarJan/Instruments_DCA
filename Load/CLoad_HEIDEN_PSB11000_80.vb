'Class CLoad_HEIDEN_PSB11000_80
'23.06.2020 JaBe,  inherited from BCLoad 

Imports Ivi.Visa

Public Class CLoad_HEIDEN_PSB11000_80
    Inherits BCLoad
    Implements ILoad

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 1000
        CurrentMax = 80
        PowerMax = 30000
        ChannelsCount = 1
        ChannelNo = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()

        Visa.SendString("*CLS;*RST")
        Visa.SendString("SYST:LOCK ON")
        Visa.SendString("FUNC:GEN:SEL NONE")

        Visa.SendString("SINK:CURRENT:LIMIT:HIGH " & CurrentMax)
        Visa.SendString("SINK:POWER:LIMIT:HIGH " & PowerMax)

    End Sub

#End Region

#Region "Interface ILoad Methodes"
    Public Overrides Sub SetCC(Current As Single) Implements ILoad.SetCC

        If Current > CurrentMax Then Current = CurrentMax

        If Current = 0 Then
            Visa.SendString("OUTP OFF")
        Else
            Visa.SendString("OUTP ON")
        End If
        Visa.SendString("SINK:CURR " & Current)

    End Sub

    Public Overrides Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR

        Visa.SendString("SINK:RES " & Resistance)

    End Sub

    Public Overrides Sub SetCV(Voltage As Single) Implements ILoad.SetCV

        If Voltage > VoltageMax Then Voltage = VoltageMax

        If Voltage = 0 Then
            Visa.SendString("OUT OFF")
        Else
            Visa.SendString("OUT ON")
        End If

        Visa.SendString("SOURCE:VOLT " & Voltage)
        Visa.SendString("SINK:POW  " & PowerMax)

    End Sub

    Public Overrides Sub SetCP(Pow As Single) Implements ILoad.SetCP

        If Pow > PowerMax Then Pow = PowerMax

        If Pow = 0 Then
            Visa.SendString("OUT OFF")
        Else
            Visa.SendString("OUT ON")
        End If

        Visa.SendString("SINK:POW")
        Visa.SendString("POW " & PowerMax)

    End Sub

    Public Overrides Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Overrides Sub SetCCPulse(ByVal CCInitial As Single, ByVal CurrPulse As Single, Duration_ms As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Overrides Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        Visa.SendString("OUT OFF")
    End Sub

    Public Overrides Sub SetOutputON() Implements ILoad.SetOutputON
        Visa.SendString("OUT ON")
    End Sub

    Public Overrides Sub SetShortON() Implements ILoad.SetShortON
        Visa.SendString("SINK:RES MIN")
        Visa.SendString("OUT ON")
    End Sub

    Public Overrides Sub SetShortOFF() Implements ILoad.SetShortOFF
        Visa.SendString("SINK:RES MAX")
        Visa.SendString("OUT OFF")
    End Sub

    Public Overrides Sub SetModeCC() Implements ILoad.SetModeCC
        Visa.SendString("SINK:CURR")
    End Sub

    Public Overrides Sub SetModeCR() Implements ILoad.SetModeCR
        Visa.SendString("SINK:RES")
    End Sub

    Public Overrides Sub SetModeCV() Implements ILoad.SetModeCV
        Visa.SendString("SINK:VOLT")
    End Sub

    Public Overrides Sub SetModeCP() Implements ILoad.SetModeCP
        Visa.SendString("SINK:POW")
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
