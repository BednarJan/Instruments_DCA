'Class CLoad_CHROMA_63804
'23.06.2020 JaBe,  inherited from BCLoad 

Imports Ivi.Visa

Public Class CLoad_CHROMA_63804
    Inherits BCLoad
    Implements ILoad

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 350
        CurrentMax = 45
        PowerMax = 4500
        ChannelsCount = 1
        ChannelNo = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()

        Visa.SendString("*CLS;*RST")
        Visa.SendString("LOAD OFF")
        Visa.SendString("CURR:MAX " & CurrentMax)
        Visa.SendString("SYSTEM:SETUP:MODE AC")
        Visa.SendString("ABA OFF")

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
        Throw New NotImplementedException() 'NOT supported from instrument
    End Sub

    Public Overrides Sub SetCCPulse(ByVal CCInitial As Single, ByVal CurrPulse As Single, Duration_ms As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Throw New NotImplementedException() 'NOT supported from instrument
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

#Region "Public Special Functions "


#End Region

End Class
