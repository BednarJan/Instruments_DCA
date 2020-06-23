'Class CLoad_ZENTRO_EL3000
'06.04.2019, A. Zahler
'Compatible Instruments:
'- Zentro EL3000 (single channel)

Imports Ivi.Visa

Public Class CLoad_ZENTRO_EL3000
    Inherits BCLoad
    Implements ILoad

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 160
        CurrentMax = 100
        PowerMax = 3000
        ChannelsCount = 3
        ChannelNo = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()
        Visa.SendString("*RST;*CLS" & vbLf)
        Visa.SendString("IMODE" & vbLf)
        Visa.SendString("CURR1 " & Format(0, "00.00") & vbLf)
        Visa.SendString("LOAD OFF" & vbLf)
    End Sub

#End Region

#Region "Interface ILoad Methodes"
    Public Overrides Sub SetCC(Current As Single) Implements ILoad.SetCC

        If Current > CurrentMax Then Current = CurrentMax

        If Current = 0 Then
            Visa.SendString("LOAD OFF" & vbLf)
            Visa.SendString("CURR1 " & Format(0, "00.00") & vbLf)
        Else
            Visa.SendString("IMODE" & vbLf)
            Visa.SendString("CURR1 " & Format(Current, "00.00") & vbLf)
            Visa.SendString("LOAD ON" & vbLf)
        End If
    End Sub

    Public Overrides Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR
        Dim Conductance As Single

        Conductance = 1 / Resistance
        If Conductance > 80 Then Conductance = 80

        Visa.SendString("GMODE" & vbLf)
        Visa.SendString("COND1 " & Format(Conductance, "00.0000") & vbLf)
        Visa.SendString("LOAD ON" & vbLf)
    End Sub

    Public Overrides Sub SetCV(Voltage As Single) Implements ILoad.SetCV

        If Voltage > VoltageMax Then Voltage = VoltageMax

        Visa.SendString("UMODE" & vbLf)
        Visa.SendString("VOLT1 " & Format(Voltage, "00.00") & vbLf)
        Visa.SendString("LOAD ON" & vbLf)
    End Sub

    Public Overrides Sub SetCP(Power As Single) Implements ILoad.SetCP

        If Power > PowerMax Then Power = PowerMax

        Visa.SendString("PMODE" & vbLf)
        Visa.SendString("POW1 " & Format(Power, "00.00") & vbLf)
        Visa.SendString("LOAD ON" & vbLf)
    End Sub

    Public Overrides Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn

        If CC1 > CurrentMax Then CC1 = CurrentMax
        If CC2 > CurrentMax Then CC2 = CurrentMax

        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub SetCCPulse(ByVal CCInitial As Single, ByVal CCPulse As Single, Duration_s As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub SetOutputON() Implements ILoad.SetOutputON
        MyBase.SetOutputON()
    End Sub

    Public Overrides Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        MyBase.SetOutputOFF()
    End Sub

    Public Overrides Sub SetShortON() Implements ILoad.SetShortON
        Throw New NotImplementedException() 'Not supported from this load
    End Sub

    Public Overrides Sub SetShortOFF() Implements ILoad.SetShortOFF
        Throw New NotImplementedException() 'Not supported from this load
    End Sub

    Public Overrides Sub SetModeCC() Implements ILoad.SetModeCC
        'This command sets LOAD OFF
        Visa.SendString("IMODE" & vbLf)
    End Sub

    Public Overrides Sub SetModeCR() Implements ILoad.SetModeCR
        'This command sets LOAD OFF
        Visa.SendString("GMODE" & vbLf)
    End Sub

    Public Overrides Sub SetModeCV() Implements ILoad.SetModeCV
        'This command sets LOAD OFF
        Visa.SendString("UMODE" & vbLf)
    End Sub

    Public Overrides Sub SetModeCP() Implements ILoad.SetModeCP
        'This command sets LOAD OFF
        Visa.SendString("PMODE" & vbLf)
    End Sub

    Public Overrides Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        Throw New NotImplementedException() 'Not supported from this load
    End Sub

    Public Overrides Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        Throw New NotImplementedException() 'Not supported from this load
    End Sub

#End Region

#Region "Public Special Functions "
    Overrides Sub ClearProt()
        Throw New NotImplementedException()
    End Sub

    Overrides Function GetStatus() As Single
        Visa.SendString("STATUS" & vbLf)
        Return Visa.ReceiveString()
    End Function
#End Region

End Class
