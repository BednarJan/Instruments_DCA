
'Class BCLoad inherited from the BCDevice
'02.05.2019, J. Bednar  
'23.06.2020, JaBe
'Base clase for the DC sources
Imports Ivi.Visa

Public Class BCLoad
    Inherits BCDevice
    Implements ILoad

#Region "Shorthand Properties"

    Public Property VoltageMax As Single Implements ILoad.VoltageMax
    Public Property CurrentMax As Single Implements ILoad.CurrentMax
    Public Property PowerMax As Single Implements ILoad.PowerMax
    Public Property ChannelsCount As Byte Implements ILoad.ChannelsCount
    Public Property ChannelNo As Byte = 1

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()
        Visa.SendString("*RST;*CLS")
        Visa.SendString("LOAD OFF")
        Visa.SendString("MODE CRH")
        Visa.SendString("RES:L1 6")
        Visa.SendString("RES:L2 6")
        Visa.SendString("RES:RISE 1")
        Visa.SendString("RES:FALL 1")
        Visa.SendString("MODE CVH")
        Visa.SendString("VOLT:MODE SLOW")
        Visa.SendString("VOLT:STAT:L1 " & VoltageMax)
        Visa.SendString("VOLT:STAT:L2 " & VoltageMax)
        Visa.SendString("VOLT:CURR MAX")
        Visa.SendString("MODE CCH")
        Visa.SendString("CURR:STAT:L1 0")
        Visa.SendString("CURR:STAT:L2 0")
        Visa.SendString("CURR:STAT:RISE 1")
        Visa.SendString("CURR:STAT:FALL 1")
        Visa.SendString("LOAD OFF")
    End Sub

#End Region

#Region "Interface ILoad Methodes"

    Public Overridable Sub SetCC(Current As Single) Implements ILoad.SetCC

        If Current > CurrentMax Then Current = CurrentMax

        If Current = 0 Then
            Visa.SendString("LOAD OFF")
            Visa.SendString("CURR:STAT:L1 0")
        Else
            Visa.SendString("MODE CCH")
            Visa.SendString("CURR:STAT:L1 " & Current)
            Visa.SendString("LOAD ON")
        End If

    End Sub

    Public Overridable Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR
        Visa.SendString("MODE CRH")
        Visa.SendString("RES:L1 " & Resistance)
        Visa.SendString("LOAD ON")
    End Sub

    Public Overridable Sub SetCV(Voltage As Single) Implements ILoad.SetCV
        If Voltage > VoltageMax Then Voltage = VoltageMax

        Visa.SendString("MODE CVH")
        Visa.SendString("VOLT:L1 " & Voltage)
        Visa.SendString("LOAD ON")

    End Sub

    Public Overridable Sub SetCP(Pow As Single) Implements ILoad.SetCP
        If Pow > PowerMax Then Pow = PowerMax

        Visa.SendString("MODE CPH")
        Visa.SendString("POW:L1 " & Pow)
        Visa.SendString("LOAD ON")

    End Sub

    Public Overridable Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        If CC1 > CurrentMax Then CC1 = CurrentMax
        If CC2 > CurrentMax Then CC2 = CurrentMax

        Visa.SendString("LOAD OFF")
        Visa.SendString("MODE CCDH")
        Visa.SendString("CURR:DYN:L1 " & CC1)
        '_Visa.SendString("MODE CCDH") 'Required??
        Visa.SendString("CURR:DYN:L2 " & CC2)
        Visa.SendString("CURR:DYN:T1 " & ((1 / Freq) / 2))
        Visa.SendString("CURR:DYN:T2 " & ((1 / Freq) / 2))
        Visa.SendString("CURR:DYN:RISE " & SlewRate)
        Visa.SendString("CURR:DYN:FALL " & SlewRate)
        Visa.SendString("LOAD ON")
    End Sub

    Public Overridable Sub SetCCPulse(ByVal CCInitial As Single, ByVal CCPulse As Single, Duration_s As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        Visa.SendString("LOAD OFF")
        Visa.SendString("MODE CCDH")
        Visa.SendString("CURR:DYN:L1 " & CCInitial)
        Visa.SendString("CURR:DYN:L2 " & CCPulse)
        Visa.SendString("CURR:DYN:T1 1s")
        Visa.SendString("CURR:DYN:T2 " & Duration_s & "s")
        Visa.SendString("CURR:DYN:RISE " & SlewRate)
        Visa.SendString("CURR:DYN:FALL " & SlewRate)
        Visa.SendString("LOAD ON")
    End Sub

    Public Overridable Sub SetOutputON() Implements ILoad.SetOutputON
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("LOAD ON")
    End Sub

    Public Overridable Sub SetOutputOFF() Implements ILoad.SetOutputOFF
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("LOAD OFF")
    End Sub

    Public Overridable Sub SetShortON() Implements ILoad.SetShortON
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("LOAD ON")
        Visa.SendString("LOAD:SHOR ON")
    End Sub

    Public Overridable Sub SetShortOFF() Implements ILoad.SetShortOFF
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("LOAD:SHOR OFF")
    End Sub

    Public Overridable Sub SetModeCC() Implements ILoad.SetModeCC
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("MODE CCH")
    End Sub

    Public Overridable Sub SetModeCV() Implements ILoad.SetModeCV
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("MODE CV")
    End Sub

    Public Overridable Sub SetModeCR() Implements ILoad.SetModeCR
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("MODE CRH")
    End Sub

    Public Overridable Sub SetModeCP() Implements ILoad.SetModeCP
        Visa.SendString("MODE CPH")
    End Sub

    Public Overridable Sub SetRangeLow() Implements ILoad.SetRangeLow
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub SetRangeHigh() Implements ILoad.SetRangeHigh
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub SetRangeAuto() Implements ILoad.SetRangeAuto
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("CURR:DYN:RISE " & SlewRateRise)
        Visa.SendString("CURR:DYN:FALL " & SlewRateFall)
    End Sub

    Public Overridable Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        If ChannelsCount > 1 Then
            Visa.SendString("CHAN " & GetChannel(ChannelNo))
        End If
        Visa.SendString("RES:RISE " & SlewRateRise)
        Visa.SendString("RES:FALL " & SlewRateFall)
    End Sub

#End Region

#Region "private methodes and functions"

    Overridable Function GetChannel(ByVal chan As Integer)

        Return chan

    End Function

    Overridable Sub ClearProt()
        Visa.SendString("LOAD:PROT:CLE")
    End Sub

    Overridable Function GetStatus() As Single
        Visa.SendString("STAT:QUES:COND?")
        Return Visa.ReceiveString()
    End Function


#End Region

End Class
