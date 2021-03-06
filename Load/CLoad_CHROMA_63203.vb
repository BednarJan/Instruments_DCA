﻿'Class CLoad_CHROMA_63203
'10.02.2019, A. Zahler
'Compatible Instruments:
'- Chroma 63203 (single channel)
'23.06.2020 JaBe,  inherited from BCLoad 

Imports Ivi.Visa

Public Class CLoad_CHROMA_63203
    Inherits BCLoad
    Implements ILoad

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)
        VoltageMax = 80
        CurrentMax = 600
        PowerMax = 5200
        ChannelsCount = 1
        ChannelNo = 1

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()
        MyBase.Initialize()
    End Sub
#End Region

#Region "Interface ILoad Methodes"
    Public Overrides Sub SetCC(Current As Single) Implements ILoad.SetCC
        MyBase.SetCC(Current)
    End Sub

    Public Overrides Sub SetCR(ByVal Resistance As Single) Implements ILoad.SetCR
        MyBase.SetCR(Resistance)
    End Sub

    Public Overrides Sub SetCV(Voltage As Single) Implements ILoad.SetCV
        MyBase.SetCV(Voltage)
    End Sub

    Public Overrides Sub SetCP(Pow As Single) Implements ILoad.SetCP
        MyBase.SetCP(Pow)
    End Sub

    Public Overrides Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        MyBase.SetCCDyn(CC1, CC2, Freq, SlewRate)
    End Sub


    Public Overrides Sub SetCCPulse(ByVal CCInitial As Single, ByVal CCPulse As Single, Duration_s As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        MyBase.SetCCPulse(CCInitial, CCPulse, Duration_s)
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
        Visa.SendString("MODE CVH")
    End Sub

    Public Overrides Sub SetModeCP() Implements ILoad.SetModeCP
        MyBase.SetModeCP()
    End Sub

    Public Overrides Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCC
        Visa.SendString("CURR:STAT:RISE " & SlewRateRise)
        Visa.SendString("CURR:STAT:FALL " & SlewRateFall)
    End Sub

    Public Overrides Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        Visa.SendString("RES:RISE " & SlewRateRise)
        Visa.SendString("RES:FALL " & SlewRateFall)
    End Sub

#End Region

#Region "Public Special Functions CHROMA 63203"

#End Region

End Class
