'Class CLoad_CHROMA_63210E_600_700
'16.03.2019, A. Zahler
'Compatible Instruments:
'- Chroma 63210E-600-700 (single channel)
'23.06.2020,  JaBe   Inherited from BCLoad
Imports Ivi.Visa

Public Class CLoad_CHROMA_63210E_600_700
    Inherits BCLoad
    Implements ILoad

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 600
        CurrentMax = 700
        PowerMax = 10000
        ChannelsCount = 3
        ChannelNo = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

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

    Public Overrides Sub SetCP(Power As Single) Implements ILoad.SetCP
        MyBase.SetCP(Power)
    End Sub

    Public Overrides Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single) Implements ILoad.SetCCDyn
        MyBase.SetCCDyn(CC1, CC2, Freq, SlewRate)
    End Sub

    Public Overrides Sub SetCCPulse(ByVal CCInitial As Single, ByVal CCPulse As Single, Duration_s As Single, Optional SlewRate As Single = 1) Implements ILoad.SetCCPulse
        MyBase.SetCCPulse(CCInitial, CCPulse, Duration_s, SlewRate)
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
        MyBase.SetSlewRateCC(SlewRateRise, SlewRateFall)
    End Sub

    Public Overrides Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single) Implements ILoad.SetSlewRateCR
        MyBase.SetSlewRateCR(SlewRateRise, SlewRateFall)
    End Sub
#End Region

#Region "Public Special Functions CHROMA 63210E-600-700"

#End Region

End Class
