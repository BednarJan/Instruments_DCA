Public Interface ILoad

#Region "Properties"

    ReadOnly Property VoltageMax As Single
    ReadOnly Property CurrentMax As Single
    ReadOnly Property PowerMax As Single
    Property ChannelsCount As Byte

#End Region

#Region "Methods (Sub & Functions)"
    Sub SetCC(Current As Single)
    Sub SetCR(ByVal Resistance As Single)
    Sub SetCV(Voltage As Single)
    Sub SetCP(Pow As Single)

    Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single)

    Sub SetCCPulse(ByVal CCInitial As Single, ByVal CCPulse As Single, Duration_s As Single, Optional SlewRate As Single = 1)

    Sub SetOutputON()
    Sub SetOutputOFF()

    Sub SetShortON()
    Sub SetShortOFF()

    Sub SetModeCC()
    Sub SetModeCV()
    Sub SetModeCR()
    Sub SetModeCP()

    Sub SetRangeLow()  'Required?
    Sub SetRangeHigh() 'Required?
    Sub SetRangeAuto() 'Required?

    Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single)
    Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single)

#End Region
End Interface
