Public Interface ILoad

#Region "Properties"
    ReadOnly Property VoltageMax As Single
    ReadOnly Property CurrentMax As Single
    ReadOnly Property PowerMax As Single
#End Region

#Region "Methods (Sub & Functions)"
    Sub SetCC(Val As Single)
    Sub SetCR(Val As Single)
    Sub SetCV(Val As Single)
    Sub SetCP(Val As Single)

    'Chroma 63200/E/A series and Prodigit 3302 offers dynamic operation in CC mode only
    'HP 6050A offers dynamic (= Transient) operation in CC, CR and CV mode
    'Prodigit 3311 offers dynamic operation in CC mode only
    'Zentro EL3000 offers dynamic operation in CC, CR, CV and CP mode
    Sub SetCCDyn(CC1 As Single, CC2 As Single, Freq As Single, SlewRate As Single)
    'Sub SetCVDyn(CV1 As Single, CV2 As Single, Freq As Single, SlewRate As Single) 'Required?
    'Sub SetCRDyn(CR1 As Single, CR2 As Single, Freq As Single, SlewRate As Single) 'Required?
    'Sub SetCPDyn(CP1 As Single, CP2 As Single, Freq As Single, SlewRate As Single) 'Required?

    Sub SetCCPulse(ByVal CCInitial As Single, ByVal CCPulse As Single, Duration_s As Single, Optional SlewRate As Single = 1)

    Sub SetOutputON()
    Sub SetOutputOFF()

    Sub SetShortON()
    Sub SetShortOFF()

    Sub SetModeCC()
    Sub SetModeCV()
    Sub SetModeCR()
    Sub SetModeCP()

    'Sub SetRangeLow()  'Required?
    'Sub SetRangeHigh() 'Required?
    'Sub SetRangeAuto() 'Required?

    Sub SetSlewRateCC(SlewRateRise As Single, SlewRateFall As Single)
    Sub SetSlewRateCR(SlewRateRise As Single, SlewRateFall As Single)

    'Sub ClearProt()
    'Function GetStatus() As Single

#End Region
End Interface
