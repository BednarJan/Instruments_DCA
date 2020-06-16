Public Interface ISource_DC
    Inherits IDevice

#Region "Properties (Set & Get)"

    ReadOnly Property VoltageMax As Single
    ReadOnly Property CurrentMax As Single
    ReadOnly Property PowerMax As Single

#End Region

#Region "Methods (Sub & Functions)"

    Sub SetOutputON()
    Sub SetOutputOFF()

    Sub SetVoltage(Voltage As Single, CurrentLim As Single, Optional SetON As Boolean = True)
    Sub SetVoltage(Voltage As Single, Optional SetON As Boolean = True)
    Sub SetVoltPuls(ByVal vPulse As Single, ByVal Width As Single, ByVal vEnd As Single)


    Sub SetCurrentLim(CurrentLim As Single)

#End Region
End Interface
