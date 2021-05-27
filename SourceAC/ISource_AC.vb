Imports Instruments

Public Interface ISource_AC
    Inherits IDevice

#Region "Enum definitions"
    Enum EPhaseMode
        SinglePhase = 1
        TwoPhase = 2
        ThreePhase = 3
    End Enum

    Enum EPhaseNumber
        All
        Phase1
        Phase2
        Phase3
    End Enum

    Enum ERange
        LOW
        HIGH
        AUTO
    End Enum

    Enum EShape
        A
        B
    End Enum

    Enum EBase
        TIME
        CYCLE
    End Enum

#End Region

#Region "Properties (Set & Get)"
    ReadOnly Property VoltageMax As Single
    ReadOnly Property CurrentMax As Single
    ReadOnly Property PowerMax As Single
#End Region

#Region "Methods (Sub & Functions)"
    Sub SetOutputON()
    Sub SetOutputOFF()

    Sub SetVoltage(Voltage As Single, CurrentLim As Single, Optional Phase As EPhaseNumber = 0, Optional SetOutON As Boolean = True)
    Sub SetVoltage(Voltage As Single, Optional SetON As Boolean = True)

    Sub SetCurrentLim(CurrentLim As Single)
    Sub SetFreq(Freq As Single)

    Sub SetRange(Range As ERange)

    Sub SetPhaseMode(Mode As EPhaseMode)

    Sub SetVoltPuls(ByVal vStart As Single, ByVal vPulse As Single, ByVal Width As Single, Optional ByVal phase As Single = 90)

    Function GetStatus() As Single

    Function GetVolt(Optional nPhase As Integer = 1) As Single

    Function GetCurrent(Optional nPhase As Integer = 1) As Single

    Function GetPower(Optional nPhase As Integer = 1) As Single

#End Region

End Interface
