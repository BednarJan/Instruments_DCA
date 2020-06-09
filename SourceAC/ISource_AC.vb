#Region "Enums"
Imports Instruments

Public Enum EPhaseMode
    SinglePhase
    ThreePhase
End Enum

Public Enum EPhaseNumber
    All
    Phase1
    Phase2
    Phase3
End Enum

Public Enum ERange
    LOW
    HIGH
    AUTO
End Enum

Public Enum EShape
    A
    B
End Enum

Public Enum EWaveForm ' Define only valid Waveform Names, since they will be used to send to the source as Settings
    SINe
    DST01
    DST02
    DST03
    DST04
    DST05
    DST06
    DST07
    DST08
    DST09
    DST10
    DST11
    DST12
    DST13
    DST14
    DST15
    DST16
    DST17
    DST18
    DST19
    DST20
    DST21
    DST22
    DST23
    DST24
    DST25
    DST26
    DST27
    DST28
    DST29
    DST30
    SQUAre
    CSIN
End Enum

Public Enum EBase
    TIME
    CYCLE
End Enum

#End Region
Public Interface ISource_AC

#Region "Properties (Set & Get)"
    ReadOnly Property VoltageMax As Single
    ReadOnly Property CurrentMax As Single
    ReadOnly Property PowerMax As Single
#End Region

#Region "Methods (Sub & Functions)"
    Sub SetOutputON()
    Sub SetOutputOFF()

    Sub SetSource(Voltage As Single, CurrentLim As Single, Freq As Single, Optional Phase As EPhaseNumber = 0)
    Sub SetVoltage(Voltage As Single, Optional Phase As EPhaseNumber = 0)
    Sub SetCurrentLim(CurrentLim As Single)
    Sub SetFreq(Freq As Single)

    Sub SetRange(Range As ERange)

    Sub SetShape(Shape As EShape)
    Sub SetWaveForm(Shape As EShape, Waveform As EWaveForm)

    Sub SetStep(VoltageAC As Single, VoltageDC As Single, Frequency As Single, Shape As EShape, StartAngle As Single,
                VoltageStepAC As Single, VoltageStepDC As Single, FrequencyStep As Single, StepDwell As Single, Count As Integer)

    Sub SetList(Count As UInteger, Dwell() As Single, Shape() As EShape, Base As EBase,
                VoltageStartAC() As Single, VoltageEndAC() As Single, VoltageStartDC() As Single, VoltageEndDC() As Single,
                FrequencyStart() As Single, FrequencyEnd() As Single, StartAngle() As Single)

    Sub SetPulse(VoltageAC As Single, VoltageDC As Single, Frequency As Single, Shape As EShape,
                 StartAngle As Single, Count As UInteger, DutyCycle As Single, Duration As Single)

    Sub SetBrownOut(VoltageStartAC As Single, VoltageEndAC As Single, Duration As Single, Frequency As Single)

    Sub ClearProt()
    Function GetStatus() As Single
#End Region

End Interface
