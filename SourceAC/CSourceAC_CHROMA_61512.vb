﻿'Class CSourceAC_CHROMA_61512
'14.04.2019, A. Zahler
'16.04.2020, JaBe  ..........  must be actualised
'Compatible Instruments:
'- Chroma 6530 (single phase)
Imports Ivi.Visa

Public Class CSourceAC_CHROMA_61512
    Inherits BCSourceAC_CHROMA
    Implements ISource_AC

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

        CurrentMaxHi = 30
        CurrentMaxLo = 30

        VoltageMaxHi = 300
        VoltageMaxLo = 150

        PowerMaxHi = 18000
        PowerMaxLo = 18000

        Initialize()

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        MyBase.Initialize()

    End Sub
#End Region

#Region "Interface Methodes ISource_AC"
    Public Overrides Sub SetOutputON() Implements ISource_AC.SetOutputON
        MyBase.SetOutputON()
    End Sub

    Public Overrides Sub SetOutputOFF() Implements ISource_AC.SetOutputOFF
        MyBase.SetOutputOFF()
    End Sub

    Public Overrides Sub SetVoltage(Voltage As Single, CurrentLim As Single, Optional Phase As ISource_AC.EPhaseNumber = ISource_AC.EPhaseNumber.All, Optional SetOutON As Boolean = True) Implements ISource_AC.SetVoltage
        MyBase.SetVoltage(Voltage, CurrentLim, ISource_AC.EPhaseNumber.All, SetOutON)
    End Sub

    Public Overrides Sub SetVoltage(Voltage As Single, Optional SetOutON As Boolean = True) Implements ISource_AC.SetVoltage
        MyBase.SetVoltage(Voltage, SetOutON)
    End Sub

    Public Overrides Sub SetVoltPuls(ByVal vStart As Single, ByVal vPulse As Single, ByVal Width As Single, Optional ByVal phase As Single = 90) Implements ISource_AC.SetVoltPuls

        MyBase.SetVoltPuls(vStart, vPulse, Width, phase)

    End Sub

    Public Overrides Sub SetCurrentLim(CurrentLim As Single) Implements ISource_AC.SetCurrentLim
        MyBase.SetCurrentLim(CurrentLim)
    End Sub

    Public Overrides Sub SetFreq(Freq As Single) Implements ISource_AC.SetFreq
        MyBase.SetFreq(Freq)
    End Sub

    Public Overrides Sub SetRange(Range As ISource_AC.ERange) Implements ISource_AC.SetRange

        MyBase.SetRange(Range)
    End Sub

    Public Overrides Sub SetPhaseMode(PhaseMode As ISource_AC.EPhaseMode) Implements ISource_AC.SetPhaseMode

        MyBase.SetPhaseMode(PhaseMode)

        Select Case PhaseMode
            Case ISource_AC.EPhaseMode.SinglePhase
                CurrentMaxHi = 90
            Case ISource_AC.EPhaseMode.ThreePhase
                CurrentMaxHi = 30

        End Select

    End Sub

    Public Overrides Function GetStatus() As Single Implements ISource_AC.GetStatus

        Return MyBase.GetStatus

    End Function




#End Region

#Region "Public Special Functions CHROMA 61512"



#End Region

End Class
