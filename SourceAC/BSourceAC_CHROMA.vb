'Base Class BCSourceAC_CHROMA
'26.05.2021, JaBe  ..........  common for all AC=sources
Imports Ivi.Visa

Public Class BCSourceAC_CHROMA
    Inherits BCSourceAC
    Implements ISource_AC

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        Visa.SendString("*RST;*CLS:")
        Visa.SendString("OUTP:PROT:CLE")
        Visa.SendString("RANG LOW")

        VoltageMax = VoltageMaxLo
        CurrentMax = CurrentMaxLo
        PowerMax = PowerMaxLo


        Visa.SendString("SOUR:FREQ:IMM 50")
        Visa.SendString("SOUR:FUNC:SHAP A")
        Visa.SendString("TPH 0")
        Visa.SendString("TPH:SYNC PHAS")
        Visa.SendString("OREL ON")
        Visa.SendString("SOUR:VOLT 0")
        Visa.SendString("OUTP:STAT OFF")
        Visa.SendString("SOUR:FUNC:TRAN OFF")

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

        Select Case Range
            Case ISource_AC.ERange.LOW

                VoltageMax = VoltageMaxLo
                CurrentMax = CurrentMaxLo
                PowerMax = PowerMaxLo

                Visa.SendString("RANG LOW")
            Case ISource_AC.ERange.HIGH

                VoltageMax = VoltageMaxHi
                CurrentMax = CurrentMaxHi
                PowerMax = PowerMaxHi

                Visa.SendString("RANG HIGH")
            Case Else
                Visa.SendString("RANG AUTO")
        End Select
    End Sub

    Public Overrides Sub SetPhaseMode(PhaseMode As ISource_AC.EPhaseMode) Implements ISource_AC.SetPhaseMode
        Dim sPhases As String = String.Empty
        Select Case PhaseMode
            Case ISource_AC.EPhaseMode.SinglePhase
                sPhases = "SINGLE"
            Case ISource_AC.EPhaseMode.ThreePhase
                sPhases = "THREE"

        End Select
        Visa.SendString("INSTrument:PHASe " & sPhases)
    End Sub

    Public Overrides Function GetStatus() As Single Implements ISource_AC.GetStatus
        Return MyBase.GetStatus
    End Function

    Public Overrides Function GetVolt(Optional nPhase As Integer = 1) As Single Implements ISource_AC.GetVolt
        Return MyBase.GetVolt(nPhase)
    End Function

    Public Overrides Function GetCurrent(Optional nPhase As Integer = 1) As Single Implements ISource_AC.GetCurrent
        Return MyBase.GetCurrent(nPhase)
    End Function

    Public Overrides Function GetPower(Optional nPhase As Integer = 1) As Single Implements ISource_AC.GetPower
        Return MyBase.GetPower(nPhase)
    End Function

#End Region

#Region "Public Special Functions CHROMA 6530"



#End Region

End Class
