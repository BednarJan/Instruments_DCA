'Class CSourceAC_CHROMA_61512
'14.04.2019, A. Zahler
'16.04.2020, JaBe  ..........  must be actualised
'Compatible Instruments:
'- Chroma 6530 (single phase)
Imports Ivi.Visa

Public Class CSourceAC_CHROMA_61512
    Inherits BCSourceAC
    Implements ISource_AC

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)

        CurrentMaxHi = 30
        CurrentMaxLo = 30

        VoltageMaxHi = 300
        VoltageMaxLo = 300

        PowerMaxHi = 3000
        PowerMaxLo = 3000

        Initialize()

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
        Visa.SendString("OUTP:STAT ON")
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
        MyBase.SetPhaseMode(PhaseMode)
    End Sub

    Public Overrides Function GetStatus() As Single Implements ISource_AC.GetStatus
        Return MyBase.GetStatus
    End Function
#End Region

#Region "Public Special Functions CHROMA 6530"
    Public Function GetVolt() As Single
        Visa.SendString("MEAS:VOLT:AC?")
        Return Visa.ReceiveValue()
    End Function

#End Region

End Class
