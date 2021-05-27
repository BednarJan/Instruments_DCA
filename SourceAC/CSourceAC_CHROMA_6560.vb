'Class CSourceAC_CHROMA_6560
'26.05.2021, JaBe  
'Compatible Instruments:
Imports Ivi.Visa

Public Class CSourceAC_CHROMA_6560
    Inherits BCSourceAC
    Implements ISource_AC

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        CurrentMaxHi = 23
        CurrentMaxLo = 23

        VoltageMaxHi = 300
        VoltageMaxLo = 150

        PowerMaxHi = 6000
        PowerMaxLo = 6000

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

        Throw New CInstrumentException(New NotImplementedException, Visa.Session.ResourceName)

    End Sub

#End Region

#Region "Public Special Functions CHROMA 6560"


#End Region

End Class
