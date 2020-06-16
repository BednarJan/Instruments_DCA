'Class BCSourceAC common for all AC ources,  inherited fro the BCDevice
'16.06.2020, JaBe
'Base clase for the AC sources
Imports Ivi.Visa



Public Class BCSourceAC
    Inherits BCDevice
    Implements ISource_AC


#Region "Shorthand Properties"

    Public Property VoltageMaxLo As Single
    Public Property VoltageMax As Single Implements ISource_AC.VoltageMax
    Public Property VoltageMaxHi As Single

    Public Property CurrentMaxLo As Single
    Public Property CurrentMax As Single Implements ISource_AC.CurrentMax
    Public Property CurrentMaxHi As Single

    Public Property PowerMaxLo As Single
    Public Property PowerMax As Single Implements ISource_AC.PowerMax
    Public Property PowerMaxHi As Single

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Function IDN() As String Implements IDevice.IDN
        Return MyBase.IDN()
    End Function

    Public Overrides Sub RST() Implements IDevice.RST
        MyBase.RST()
    End Sub

    Public Overrides Sub CLS() Implements IDevice.CLS
        MyBase.CLS()
    End Sub

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        Visa.SendString("*RST;*CLS")
        cHelper.Delay(2)

        Visa.SendString("OUTP:PROT:CLE")
        Visa.SendString("RANG LOW")

        VoltageMax = VoltageMaxLo
        CurrentMax = CurrentMaxLo
        PowerMax = PowerMaxLo

        Visa.SendString("SOUR:FREQ:IMM 50")
        Visa.SendString("SOUR:VOLT:AC 0")
        Visa.SendString("SOUR:VOLT:DC 0")

        Visa.SendString("OUTP:TTLTrg:MODE FSTR")
        Visa.SendString("CURR:PROT:DEL 0.3") 'Set Current Limit Trip Delay
        Visa.SendString("SOURce:PONSetup:OLOad CVOL")

        Visa.SendString(":OUTPUT:STATE OFF")

    End Sub
#End Region

#Region "Interface Methodes ISource_AC"

    Public Overridable Sub SetOutputON() Implements ISource_AC.SetOutputON
        Visa.SendString(":OUTPUT:STATE ON")
    End Sub

    Public Overridable Sub SetOutputOFF() Implements ISource_AC.SetOutputOFF
        Visa.SendString(":OUTPUT:STATE OFF")
    End Sub

    Public Overridable Sub SetFreq(Freq As Single) Implements ISource_AC.SetFreq
        Freq = Math.Round(Freq, 2)
        Visa.SendString("SOUR:FREQ: IMM " & Freq.ToString)

    End Sub

    Public Overridable Sub SetVoltage(ByVal Voltage As Single, CurrentLim As Single, Optional Phase As ISource_AC.EPhaseNumber = ISource_AC.EPhaseNumber.All, Optional SetOutON As Boolean = True) Implements ISource_AC.SetVoltage

        If Voltage > _VoltageMax Then Voltage = _VoltageMax
        If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        Voltage = Math.Round(Voltage, 1)

        SetCurrentLim(CurrentLim)

        If Voltage = 0 Then
            Visa.SendString("SOUR:VOLTAGE " & Voltage.ToString)
            Visa.SendString(":OUTPUT:STATE OFF")
        Else
            Visa.SendString("SOURCE:VOLTAGE " & Voltage.ToString)
            If SetOutON Then Visa.SendString(":OUTPUT:STATE ON")
        End If

    End Sub

    Public Overridable Sub SetVoltage(Voltage As Single, Optional SetOutON As Boolean = True) Implements ISource_AC.SetVoltage

        SetVoltage(Voltage, _CurrentMax, SetOutON)

    End Sub

    Public Overridable Sub SetVoltPuls(ByVal vStart As Single, ByVal vPulse As Single, ByVal Width As Single, Optional ByVal phase As Single = 90) Implements ISource_AC.SetVoltPuls

        'PULSE: It makes the output change to its set value for a specific period of time as specified by parameters in the PULSE menu
        vStart = Math.Round(vStart, 1)
        vPulse = Math.Round(vPulse, 1)
        Width = Math.Round(Width, 1)
        phase = Math.Round(phase, 1)


        Visa.SendString("OUTP:TTLT:MODE FSTR")
        Visa.SendString("OUTP:TTLT ON")
        Visa.SendString("INIT TRAN")
        Visa.SendString("VOLT:MODE LIST")
        Visa.SendString("LIST:VOLTAGE " & CStr(vPulse) & "," & CStr(vStart))
        Visa.SendString("LIST:DWEL " & CStr(Width) & ",1")
        Visa.SendString("LIST:COUNT 1")
        Visa.SendString("TRIG:SYNC:SOUR PHAS  " & CStr(phase))
        Visa.SendString("INIT:IMM ")

        '   next 3 commands will set the output trigger mode to its
        '   initial state = trigger on every voltage change
        '   dont delete them

        Visa.SendString("TRIG:SYNC:SOUR IMM")
        Visa.SendString("VOLT:MODE FIX")
        Visa.SendString("OUTP:TTLT:MODE FSTR")

    End Sub


    Public Overridable Sub SetCurrentLim(CurrentLim As Single) Implements ISource_AC.SetCurrentLim

        If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        CurrentLim = Math.Round(CurrentLim, 2)

        Visa.SendString("SOUR:CURRENT " & CurrentLim.ToString)

    End Sub

    Public Overridable Sub SetRange(Range As ISource_AC.ERange) Implements ISource_AC.SetRange
        Select Case Range
            Case ISource_AC.ERange.LOW

                VoltageMax = VoltageMaxLo
                CurrentMax = CurrentMaxLo
                PowerMax = PowerMaxLo

            Case ISource_AC.ERange.HIGH

                VoltageMax = VoltageMaxHi
                CurrentMax = CurrentMaxHi
                PowerMax = PowerMaxHi

        End Select
        Visa.SendString("SOURce:VOLTage:RANGE " & VoltageMax)

    End Sub

    Public Overridable Sub SetPhaseMode(PhaseMode As ISource_AC.EPhaseMode) Implements ISource_AC.SetPhaseMode

        Dim nPhaseCount As Integer
        Select Case PhaseMode
            Case ISource_AC.EPhaseMode.SinglePhase
                nPhaseCount = 1
            Case ISource_AC.EPhaseMode.TwoPhase
                nPhaseCount = 2
            Case ISource_AC.EPhaseMode.ThreePhase
                nPhaseCount = 3

        End Select
        Visa.SendString("SYST:CONF:NOUT " & nPhaseCount)
    End Sub


    Public Overridable Function GetStatus() As Single Implements ISource_AC.GetStatus
        Visa.SendString("SYST:ERR?")
        Return Visa.ReceiveString()
    End Function


#End Region

End Class
