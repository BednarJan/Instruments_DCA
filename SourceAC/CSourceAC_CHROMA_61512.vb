'Class CSourceAC_CHROMA_61512
'14.04.2019, A. Zahler
'Compatible Instruments:
'- Chroma 61512 (three phase)
Imports Ivi.Visa

Public Class CSourceAC_CHROMA_61512
    Implements IDevice
    Implements ISource_AC

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMax As Single = 300 Implements ISource_AC.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 144 Implements ISource_AC.CurrentMax
    Public ReadOnly Property PowerMax As Single = 18000 Implements ISource_AC.PowerMax
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Visa = New CVisaDeviceNI(Session, ErrorLogger)
        _ErrorLogger = ErrorLogger
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"
    Public Function IDN() As String Implements IDevice.IDN
        Dim ErrorMessages(1) As String

        _Visa.SendString("*IDN?", ErrorMessages(0))
        Return _Visa.ReceiveString(ErrorMessages(1))
    End Function

    Public Sub RST() Implements IDevice.RST
        Dim ErrorMessage As String = ""

        _Visa.SendString("*RST", ErrorMessage)
    End Sub

    Public Sub CLS() Implements IDevice.CLS
        Dim ErrorMessage As String = ""

        _Visa.SendString("*CLS", ErrorMessage)
    End Sub

    Public Sub Initialize() Implements IDevice.Initialize
        _Visa.SendString("*RST;*CLS:")
        _Visa.SendString("OUTP:PROT:CLE")
        _Visa.SendString("SYSTEM:DATE " & DateTime.Now.Year.ToString & "," & DateTime.Now.Month.ToString & "," & DateTime.Now.Day.ToString)
        _Visa.SendString("SYSTEM:TIME " & DateTime.Now.Hour.ToString & "," & DateTime.Now.Minute.ToString & "," & DateTime.Now.Second.ToString)
        _Visa.SendString("OUTPUT:STATE OFF")
        _Visa.SendString("INSTRUMENT:EDIT ALL")
        _Visa.SendString("INSTRUMENT:COUPLE ALL")
        _Visa.SendString("OUTPUT:PROTECTION:CLEAR")
        _Visa.SendString("OUTPUT:MODE FIXED")
        _Visa.SendString("OUTPUT:SLEW:VOLTAGE:AC 000.000")
        _Visa.SendString("OUTPUT:SLEW:VOLTAGE:DC 000.000")
        _Visa.SendString("OUTPUT:SLEW:FREQUENCY 000.000")
        _Visa.SendString("SOURCE:CONFIGURE:INHIBIT DISABLE") 'Set to ENABLE to activate remote inhibit
        _Visa.SendString("SOURCE:CURRENT:LIMIT 48") '3-phase: Max. 48A, 1-phase: Max. 144A
        _Visa.SendString("SOURCE:CURRENT:DELAY 5")
        _Visa.SendString("SOURCE:FREQUENCY:IMMEDIATE 50")
        '_Visa.SendString("SOURCE:FREQUENCY:LIMIT 70")
        _Visa.SendString("SOURCE:POWER:PROTECTION " & PowerMax)
        _Visa.SendString("SOURCE:VOLTAGE:LEVEL:IMMEDIATE:AMPLITUDE:AC 0")
        _Visa.SendString("SOURCE:VOLTAGE:LEVEL:IMMEDIATE:AMPLITUDE:DC 0")
        _Visa.SendString("SOURCE:VOLTAGE:LIMIT:AC 300")
        _Visa.SendString("SOURCE:VOLTAGE:LIMIT:DC:PLUS 0")
        _Visa.SendString("SOURCE:VOLTAGE:LIMIT:DC:MINUS 0")
        _Visa.SendString("SOURCE:VOLTAGE:RANGE HIGH")
        _Visa.SendString("SOURCE:FUNCTION:SHAPE A")
        _Visa.SendString("SOURCE:PHASE:SEQUENCE POS")
        _Visa.SendString("SOURCE:PHASE:ON 0")
        _Visa.SendString("SOURCE:PHASE:OFF 0")
        _Visa.SendString("SOURCE:PHASE:P12 120")
        _Visa.SendString("SOURCE:PHASE:P13 240")
        _Visa.SendString("OUTPUT:RELAY ON")
        _Visa.SendString("OUTPUT:STATE ON")
    End Sub
#End Region

#Region "Interface Methodes ISource_AC"
    Public Sub SetOutputON() Implements ISource_AC.SetOutputON
        _Visa.SendString("OUTP:STAT ON")
    End Sub

    Public Sub SetOutputOFF() Implements ISource_AC.SetOutputOFF
        _Visa.SendString("OUTP:STAT OFF")
    End Sub

    Public Sub SetSource(Voltage As Single, CurrentLim As Single, Freq As Single, Optional Phase As EPhaseNumber = 0) Implements ISource_AC.SetSource
        If Voltage > _VoltageMax Then Voltage = _VoltageMax
        If CurrentLim > _CurrentMax Then CurrentLim = _CurrentMax

        Voltage = Math.Round(Voltage, 1)
        CurrentLim = Math.Round(CurrentLim, 2)
        Freq = Math.Round(Freq, 2)

        _Visa.SendString("SOURCE:CURRENT:LIMIT " & CurrentLim)
        _Visa.SendString("SOURCE:FREQUENCY:IMMEDIATE " & Freq)

        Select Case Phase
            Case EPhaseNumber.All
                _Visa.SendString("INSTRUMENT:COUPLE ALL")
            Case Else
                _Visa.SendString("INSTRUMENT:COUPLE NONE")
                _Visa.SendString("INSTRUMENT:NSELECT " & Phase)
        End Select

        _Visa.SendString("OUTPUT:MODE FIXED")
        _Visa.SendString("SOURCE:VOLTAGE:LEVEL:IMMEDIATE:AC " & Voltage)
    End Sub

    Public Sub SetVoltage(Voltage As Single, Optional Phase As EPhaseNumber = 0) Implements ISource_AC.SetVoltage
        If Voltage > _VoltageMax Then Voltage = _VoltageMax

        Voltage = Math.Round(Voltage, 1)

        Select Case Phase
            Case EPhaseNumber.All
                _Visa.SendString("INSTRUMENT:COUPLE ALL")
            Case Else
                _Visa.SendString("INSTRUMENT:COUPLE NONE")
                _Visa.SendString("INSTRUMENT:NSELECT " & Phase)
        End Select
        _Visa.SendString("OUTPUT:MODE FIXED")
        _Visa.SendString("SOURCE:VOLTAGE:LEVEL:IMMEDIATE:AC " & Voltage)

    End Sub

    Public Sub SetCurrentLim(CurrentLim As Single) Implements ISource_AC.SetCurrentLim
        If CurrentLim > _CurrentMax Then CurrentLim = _CurrentMax

        CurrentLim = Math.Round(CurrentLim, 2)

        _Visa.SendString("SOUR:CURR:LIM " & CurrentLim)
    End Sub

    Public Sub SetFreq(Freq As Single) Implements ISource_AC.SetFreq
        Freq = Math.Round(Freq, 2)

        _Visa.SendString("SOUR:FREQ:IMM " & Freq)
    End Sub

    Public Sub SetRange(Range As ERange) Implements ISource_AC.SetRange
        Select Case Range
            Case ERange.LOW
                _Visa.SendString("RANG LOW")
            Case ERange.HIGH
                _Visa.SendString("RANG HIGH")
            Case Else
                _Visa.SendString("RANG AUTO") 'Check if Auto is possible with Chroma 61512
        End Select
    End Sub

    Public Sub SetShape(Shape As EShape) Implements ISource_AC.SetShape
        Select Case Shape
            Case EShape.A
                _Visa.SendString("FUNC:SHAP A")
            Case EShape.B
                _Visa.SendString("FUNC:SHAP B")
        End Select
    End Sub

    Public Sub SetWaveForm(Shape As EShape, Waveform As EWaveForm) Implements ISource_AC.SetWaveForm
        Dim sCmd As String = String.Empty
        Select Case Waveform
            Case EWaveForm.SINe
                sCmd = "FUNC:SHAP:"
                sCmd &= [Enum].GetName(Shape.GetType, Shape)
                sCmd &= " SINE"
                _Visa.SendString(sCmd)
            Case EWaveForm.SQUAre
                sCmd = "FUNC:SHAP:"
                sCmd &= [Enum].GetName(Shape.GetType, Shape)
                sCmd &= " SQUA"
                _Visa.SendString(sCmd)
            Case EWaveForm.CSIN
                sCmd = "FUNC:SHAP:"
                sCmd &= [Enum].GetName(Shape.GetType, Shape)
                sCmd &= " CSIN"
                _Visa.SendString(sCmd)
            Case Else
                sCmd = "FUNC:SHAP:"
                sCmd &= [Enum].GetName(Shape.GetType, Shape)
                sCmd &= " "
                sCmd &= [Enum].GetName(Waveform.GetType, Waveform)
                _Visa.SendString(sCmd)
        End Select
    End Sub

    Public Sub SetStep(VoltageAC As Single,
                       VoltageDC As Single,
                       Frequency As Single,
                       Shape As EShape,
                       StartAngle As Single,
                       VoltageStepAC As Single,
                       VoltageStepDC As Single,
                       FrequencyStep As Single,
                       StepDwell As Single,
                       Count As Integer) Implements ISource_AC.SetStep

        'STEP: It makes the output change to its set values step by step
        'VoltageAC      = 0...150VAC (low Range) | 0...300VAC (high Range)
        'VoltageDC      = -212.1...212.1VDC (low Range) | -424.2...424.2VDC (high Range)
        'Frequency      = 15...1500Hz
        'Shape          = "A" | "B"
        'StartPhase     = 0...359.9°
        'VoltageStepAC  = -150...150VAC (low Range) | -300...3000 (high Range)
        'VoltageStepDC  = -212.1...212.1VDC (low Range) | -424.2...424.2VDC (high Range)
        'FrequencyStep  = -1500...1500Hz
        'StepDwell      = 0...99999999.9ms
        'Count          = 0...65535

        If VoltageAC > _VoltageMax Then VoltageAC = _VoltageMax

        VoltageAC = Math.Round(VoltageAC, 1)
        VoltageDC = Math.Round(VoltageDC, 1)
        Frequency = Math.Round(Frequency, 1)
        StartAngle = Math.Round(StartAngle, 1)
        VoltageStepAC = Math.Round(VoltageStepAC, 1)
        VoltageStepDC = Math.Round(VoltageStepDC, 1)
        FrequencyStep = Math.Round(FrequencyStep, 1)
        StepDwell = Math.Round(StepDwell, 1)

        _Visa.SendString("TRIG OFF")
        _Visa.SendString("OUTPUT:MODE FIXED")
        _Visa.SendString("SOURCE:STEP:VOLTAGE:AC " & VoltageAC)
        _Visa.SendString("SOURCE:STEP:VOLTAGE:DC " & VoltageDC)
        _Visa.SendString("SOURCE:STEP:FREQUENCY " & Frequency)
        _Visa.SendString("SOURCE:STEP:SHAPE " & Shape)
        _Visa.SendString("SOURCE:STEP:SPHASE " & StartAngle)
        _Visa.SendString("SOURCE:STEP:DVOLTAGE:AC " & VoltageStepAC)
        _Visa.SendString("SOURCE:STEP:DVOLTAGE:DC " & VoltageStepDC)
        _Visa.SendString("SOURCE:STEP:DFREQUENCY " & FrequencyStep)
        _Visa.SendString("SOURCE:STEP:DWELL " & StepDwell)
        _Visa.SendString("SOURCE:STEP:COUNT " & Count)
        _Visa.SendString("OUTPUT:MODE STEP")
        _Visa.SendString("TRIG ON")
    End Sub

    Public Sub SetList(Count As UInteger,
               Dwell() As Single,
               Shape() As EShape,
               Base As EBase,
               VoltageStartAC() As Single,
               VoltageEndAC() As Single,
               VoltageStartDC() As Single,
               VoltageEndDC() As Single,
               FrequencyStart() As Single,
               FrequencyEnd() As Single,
               StartAngle() As Single) Implements ISource_AC.SetList

        'LIST: It makes the output change through a number of values specified by parameters in the LIST menu
        'Count          = 0...65535
        'Dwell          = Array[S1,...,S100] 0...99999999.9ms
        'Shape          = Array[S1,...,S100] "A" | "B"
        'Base           = TIME | CYCLE
        'VoltageStartAC = Array[S1,...,S100] 0...150VAC (low Range) | 0...300VAC (high Range)
        'VoltageEndAC   = Array[S1,...,S100] 0...150VAC (low Range) | 0...300VAC (high Range)
        'VoltageStartDC = Array[S1,...,S100] -212.1...212.1VDC (low Range) | -424.2...424.2VDC (high Range)
        'VoltageEndDC   = Array[S1,...,S100] -212.1...212.1VDC (low Range) | -424.2...424.2VDC (high Range)
        'FrequencyStart = Array[S1,...,S100] 15...1500Hz
        'FrequencyEnd   = Array[S1,...,S100] 15...1500Hz
        'StartAngle     = Array[S1,...,S100] 0.0...359.9°

        Dim DwellTime As String = String.Empty
        Dim WaveShape As String = String.Empty
        Dim VoltStartAC As String = String.Empty, VoltEndAC As String = String.Empty
        Dim VoltStartDC As String = String.Empty, VoltEndDC As String = String.Empty
        Dim FreqStart As String = String.Empty, FreqEnd As String = String.Empty
        Dim Angle As String = String.Empty
        Dim i As Integer

        'Check if every array has the same size to be sure the Sequences are proper set
        If (UBound(Dwell) = UBound(Shape)) And
           (UBound(Shape) = UBound(VoltageStartAC)) And
           (UBound(VoltageStartAC) = UBound(VoltageEndAC)) And
           (UBound(VoltageEndAC) = UBound(VoltageStartDC)) And
           (UBound(VoltageStartDC) = UBound(VoltageEndDC)) And
           (UBound(VoltageEndDC) = UBound(FrequencyStart)) And
           (UBound(FrequencyStart) = UBound(FrequencyEnd)) And
           (UBound(FrequencyEnd) = UBound(StartAngle)) Then

            'Make Strings from Array
            For i = 0 To UBound(Dwell)
                DwellTime = DwellTime & Math.Round(Dwell(i), 1) & IIf(i = UBound(Dwell), "", ",")
            Next i
            For i = 0 To UBound(Shape)
                WaveShape = WaveShape & Shape(i) & IIf(i = UBound(Dwell), "", ",")
            Next i
            For i = 0 To UBound(VoltageStartAC)
                VoltStartAC = VoltStartAC & Math.Round(VoltageStartAC(i), 1) & IIf(i = UBound(Dwell), "", ",")
            Next i
            For i = 0 To UBound(VoltageEndAC)
                VoltEndAC = VoltEndAC & Math.Round(VoltageEndAC(i), 1) & IIf(i = UBound(Dwell), "", ",")
            Next i
            For i = 0 To UBound(VoltageStartDC)
                VoltStartDC = VoltStartDC & Math.Round(VoltageStartDC(i), 1) & IIf(i = UBound(Dwell), "", ",")
            Next i
            For i = 0 To UBound(VoltageEndDC)
                VoltEndDC = VoltEndDC & Math.Round(VoltageEndDC(i), 1) & IIf(i = UBound(Dwell), "", ",")
            Next i
            For i = 0 To UBound(FrequencyStart)
                FreqStart = FreqStart & Math.Round(FrequencyStart(i), 2) & IIf(i = UBound(Dwell), "", ",")
            Next i
            For i = 0 To UBound(FrequencyEnd)
                FreqEnd = FreqEnd & Math.Round(FrequencyEnd(i), 2) & IIf(i = UBound(Dwell), "", ",")
            Next i
            For i = 0 To UBound(StartAngle)
                Angle = Angle & Math.Round(StartAngle(i), 1) & IIf(i = UBound(Dwell), "", ",")
            Next i

            _Visa.SendString("TRIG OFF")
            _Visa.SendString("OUTPUT:MODE FIXED")
            _Visa.SendString("SOURCE:LIST:COUPLING ALL")
            _Visa.SendString("SOURCE:LIST:COUNT " & Count)
            _Visa.SendString("SOURCE:LIST:DWELL " & DwellTime)
            _Visa.SendString("SOURCE:LIST:SHAPE " & WaveShape)
            _Visa.SendString("SOURCE:LIST:BASE " & Base)
            _Visa.SendString("SOURCE:LIST:VOLTAGE:AC:START " & VoltStartAC)
            _Visa.SendString("SOURCE:LIST:VOLTAGE:AC:END " & VoltEndAC)
            _Visa.SendString("SOURCE:LIST:VOLTAGE:DC:START " & VoltStartDC)
            _Visa.SendString("SOURCE:LIST:VOLTAGE:DC:END " & VoltEndDC)
            _Visa.SendString("SOURCE:LIST:FREQUENCY:START " & FreqStart)
            _Visa.SendString("SOURCE:LIST:FREQUENCY:END " & FreqEnd)
            _Visa.SendString("SOURCE:LIST:DEGREE " & Angle)
            _Visa.SendString("OUTPUT:MODE LIST")
            _Visa.SendString("TRIG ON")
        Else
            Call SetOutputOFF()
            _ErrorLogger.LogInfo("Chroma 61512 SetList: Array length not consistent, Source set OFF")
        End If
    End Sub

    Public Sub SetPulse(VoltageAC As Single,
                            VoltageDC As Single,
                            Frequency As Single,
                            Shape As EShape,
                            StartAngle As Single,
                            Count As UInteger,
                            DutyCycle As Single,
                            Duration As Single) Implements ISource_AC.SetPulse

        'PULSE: It makes the output change to its set value for a specific period of time as specified by parameters in the PULSE menu
        'VoltageAC      = 0...150VAC (low Range) | 0...300VAC (high Range)
        'VoltageDC      = -212.1...212.1VDC (low Range) | -424.2...424.2VDC (high Range)
        'Frequency      = 15...1500Hz
        'Shape          = "A" | "B"
        'StartPhase     = 0...359.9°
        'Count          = 0...65535
        'DutyCycle      = 0...100%
        'Duration       = 0...99999999.9ms

        If VoltageAC > _VoltageMax Then VoltageAC = _VoltageMax

        VoltageAC = Math.Round(VoltageAC, 1)
        VoltageDC = Math.Round(VoltageDC, 1)
        StartAngle = Math.Round(StartAngle, 1)
        Frequency = Math.Round(Frequency, 1)
        DutyCycle = Math.Round(DutyCycle, 0)
        Duration = Math.Round(Duration, 1)

        If DutyCycle = 50 Then Duration = Duration * 2

        _Visa.SendString("TRIG OFF")
        _Visa.SendString("OUTPUT:MODE FIXED")
        _Visa.SendString("SOURCE:PULSE:VOLTAGE:AC " & VoltageAC)
        _Visa.SendString("SOURCE:PULSE:VOLTAGE:DC " & VoltageDC)
        _Visa.SendString("SOURCE:PULSE:FREQUENCY " & Frequency)
        _Visa.SendString("SOURCE:PULSE:SHAPE " & Shape)
        _Visa.SendString("SOURCE:PULSE:SPHASE " & StartAngle)
        _Visa.SendString("SOURCE:PULSE:COUNT " & Count)
        _Visa.SendString("SOURCE:PULSE:DCYCLE " & DutyCycle)
        _Visa.SendString("SOURCE:PULSE:PERIOD " & Duration)
        _Visa.SendString("OUTPUT:MODE PULSE")
        _Visa.SendString("TRIG ON")
    End Sub

    Public Sub SetBrownOut(VoltageStartAC As Single,
                           VoltageEndAC As Single,
                           Duration As Single,
                           Frequency As Single) Implements ISource_AC.SetBrownOut

        Dim StepDVolt As Single, StepCount As Integer, StepDwell As Single

        If VoltageStartAC > 300 Then
            Call SetOutputOFF()
            _ErrorLogger.LogInfo("Chroma 61512 SetBrownOut: VoltageStartAC over VoltageMax, Source set OFF")
        ElseIf VoltageEndAC > 300 Then
            Call SetOutputOFF()
            _ErrorLogger.LogInfo("Chroma 61512 SetBrownOut: VoltageEndAC over VoltageMax, Source set OFF")
        Else
            'If VoltageEndAC > 300 Then VoltageEndAC = 300

            If VoltageStartAC < VoltageEndAC Then StepDVolt = 0.1 Else StepDVolt = -0.1
            StepCount = Math.Round(Math.Abs(VoltageStartAC - VoltageEndAC) / Math.Abs(StepDVolt))
            StepDwell = Math.Round(Duration / StepCount, 3)
            VoltageStartAC = Math.Round(VoltageStartAC, 1)
            VoltageEndAC = Math.Round(VoltageEndAC, 1)
            Frequency = Math.Round(Frequency, 1)

            _Visa.SendString("TRIG OFF")
            _Visa.SendString("OUTPUT:MODE FIXED")
            _Visa.SendString("STEP:SYNC IMMediate")
            _Visa.SendString("STEP:COUNt " & StepCount)
            _Visa.SendString("STEP:DWELl " & StepDwell)
            _Visa.SendString("STEP:FREQuency " & Frequency)
            _Visa.SendString("STEP:DFREQuency 0")
            _Visa.SendString("STEP:VOLTage:AC " & VoltageStartAC)
            _Visa.SendString("STEP:DVOLTage:AC " & StepDVolt)
            _Visa.SendString("OUTP:MODE STEP")
            _Visa.SendString("TRIG ON")
        End If
    End Sub

    Public Sub ClearProt() Implements ISource_AC.ClearProt
        _Visa.SendString("OUTP:PROT:CLEar")
    End Sub

    Public Function GetStatus() As Single Implements ISource_AC.GetStatus
        _Visa.SendString("STAT:QUES:COND?")
        Return _Visa.ReceiveString()
    End Function
#End Region

#Region "Public Special Functions CHROMA 61512"
    Private Sub SetPhaseMode(ByVal PhaseMode As EPhaseMode)
        '"THREE", "SINGLE"
        Select Case PhaseMode
            Case EPhaseMode.SinglePhase
                Visa.SendString("INSTRUMENT:PHASE SINGLE")
            Case EPhaseMode.ThreePhase
                Visa.SendString("INSTRUMENT:PHASE THREE")
        End Select
    End Sub

    Private Function getPhaseMode() As EPhaseMode
        Dim PhaseMode As String
        Dim bRet As EPhaseMode = EPhaseMode.SinglePhase

        _Visa.SendString("INSTRUMENT:PHASE?")
        PhaseMode = _Visa.ReceiveString()

        If PhaseMode = Left(PhaseMode, 6) = "SINGLE" Then bRet = EPhaseMode.SinglePhase
        If PhaseMode = Left(PhaseMode, 5) = "THREE" Then bRet = EPhaseMode.ThreePhase

        Return bRet

    End Function

    'Private Function checkPhaseMode(ByVal Mode As String)
    '    'THREE | SINGLE
    '    If Me.getPhaseMode <> Mode Then
    '        Call MsgBox("The Source has wrong Phase Setting! " & Chr(10) &
    '"Please check the connection behind the Source and press Enter if it's correct!" & Chr(10) & Chr(10) &
    '"Wrong connection can cause damage on the Source and other Equipment!", vbCritical, "Chroma 61512 wrong Phase Setting")
    '        Call MsgBox("Press 'OK' if it's save to change the Phase Mode Setting", vbExclamation, "Save to change Phase Mode")
    '        Me.setPhaseMode(Mode)
    '    End If
    'End Function

    Private Sub SetInterHarmonics(VoltageAC As Single,
                                  VoltageFrequency As Single,
                                  InterHarm_Fstart As Single,
                                  InterHarm_Fend As Single,
                                  InterHarm_Level As Single,
                                  InterHarm_Fdwell As Single)

        'INTERHARMONICS: Adds another frequency of variable voltage component to the fundamental voltage output
        'VoltageAC          = 0...150VAC (low Range) | 0...300VAC (high Range)
        'VoltageFrequency   = 15...1500Hz
        'InterHarm_Fstart   = 0.01...2400Hz
        'InterHarm_Fend     = 0.01...2400Hz
        'InterHarm_Level    = 0...30% @ 0.01Hz...500Hz      (The rms of scanning wave that is
        '                     0...20% @ 500.01...1000Hz     the percentage of fundamental voltage)
        '                     0...10% @ 1000.01...2400Hz
        'InterHarm_Fdwell   = 0.00...99999.99s (mandatory value. shouldn't be 0)

        If VoltageAC > _VoltageMax Then VoltageAC = _VoltageMax

        VoltageAC = Math.Round(VoltageAC, 1)
        VoltageFrequency = Math.Round(VoltageFrequency, 1)
        InterHarm_Fstart = Math.Round(InterHarm_Fstart, 2)
        InterHarm_Fend = Math.Round(InterHarm_Fend, 2)
        InterHarm_Level = Math.Round(InterHarm_Level, 0)
        InterHarm_Fdwell = Math.Round(InterHarm_Fdwell, 2)

        If InterHarm_Level > 30 Then InterHarm_Level = 30
        'If InterHarm_Level > 30 And (InterHarm_Fstart > 500 Or InterHarm_Fend > 500) Then InterHarm_Level = 30
        If InterHarm_Level > 20 And (InterHarm_Fstart > 1000 Or InterHarm_Fend > 1000) Then InterHarm_Level = 20
        If InterHarm_Level > 10 And (InterHarm_Fstart > 2400 Or InterHarm_Fend > 2400) Then InterHarm_Level = 10

        _Visa.SendString("TRIG OFF")
        _Visa.SendString("OUTPUT:MODE FIXED")
        _Visa.SendString("SOURCE:VOLTAGE:AC " & VoltageAC)
        _Visa.SendString("SOURCE:FREQUENCY " & VoltageFrequency)
        _Visa.SendString("SOURCE:INTERHARMONICS:FREQUENCY:START " & InterHarm_Fstart)
        _Visa.SendString("SOURCE:INTERHARMONICS:FREQUENCY:END " & InterHarm_Fend)
        _Visa.SendString("SOURCE:INTERHARMONICS:LEVEL " & InterHarm_Level)
        _Visa.SendString("SOURCE:INTERHARMONICS:DWELL " & InterHarm_Fdwell)
        _Visa.SendString("OUTPUT:MODE INTERHAR")
        _Visa.SendString("TRIG ON")
    End Sub

    Public Function GetVolt() As Single
        _Visa.SendString("MEAS:VOLT:AC?")
        Return _Visa.ReceiveValue()
    End Function

    Public Function GetInrushCurrent() As Single
        _Visa.SendString("MEAS:CURR:INR?")
        Return _Visa.ReceiveValue()
    End Function
#End Region

End Class
