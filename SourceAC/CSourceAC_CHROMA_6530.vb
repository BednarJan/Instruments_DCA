'Class CSourceAC_CHROMA_6530
'14.04.2019, A. Zahler
'Compatible Instruments:
'- Chroma 6530 (single phase)
Imports Ivi.Visa

Public Class CSourceAC_CHROMA_6530
    Implements IDevice
    Implements ISource_AC

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMax As Single = 300 Implements ISource_AC.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 30 Implements ISource_AC.CurrentMax
    Public ReadOnly Property PowerMax As Single = 3000 Implements ISource_AC.PowerMax
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
        _Visa.SendString("RANG LOW")
        _Visa.SendString("SOUR:FREQ:IMM 50")
        _Visa.SendString("SOUR:FUNC:SHAP A")
        _Visa.SendString("TPH 0")
        _Visa.SendString("TPH:SYNC PHAS")
        _Visa.SendString("OREL ON")
        _Visa.SendString("SOUR:VOLT 0")
        _Visa.SendString("OUTP:STAT ON")
        _Visa.SendString("SOUR:FUNC:TRAN OFF")
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

        _Visa.SendString("SOUR:FREQ " & Freq)
        _Visa.SendString("SOUR:CURR " & CurrentLim & ";VOLT " & Voltage)
    End Sub

    Public Sub SetVoltage(Voltage As Single, Optional Phase As EPhaseNumber = 0) Implements ISource_AC.SetVoltage
        If Voltage > _VoltageMax Then Voltage = _VoltageMax

        Voltage = Math.Round(Voltage, 1)

        _Visa.SendString("SOUR:VOLT " & Voltage)
    End Sub

    Public Sub SetCurrentLim(CurrentLim As Single) Implements ISource_AC.SetCurrentLim
        If CurrentLim > _CurrentMax Then CurrentLim = _CurrentMax

        CurrentLim = Math.Round(CurrentLim, 2)

        _Visa.SendString("SOUR:CURR " & CurrentLim)
    End Sub

    Public Sub SetFreq(Freq As Single) Implements ISource_AC.SetFreq
        Freq = Math.Round(Freq, 2)

        _Visa.SendString("SOUR:FREQ " & Freq)
    End Sub

    Public Sub SetRange(Range As ERange) Implements ISource_AC.SetRange
        Select Case Range
            Case ERange.LOW
                _Visa.SendString("RANG LOW")
            Case ERange.HIGH
                _Visa.SendString("RANG HIGH")
            Case Else
                _Visa.SendString("RANG AUTO")
        End Select
    End Sub

    Public Sub SetShape(Shape As EShape) Implements ISource_AC.SetShape
        'Call SetOFF()
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
                sCmd &= " SIN"
                _Visa.SendString(sCmd)
            Case EWaveForm.SQUAre
                sCmd = "FUNC:SHAP:"
                sCmd &= [Enum].GetName(Shape.GetType, Shape)
                sCmd &= " SQU"
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
        VoltageAC = Math.Round(VoltageAC, 1)
        VoltageDC = Math.Round(VoltageDC, 1) 'Not supported by Chroma 6560
        Frequency = Math.Round(Frequency, 1)
        StartAngle = Math.Round(StartAngle, 1)
        VoltageStepAC = Math.Round(VoltageStepAC, 1)
        VoltageStepDC = Math.Round(VoltageStepDC, 1) 'Not supported by Chroma 6560
        FrequencyStep = Math.Round(FrequencyStep, 1)
        StepDwell = Math.Round(StepDwell, 1)

        _Visa.SendString("SOUR:FUNC:TRAN ON")
        _Visa.SendString("VOLT:MODE FIXED")
        _Visa.SendString("SOURCE:STEP:VOLTAGE " & VoltageAC)
        _Visa.SendString("SOURCE:STEP:FREQUENCY " & Frequency)
        _Visa.SendString("FUNC:SHAP " & [Enum].GetName(Shape.GetType, Shape))
        _Visa.SendString("SOURCE:STEP:SPHASE " & StartAngle)
        _Visa.SendString("SOURCE:STEP:DVOLTAGE " & VoltageStepAC)
        _Visa.SendString("SOURCE:STEP:DFREQUENCY " & FrequencyStep)
        _Visa.SendString("SOURCE:STEP:DWELL " & StepDwell)
        _Visa.SendString("SOURCE:STEP:COUNT " & Count)
        _Visa.SendString("VOLT:MODE STEP")
        _Visa.SendString("INIT")
        _Visa.SendString("TRIGger")
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
        Dim DwellTime As String = String.Empty
        Dim WaveShape As String = String.Empty
        Dim VoltStartAC As String = String.Empty, VoltEndAC As String = String.Empty
        Dim VoltStartDC As String = String.Empty, VoltEndDC As String = String.Empty
        Dim FreqStart As String = String.Empty, FreqEnd As String = String.Empty
        Dim Angle As String = String.Empty
        Dim i As Integer

        'Check if every array has the same size to be sure the Sequences are proper set
        'UBound(Shape)) And (UBound(Shape)=
        If (UBound(Dwell) = UBound(VoltageStartAC)) And
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
            'For i = 0 To UBound(Shape)
            '    WaveShape = WaveShape & Shape(i) & IIf(i = UBound(Dwell), "", ",")
            'Next i
            WaveShape = Shape.ToString

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
            'For i = 0 To UBound(StartAngle)
            Angle = Math.Round(StartAngle(1), 1)
            'Next i

            _Visa.SendString("SOUR:FUNC:TRAN ON")
            _Visa.SendString("VOLT:MODE FIXED")
            _Visa.SendString("SOURCE:LIST:COUNT " & Count)
            _Visa.SendString("SOURCE:LIST:DWELL " & DwellTime)
            '_Visa.SendString("SOURCE:LIST:SHAPE A;A;A;A;A") ' & WaveShape)
            _Visa.SendString("SOURCE:LIST:BASE " & [Enum].GetName(Base.GetType, Base))
            _Visa.SendString("SOURCE:LIST:VOLTAGE:START " & VoltStartAC)
            _Visa.SendString("SOURCE:LIST:VOLTAGE:END " & VoltEndAC)
            _Visa.SendString("SOURCE:LIST:FREQUENCY " & FreqStart)
            _Visa.SendString("SOURCE:LIST:SPHase " & Angle)
            _Visa.SendString("VOLT:MODE LIST")
            _Visa.SendString("INIT")
            _Visa.SendString("TRIGger")
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

        If VoltageStartAC < VoltageEndAC Then StepDVolt = 0.1 Else StepDVolt = -0.1
        StepCount = Math.Round(Math.Abs(VoltageStartAC - VoltageEndAC) / Math.Abs(StepDVolt))
        StepDwell = Math.Round(Duration / StepCount, 3)

        'If VoltageStartAC > 300 Then VoltageStartAC = 300
        'If VoltageEndAC > 300 Then VoltageEndAC = 300

        VoltageStartAC = Math.Round(VoltageStartAC, 1)
        VoltageEndAC = Math.Round(VoltageEndAC, 1)
        Frequency = Math.Round(Frequency, 1)

        _Visa.SendString("SOUR:FUNC:TRAN ON")
        _Visa.SendString("SOUR:VOLT:MODE STEP")
        _Visa.SendString("STEP:SYNC IMMediate")
        _Visa.SendString("STEP:COUNt " & StepCount)
        _Visa.SendString("STEP:DWELl " & StepDwell)
        _Visa.SendString("STEP:FREQuency " & Frequency)
        _Visa.SendString("STEP:DFREQuency 0")
        _Visa.SendString("STEP:VOLTage " & VoltageStartAC)
        _Visa.SendString("STEP:DVOLTage " & StepDVolt)
        _Visa.SendString("INIT")
        _Visa.SendString("TRIGger")
    End Sub

    Public Sub ClearProt() Implements ISource_AC.ClearProt
        _Visa.SendString("OUTP:PROT:CLEar")
    End Sub

    Public Function GetStatus() As Single Implements ISource_AC.GetStatus
        _Visa.SendString("SYST:ERR?")
        Return _Visa.ReceiveString()
    End Function
#End Region

#Region "Public Special Functions CHROMA 6530"
    Public Function GetVolt() As Single
        _Visa.SendString("MEAS:VOLT:AC?")
        Return _Visa.ReceiveValue()
    End Function

    'Public Function GetInrushCurrent() As Single 'Not supported by Chroma 6530?
    '    _Visa.SendString("MEAS:CURR:INR?")
    '    Return _Visa.ReceiveValue()
    'End Function
#End Region

End Class
