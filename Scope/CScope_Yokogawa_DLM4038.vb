Imports Ivi.Visa

Public Class CScope_Yokogawa_DLM4038
    Implements IDevice

    Private _ErrorLogger As CErrorLogger

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property mChanCount As Integer = 8
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Visa = New CVisaDevice(Session, ErrorLogger)
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

    Public Sub SndString(cmdStr As String) Implements IDevice.SendString
        _Visa.SendString(cmdStr)
    End Sub

    Public Function RecieveString() As String Implements IDevice.ReceiveString
        Return _Visa.ReceiveString
    End Function

    Public Sub Initialize() Implements IDevice.Initialize
        Dim Channel As Integer = 0

        _Visa.SendString("*CLS;*RST")

        _Visa.SendString("COMMUNICATE:REMOTE ON")
        _Visa.SendString("*RST;*CLS;")
        _Visa.SendString("INITIALIZE:EXECUTE")
        _Visa.SendString("CHUTIL:ALL:DISPLAY OFF")
        _Visa.SendString(":DISPLAY:FGRID OFF")
        _Visa.SendString(":DISPLAY:INTENSITY:GRID 32")
        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString("SYSTEM:CLOCK:DATE " & DateTime.Now.ToString("yyyy/MM/dd").Replace(".", "/")) 'DateTime.Now.Year.ToString & "/" & DateTime.Now.Month.ToString & "/" & DateTime.Now.Day.ToString)
        _Visa.SendString("SYSTEM:CLOCK:TIME " & DateTime.Now.Hour.ToString & ":" & DateTime.Now.Minute.ToString & ":" & DateTime.Now.Second.ToString)
        _Visa.SendString(":CALIBRATE:MODE OFF")
    End Sub

#End Region

#Region "General Settings"
    Public Sub Calibrate()
        _Visa.SendString(":CALIBRATE:EXECUTE")
    End Sub

    Public Sub SetTimebase(ByVal Hor As Single)
        _Visa.SendString(":TIMEBASE:TDIV " & Hor)
    End Sub

    Public Function GetTimebase() As Single
        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString(":TIMEBASE:TDIV?")
        GetTimebase = _Visa.ReceiveValue()
    End Function

    Public Function GetSample_Rate() As Single
        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString(":TIMEBASE:SRATE?")
        GetSample_Rate = _Visa.ReceiveValue()
    End Function
#End Region

#Region "Channel Methodes"
    Public Sub ChannelOff(Optional ByVal Chan As UInteger = 0)

        If Chan = 0 Then
            _Visa.SendString(":CHUTIL:ALL:DISPLAY OFF")
        Else
            _Visa.SendString("CHANNEL" & Chan & ":DISPLAY OFF")
        End If

    End Sub

    Public Sub ChannelOn(Optional ByVal Chan As UInteger = 0)

        If Chan = 0 Then
            _Visa.SendString(":CHUTIL:ALL:DISPLAY ON")
        Else
            _Visa.SendString("CHANNEL" & Chan & ":DISPLAY ON")
        End If

    End Sub

    Public Sub ChannelConfig(ByVal Chan As Integer,
                      ByVal Lab As String,
                      ByVal Pos As Single,
                      ByVal Vert As Single,
                      ByVal Probe As Integer,
                      ByVal Unit As String,
                      ByVal Coupling As String,
                      ByVal Offs As Single,
                      ByVal Bandwidth As String)

        Dim strUnit As String
        'Channel:  1...8
        'Label:    whatever you want
        'Position: -4 ... 4 (single)
        'Vert:     try something, it is possible to use a single type value
        'Probe:    depends on the probe you are using
        'Unit:     "V" | "A"
        'Coupling: AC | DC | DC50 | GND
        'Offset:   Channel Offset for DC Measurements
        'Bandwith: "FULL" | "200 MHz" | "100 MHz" | "20 Mhz" | "10 MHz" | "5 MHz" | "2 MHz" | "1 MHz" | "500 kHz" | "250 kHz" | "125 kHz" | "62.5 kHz" | "32 kHz" | "16 kHz" | "8 kHz"

        If Unit = "A" Then
            strUnit = "C"
        Else
            strUnit = ""
        End If

        _Visa.SendString("CHANNEL" & Chan & ":DISPLAY ON")
        _Visa.SendString("CHANNEL" & Chan & ":LABEL:DISPLAY ON")
        _Visa.SendString("CHANNEL" & Chan & ":LABEL:DEFINE '" & Lab & "'")
        _Visa.SendString(":POSITION " & Pos)
        _Visa.SendString(":PROBE:MODE " & strUnit & Probe)
        _Visa.SendString(":VDIV " & Vert)
        _Visa.SendString(":COUPLING " & Coupling)
        _Visa.SendString(":OFFSET " & Offs)
        _Visa.SendString(":BWIDTH " & Bandwidth)
        _Visa.SendString(":INVERT OFF")
        _Visa.SendString(":LSCALE:MODE OFF")
        _Visa.SendString(":DESKEW 0")

    End Sub

    Public Sub ChannelLinearscale(ByVal Chan As Integer,
                           ByVal Gain As Single,
                           ByVal Offset As Single)

        _Visa.SendString("CHANNEL" & Chan & ":LSCALE:MODE ON")
        _Visa.SendString("CHANNEL" & Chan & ":LSCALE:AVALUE " & Gain)
        _Visa.SendString("CHANNEL" & Chan & ":LSCALE:BVALUE " & Offset)

    End Sub

    Public Function ChannelGetLable(ByVal Chan As Integer) As String

        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString("CHANNEL" & Chan & ":LABEL:DEFINE?")
        ChannelGetLable = _Visa.ReceiveString()

    End Function

#End Region

#Region "Math Methodes"
    Public Sub MathOff(Optional ByVal Chan As Integer = 0)

        If Chan = 0 Then
            _Visa.SendString("MATH1:DISPLAY OFF")
            _Visa.SendString("MATH2:DISPLAY OFF")
            _Visa.SendString("MATH3:DISPLAY OFF")
            _Visa.SendString("MATH4:DISPLAY OFF")
        Else
            _Visa.SendString("MATH" & Chan & ":DISPLAY OFF")
        End If

    End Sub

    Public Sub MathOn(Optional ByVal Chan As Integer = 0)

        If Chan = 0 Then
            _Visa.SendString("MATH1:DISPLAY ON")
            _Visa.SendString("MATH2:DISPLAY ON")
            _Visa.SendString("MATH3:DISPLAY ON")
            _Visa.SendString("MATH4:DISPLAY ON")
        Else
            _Visa.SendString("MATH" & Chan & ":DISPLAY ON")
        End If

    End Sub

    Public Sub MathConfig(ByVal Chan As Integer,
                   ByVal Operation As String,
                   ByVal Lab As String,
                   ByVal Pos As Single,
                   ByVal Vert As Single,
                   ByVal Unit As String)
        'Channel:  1...4
        'Operation: E.g. "C3*C6"
        'Label:    whatever you want
        'Position: -4 ... 4 (single)
        'Vert:     try something, it is possible to use a single type value
        'Unit:     "V" | "A"

        _Visa.SendString(":MATH" & Chan & ":DISPLAY ON")
        _Visa.SendString(":MATH" & Chan & ":LABEL:MODE ON")
        _Visa.SendString(":MATH" & Chan & ":LABEL:DEFINE '" & Lab & "'")
        _Visa.SendString(":MATH" & Chan & ":UNIT:MODE USERDEFINE")
        _Visa.SendString(":MATH" & Chan & ":UNIT:DEFINE '" & Unit & "'")
        _Visa.SendString(":MATH" & Chan & ":OPERATION USERDEFINE")
        _Visa.SendString(":MATH" & Chan & ":USERDEFINE:DEFINE '" & Operation & "'")
        _Visa.SendString(":MATH" & Chan & ":SCALE:MODE MANUAL")
        _Visa.SendString(":MATH" & Chan & ":SCALE:SENSITIVITY " & Vert)
        _Visa.SendString(":MATH" & Chan & ":SCALE:CENTER " & Pos)

    End Sub
#End Region

#Region "Measure Methodes"

    'Public Function MeasureAmplitude(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "AMPL", Area)
    '    End Function

    '    Public Function MeasureAverage(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "AVER", Area)
    '    End Function

    '    Public Function MeasureAvgPeriod(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "AVGP", Area)
    '    End Function

    '    Public Function MeasureBandwith(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "BWID", Area)
    '    End Function

    '    Public Function MeasureDT(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "DT", Area)
    '    End Function

    '    Public Function MeasureDutycycle(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "DUTY", Area)
    '    End Function

    '    Public Function MeasureEnumber(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "ENUM", Area)
    '    End Function

    '    Public Function MeasureFall(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "FALL", Area)
    '    End Function

    '    Public Function MeasureFrequency(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "FREQ", Area)
    '    End Function

    '    Public Function MeasureHigh(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "HIGH", Area)
    '    End Function

    '    Public Function MeasureLow(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "LOW", Area)
    '    End Function

    '    Public Function MeasureMaximum(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "MAX", Area)
    '    End Function

    '    Public Function MeasureMinimum(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "MIN", Area)
    '    End Function

    '    Public Function MeasureNOvershoot(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "NOV", Area)
    '    End Function

    '    Public Function MeasureNWith(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "NWID", Area)
    '    End Function

    '    Public Function MeasurePeriod(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "PER", Area)
    '    End Function

    '    Public Function MeasurePNumber(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "PNUM", Area)
    '    End Function

    '    Public Function MeasurePOvershoot(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "POV", Area)
    '    End Function

    '    Public Function MeasurePeakToPeak(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "PTOP", Area)
    '    End Function

    '    Public Function MeasurePODL(ByVal bit As Integer)
    '        Try
    '            _Visa.SendString(":MEASURE:PODL" & bit & "?")
    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '            Throw ex
    '        End Try
    '    End Function

    '    Public Function MeasurePWidth(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "PWID", Area)
    '    End Function

    '    Public Function MeasureRise(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "RISE", Area)
    '    End Function

    '    Public Function MeasureRMS(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "RMS", Area)
    '    End Function

    '    Public Function MeasureSDeviation(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "SDEV", Area)
    '    End Function

    '    Public Function MeasureTY1Integ(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "TY1I", Area)
    '    End Function

    '    Public Function MeasureTY2Integ(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "TY2I", Area)
    '    End Function

    '    Public Function MeasureV1(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "V1", Area)
    '    End Function

    '    Public Function MeasureV2(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Meas(Chan, "V2", Area)
    '    End Function

    '    Public Function MeasureMathMax(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return MeasMath(Chan, "MAX", Area)
    '    End Function

    '    Public Function MeasureMathMin(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return MeasMath(Chan, "MIN", Area)
    '    End Function

    'Private Function Meas(ByVal Chan As Integer, ByVal Parameter As String, Optional ByVal Area As Integer = 1) As Single

    '    'Chan = 1...8
    '    'Parameter = AMPLitude  | AVERage  | AVGFreq   | AVGPeriod | BWIDth
    '    '            DELay      | DT       | DUTYcycle | ENUMber   | FALL
    '    '            FREQuency  | HIGH     | LOW       | MAXimum   | MINimum
    '    '            NOVershoot | NWIDth   | PERiod    | PNUMber   | POVershoot
    '    '            PTOPeak    | PWIDth   | RISE      | RMS       | SDEViation
    '    '            TY1Integ   | TY2Integ | V1        | V2
    '    'Area = 1 | 2
    '    Try
    '        _Visa.SendString("COMMUNICATE:HEADER OFF")
    '        _Visa.SendString("MEASURE:MODE ON")
    '        _Visa.SendString(":MEASURE:CHANNEL" & Chan & ":AREA" & Area & ":" & Parameter & ":STATE ON")
    '        Threading.Thread.Sleep(300)
    '        _Visa.SendString(":MEASURE:CHANNEL" & Chan & ":AREA" & Area & ":" & Parameter & ":VALUE?")

    '        Return MyBase.ReceiveString()

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Function

    'Private Function MeasMath(ByVal Chan As Integer, ByVal Parameter As String, Optional ByVal Area As Integer = 1) As Single
    '    'Chan = 1...4
    '    'Parameter = AMPLitude  | AVERage  | AVGFreq   | AVGPeriod | BWIDth
    '    '            DELay      | DT       | DUTYcycle | ENUMber   | FALL
    '    '            FREQuency  | HIGH     | LOW       | MAXimum   | MINimum
    '    '            NOVershoot | NWIDth   | PERiod    | PNUMber   | POVershoot
    '    '            PTOPeak    | PWIDth   | RISE      | RMS       | SDEViation
    '    '            TY1Integ   | TY2Integ | V1        | V2
    '    'Area = 1 | 2
    '    Try
    '        _Visa.SendString("COMMUNICATE:HEADER OFF")
    '        _Visa.SendString("MEASURE:MODE ON")
    '        _Visa.SendString(":MEASURE:MATH" & Chan & ":AREA" & Area & ":" & Parameter & ":STATE ON")
    '        Threading.Thread.Sleep(300)
    '        _Visa.SendString(":MEASURE:MATH" & Chan & ":AREA" & Area & ":" & Parameter & ":VALUE?")

    '        Return MyBase.ReceiveString()

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Function

    'Public Sub MeasureSetTRange(ByVal Range As Integer, ByVal Pos1 As Single, ByVal Pos2 As Single)
    '    'Range = 1 | 2
    '    'Pos1 = -5 ... 5
    '    'Pos2 = -5 ... 5
    '    Try
    '        _Visa.SendString("COMMUNICAT:HEADER OFF")
    '        _Visa.SendString(":MEASURE:TRANGE" & Range & " " & Pos1 & "," & Pos2)

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Sub

    'Public Sub MeasureSetRefLevel(ByVal Channel As Integer, ByVal Mode As String, ByVal Distal As Single, ByVal Mesial As Single, ByVal Proximal As Single)
    '    'Channel = 1 ... 8
    '    'Mode = "UNIT" | "PERCENT"
    '    'Distal / Mesial / Proximal = 0 ... 100% (step 1%) | Voltage level
    '    Try
    '        _Visa.SendString("COMMUNICAT:HEADER OFF")
    '        _Visa.SendString("MEASURE:CHANNEL" & Channel & ":DPROXIMAL:MODE " & Mode)

    '        If Mode = "UNIT" Then
    '            _Visa.SendString(":MEASURE:CHANNEL" & Channel & ":DPROXIMAL:UNIT " & Distal & "," & Mesial & "," & Proximal)
    '        ElseIf Mode = "PERCENT" Then
    '            _Visa.SendString(":MEASURE:CHANNEL" & Channel & ":DPROXIMAL:PERCENT " & Distal & "," & Mesial & "," & Proximal)
    '        End If

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Sub

    'Public Function MeasureDelay(ByVal MeasChan As Integer,
    '                      ByVal MeasSlope As String,
    '                      Optional ByVal Source As String = "TRIGGER",
    '                      Optional ByVal RefChan As Integer = 0,
    '                      Optional ByVal RefSlope As String = "") As Single
    '    'Source = "TRACE" | "TRIGGER"
    '    'If Source = Trigger then you don't need to define a RefChan and RefSlope
    '    Try
    '        _Visa.SendString("MEASURE:MODE ON")
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":DELAY:STATE ON")
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":AREA1:DELAY:MEASURE:SLOPE " & MeasSlope)
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":AREA1:DELAY:REFERENCE:SOURCE " & Source)
    '        If Source = "TRACE" Then
    '            _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":AREA1:DELAY:REFERENCE:TRACE " & RefChan)
    '            _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":AREA1:DELAY:REFERENCE:SLOPE " & RefSlope)
    '        End If
    '        Threading.Thread.Sleep(300)
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":DELAY:VALUE?")

    '        Return MyBase.ReceiveString()

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Function

    'Public Function MeasureDelayTriggerToAnalog(ByVal AnalogChan As Integer,
    '                                        ByVal Slope As String,
    '                                        Optional ByVal Area As Integer = 1) As Single
    '    'AnalogChan = 1...8
    '    'Slope = "RISE" | "FALL"
    '    'Area = 1 | 2
    '    Try
    '        _Visa.SendString("MEASURE:MODE ON")
    '        _Visa.SendString(":MEASURE:CHANNEL" & AnalogChan & ":DELAY:STATE ON")
    '        _Visa.SendString(":MEASURE:CHANNEL" & AnalogChan & ":AREA" & Area & ":DELAY:MEASURE:SLOPE " & Slope)
    '        _Visa.SendString(":MEASURE:CHANNEL" & AnalogChan & ":AREA" & Area & ":DELAY:REFERENCE:SOURCE TRIGGER")
    '        Threading.Thread.Sleep(300)
    '        _Visa.SendString(":MEASURE:CHANNEL" & AnalogChan & ":DELAY:VALUE?")

    '        Return MyBase.ReceiveString()

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Function

    'Public Function MeasureDelayTriggerToLogic(ByVal LogicChan As Integer,
    '                                       ByVal Slope As String,
    '                                       Optional ByVal Area As Integer = 1) As Single
    '    'LogicChan = 0...7
    '    'Slope = "RISE" | "FALL"
    '    'Area = 1 | 2
    '    Try
    '        _Visa.SendString("MEASURE:MODE ON")
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":DELAY:STATE ON")
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":AREA1:DELAY:MEASURE:SLOPE " & Slope)
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":AREA1:DELAY:REFERENCE:SOURCE TRIGGER")
    '        Threading.Thread.Sleep(300)
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":DELAY:VALUE?")

    '        Return MyBase.ReceiveString()

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Function

    'Public Function MeasureDelayAnalogToAnalog(ByVal RefChan As Integer,
    '                                       ByVal RefSlope As String,
    '                                       ByVal MeasChan As Integer,
    '                                       ByVal MeasSlope As String,
    '                                       Optional ByVal Area As Integer = 1) As Single
    '    'RefChan = 1...8
    '    'RefSlope = "RISE" | "FALL"
    '    'MeasChan = 1...8
    '    'MeasSlope = "RISE" | "FALL"
    '    'Area = 1 | 2
    '    Try
    '        _Visa.SendString("MEASURE:MODE ON")
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":DELAY:STATE ON")
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":AREA" & Area & ":DELAY:MEASURE:SLOPE " & MeasSlope)
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:SOURCE TRACE")
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:TRACE " & RefChan)
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:SLOPE " & RefSlope)
    '        Threading.Thread.Sleep(300)
    '        _Visa.SendString(":MEASURE:CHANNEL" & MeasChan & ":DELAY:VALUE?")

    '        Return MyBase.ReceiveString()

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Function

    'Public Function MeasureDelayAnalogToLogic(ByVal AnalogChan As Integer,
    '                                      ByVal AnalogSlope As String,
    '                                      ByVal LogicChan As Integer,
    '                                      ByVal LogicSlope As String,
    '                                      Optional ByVal Area As Integer = 1) As Single
    '    'AnalogChan = 1...8
    '    'AnalogSlope = "RISE" | "FALL"
    '    'LogicChan = 0...7
    '    'LogicSlope = "RISE" | "FALL"
    '    'Area = 1 | 2
    '    Try
    '        _Visa.SendString("MEASURE:MODE ON")
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":DELAY:STATE ON")
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":AREA" & Area & ":DELAY:MEASURE:SLOPE " & LogicSlope)
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":AREA" & Area & ":DELAY:REFERENCE:SOURCE TRACE")
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":AREA" & Area & ":DELAY:REFERENCE:TRACE " & AnalogChan)
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":AREA" & Area & ":DELAY:REFERENCE:SLOPE " & AnalogSlope)
    '        Threading.Thread.Sleep(300)
    '        _Visa.SendString(":MEASURE:PODL" & LogicChan & ":DELAY:VALUE?")

    '        Return MyBase.ReceiveString()

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Function

    'Public Function MeasureDelayLogicToLogic(ByVal RefChan As Integer,
    '                                      ByVal RefSlope As String,
    '                                      ByVal MeasChan As Integer,
    '                                      ByVal MeasSlope As String,
    '                                      Optional ByVal Area As Integer = 1) As Single
    '    'RefChan = 0...7
    '    'RefSlope = "RISE" | "FALL"
    '    'MeasChan = 0...7
    '    'MeasSlope = "RISE" | "FALL"
    '    'Area = 1 | 2
    '    Try
    '        _Visa.SendString("MEASURE:MODE ON")
    '        _Visa.SendString(":MEASURE:PODL" & MeasChan & ":DELAY:STATE ON")
    '        _Visa.SendString(":MEASURE:PODL" & MeasChan & ":AREA" & Area & ":DELAY:MEASURE:SLOPE " & MeasSlope)
    '        _Visa.SendString(":MEASURE:PODL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:SOURCE TRACE")
    '        _Visa.SendString(":MEASURE:PODL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:TRACE PODL" & RefChan)
    '        _Visa.SendString(":MEASURE:PODL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:SLOPE " & RefSlope)
    '        Threading.Thread.Sleep(300)
    '        _Visa.SendString(":MEASURE:PODL" & MeasChan & ":DELAY:VALUE?")

    '        Return MyBase.ReceiveString()

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Function

    'Public Sub MeasureClearAll(Optional ByVal Chan As Integer = 0)
    '    Dim i As Integer
    '    Try
    '        If Chan = 0 Then
    '            For i = 1 To 8
    '                _Visa.SendString("MEASURE:CHANNEL" & i & ":AREA1:ALL OFF")
    '            Next i
    '        Else
    '            _Visa.SendString("MEASURE:CHANNEL" & Chan & ":AREA1:ALL OFF")
    '        End If

    '    Catch ex As Exception
    '        _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '        Throw ex
    '    End Try
    'End Sub

#End Region

#Region "Cursor Methodes"

    Public Sub CursorSet(ByVal CursorType As String, Trace As Integer, ByVal Pos1 As Single, Optional Pos2 As Single = -6)
        'Pos1 / Pos2 = -5 ... +5
        'CursorType = DEGRee | HAVertical | HORizontal | MARKer | OFF | VERTical
        'Trace = 1 ... 8

        If Pos2 = -6 Then Pos2 = Pos1

        Select Case CursorType
            Case "DEGree"

            Case "HAVertical"

            Case "HORizontal"

            Case "VERTical"
                _Visa.SendString(":CURSOR:TY:TYPE VERTical")
                _Visa.SendString(":CURSOR:TY:VERTICAL:TRACE " & Trace)
                _Visa.SendString(":CURSOR:TY:VERTICAL:POSITION1 " & Pos1)
                _Visa.SendString(":CURSOR:TY:VERTICAL:POSITION2 " & Pos2)
            Case "OFF"

        End Select

    End Sub

    Public Function CURSOR_GET_VERTICAL_VALUE(ByVal Cursor As Integer) As Single

        _Visa.SendString(":CURSOR:TY:VERTICAL:V" & Cursor & ":VALUE?")
        Return _Visa.ReceiveValue()

    End Function
#End Region

#Region "Trigger Methodes"

    Public Sub TRIGGER_POSITION_SET_TEST(ByVal Pos As Single)
        _Visa.SendString(":TRIGGER:POSITION " & Pos)
    End Sub

    Public Function TRIGGER_GET_POSITION() As Single
        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString(":TRIGGER:POSITION?")
        Return _Visa.ReceiveValue
    End Function

    Public Sub TRIGGER_MODE_AUTO()
        _Visa.SendString(":TRIGGER:MODE AUTO")
    End Sub

    Public Sub TRIGGER_MODE_AUTO_LEVEL()
        _Visa.SendString(":TRIGGER:MODE ALEV")
    End Sub

    Public Sub TRIGGER_MODE_NORMAL()
        _Visa.SendString(":TRIGGER:MODE NORM")
    End Sub

    Public Sub TRIGGER_MODE_N_SINGLE()

        _Visa.SendString(":TRIGGER:SCOUNT 1")
        _Visa.SendString(":TRIGGER:MODE NSIN")

    End Sub

    'Private Sub SetTriggerMode(ByVal Mode As String)
    '    _Visa.SendString(":TRIGGER:MODE " & Mode)
    'End Sub

    Public Sub TriggerSourcAndLevel(ByVal Chan As Integer, ByVal Levl As Single)
        _Visa.SendString(":TRIGGER:SOURCE:CHANNEL" & Chan & ":LEVEL " & Levl)
    End Sub

    Public Sub TriggerEdgeConfig(ByVal Source As Integer,
                                    ByVal Slope As String,
                                    ByVal Level As Single,
                                    ByVal HFRejection As String,
                                    ByVal NoiseRejection As String)
        'Source = 1 ... 8
        'Slope = "RISE" | "FALL"
        'Level = Trigger Level
        'HFRejection = "OFF" | "15 KHz" | "20 MHz"
        'NoiseRejection = "LOW" | "HIGH"
        _Visa.SendString(":TRIGGER:ATRIGGER:SIMPLE:SOURCE " & Source)
        _Visa.SendString(":TRIGGER:ATRIGGER:SIMPLE:SLOPE " & Slope)
        _Visa.SendString(":TRIGGER:ATRIGGER:SIMPLE:LEVEL " & Level)
        _Visa.SendString(":TRIGGER:ATRIGGER:SIMPLE:HFREJECTION " & HFRejection)
        _Visa.SendString(":TRIGGER:ATRIGGER:SIMPLE:HYSTERESIS " & NoiseRejection)

    End Sub

    Public Sub TriggerEdgeConfigLogic(ByVal LogicBit As Integer, ByVal Slope As String)
        'LogicBit = 0 ... 7
        'Slope = "RISE" | "FALL"
        _Visa.SendString(":TRIGGER:ATRIGGER:SIMPLE:SOURCE PODL" & LogicBit)
        _Visa.SendString(":TRIGGER:ATRIGGER:SIMPLE:SLOPE " & Slope)

    End Sub

    Public Sub TriggerExternal()

        _Visa.SendString("TRIGGER:SIMPLE:SOURCE EXTERNAL")
        _Visa.SendString("TRIGGER:SIMPLE:LEVEL 2")
        _Visa.SendString("TRIGGER:MODE NORMAL")

    End Sub

    Private Sub SetTriggerActionSave()

        _Visa.SendString("FILE:DRIVE FLASHMEM")
        _Visa.SendString("FILE:SAVE:ASCII:TINFORMATION OFF")
        _Visa.SendString("FILE:SAVE:ASCII:TRACE 1")
        _Visa.SendString("FILE:SAVE:NAME 'PLD'")
        _Visa.SendString("TRIGGER:ACTION:SAVE ON ")
        _Visa.SendString("STOP")
        _Visa.SendString("TRIGGER:ACTION:START")

    End Sub
#End Region

#Region "Acquire Methodes"
    Public Sub ACQUIRE_MODE_AVERAGE(ByVal Samples As Integer)

        _Visa.SendString(":ACQUIRE:MODE AVERAGE")
        _Visa.SendString(":ACQUIRE:AVERAGE:COUNT " & Samples)

    End Sub

    Public Sub ACQUIRE_MODE_ENVELOPE()
        _Visa.SendString(":ACQUIRE:MODE ENVELOPE")
    End Sub

    Public Sub ACQUIRE_MODE_NORMAL()
        _Visa.SendString(":ACQUIRE:MODE NORMAL")
    End Sub

    Public Sub ACQUIRE_SET_RECORD_LENGTH(ByVal RecordLength As Long)
        _Visa.SendString(":ACQUIRE:RLENGTH " & RecordLength)
    End Sub

    Public Sub ACQUIRE_HIGHRES_ON()
        _Visa.SendString(":ACQUIRE:RESOLUTION ON")
    End Sub

    Public Sub ACQUIRE_HIGHRES_OFF()
        _Visa.SendString(":ACQUIRE:RESOLUTION OFF")
    End Sub

    Public Sub ACQUIRE_STOP()
        _Visa.SendString("STOP")
    End Sub

    Public Sub ACQUIRE_START()
        _Visa.SendString("START")
    End Sub

    Public Function ACQUIRE_SINGLE_START() As Integer

        _Visa.SendString("SSTART? 0")
        Return _Visa.ReceiveValue()

    End Function
#End Region

#Region "Waveform Methodes"
    Public Function WAVEFORM_LENGTH() As Single

        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString("WAVEFORM:LENGTH?")
        Return _Visa.ReceiveValue()

    End Function

    Public Function WAVEFORM_RANGE(ByVal Channel As Integer) As Single

        _Visa.SendString("WAVEFORM:TRACE " & Channel)
        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString("WAVEFORM:RANGE?")
        Return _Visa.ReceiveValue()

    End Function

    Public Function WAVEFORM_OFFSET(ByVal Channel As Integer) As Single

        _Visa.SendString("WAVEFORM:TRACE " & Channel)
        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString("WAVEFORM:OFFSET?")
        Return _Visa.ReceiveValue()

    End Function

    Public Function WAVEFORM_POSITION(ByVal Channel As Integer) As Single

        _Visa.SendString("WAVEFORM:TRACE " & Channel)
        _Visa.SendString("COMMUNICATE:HEADER OFF")
        _Visa.SendString("WAVEFORM:POSITION?")
        Return _Visa.ReceiveValue()

    End Function
#End Region

#Region "Zoom Methodes"
    Public Sub ZOOM_CONFIG(ByVal Display As UInteger, ByVal Format As String, ByVal Position As Single, ByVal Magnification As Single)
        'Display = 1 | 2
        'Format = "MAIN" | "DUAL" | "HEXa" | "OCTal" | "QUAD" | "SINGle" | "TRIad"
        'Position = -5 ... 5
        'Magnification = ???

        Me.ACQUIRE_STOP()
        _Visa.SendString(":ZOOM" & Display & ":ALLOCATION:ALL ON")
        _Visa.SendString(":ZOOM" & Display & ":DISPLAY ON")
        _Visa.SendString(":ZOOM" & Display & ":FORMAT " & Format)
        _Visa.SendString(":ZOOM" & Display & ":MAIN 50")
        _Visa.SendString(":ZOOM" & Display & ":POSITION " & Position)
        _Visa.SendString(":ZOOM" & Display & ":MAG " & Magnification)

    End Sub

    Public Sub ZOOM_SET_POSITION(ByVal Display As Integer, ByVal Mag As Single, ByVal PosRelative As Single)
        Dim Position As Single

        _Visa.SendString(":ZOOM" & Display & ":POSITION?")
        Position = _Visa.ReceiveString()
        Position = Position - PosRelative / Mag
        _Visa.SendString(":ZOOM" & Display & ":POSITION " & Position)

    End Sub

    Public Sub ZOOM_REMOVE(ByVal Display As Integer)
        'Display = 1 | 2
        _Visa.SendString(":ZOOM" & Display & ":DISPLAY OFF")
    End Sub
#End Region

#Region "Search Methodes"
    Public Sub SEARCH_QUALIFIED_EDGE(ByVal Channel As Integer, ByVal Slope As String, ByVal Level As Single, ByVal Condition As Boolean)
        'Channel = 1 ... 8
        'Slope = "BOTH" | "FALL" | "RISE"
        'Level = Voltage Level / Current Level
        'Condition = TRUE | FALSE
        _Visa.SendString(":SEARCH:TYPE QUALIFY")
        _Visa.SendString(":SEARCH:EDGE:SOURCE " & Channel)
        _Visa.SendString(":SEARCH:EDGE:SLOPE " & Slope)
        _Visa.SendString(":SEARCH:EDGE:LEVEL " & Level)
        _Visa.SendString(":SEARCH:QUALIFY:CONDITION " & Condition)
        _Visa.SendString(":SEARCH:EXECUTE")

    End Sub

#End Region

#Region "Logic Methodes"
    Public Sub LOGIC_SET_CONFIG(ByVal bit As Integer,
                         ByVal Label As String)
        'Bit = 0 ... 7
        'Label = whatever you want
        _Visa.SendString(":LOGIC:PODL:MODE ON")
        _Visa.SendString(":LOGIC:PODL:BIT" & bit & ":DISPLAY ON")
        _Visa.SendString(":LOGIC:PODL:BIT" & bit & ":LABEL '" & Label & "'")

    End Sub

    Public Sub LOGIC_SET_SIZE(ByVal Size As String)
        'Size = "LARGE" | "MIDIUM" | "SMALL"
        _Visa.SendString(":LOGIC:PODL:SIZE " & Size)
    End Sub

    Public Sub LOGIC_SET_POSITION(ByVal Position As String)
        'Position = -7 ... 39
        _Visa.SendString(":LOGIC:PODL:POSITION " & Position)
    End Sub

    Public Sub LOGIC_SET_THRESHOLD_TYPE(ByVal TType As String)
        'Threshold = "CMOS1" | "CMOS2" | "COMIS3" | "CMOS5" | "ECL" | "USER"
        _Visa.SendString(":LOGIC:PODL:ALL:TYPE " & TType)
    End Sub

    Public Sub LOGIC_SET_LEVEL(ByVal Level As Single)
        'Level = -10 ... 10 (in 0.1V steps)
        _Visa.SendString(":LOGIC:PODL:ALL:LEVEL " & Level)
    End Sub

    Public Sub LOGIC_ON(Optional ByVal bit As Integer = -1)
        'Bit = 0 ... 7

        If bit = -1 Then
            _Visa.SendString(":LOGIC:PODL:ALL:DISPLAY ON")
        Else
            _Visa.SendString(":LOGIC:PODL:BIT" & bit & ":DISPLAY ON")
        End If
    End Sub

    Public Sub LOGIC_OFF(Optional ByVal bit As Integer = -1)
        'Bit = 0 ... 7
        If bit = -1 Then
            _Visa.SendString(":LOGIC:PODL:ALL:DISPLAY OFF")
        Else
            _Visa.SendString(":LOGIC:PODL:BIT" & bit & ":DISPLAY OFF")
        End If
    End Sub
#End Region

#Region "Hardcopy Methodes"
    '    Public Sub PRINT_PICTURE_TO_FILE()     'druckt KO Display in die Datei bild.png im Verzeichnis C:\Temp
    '        Dim stat As ViStatus
    '        Dim dfltRM As ViSession
    '        Dim sesn As ViSession
    '        Dim retCount As Long
    '        Dim j As Long
    '        Dim strTempData As String


    '        stat = viOpenDefaultRM(dfltRM)
    '        stat = viOpen(dfltRM, "GPIB0::" & mAddr & "::INSTR", 0, 0, sesn)
    '        stat = viSetAttribute(sesn, VI_ATTR_TMO_VALUE, 5000)

    '        stat = viWrite(sesn, "STOP", 4, retCount) 'not really nessecary but for testreason ok

    '        stat = viWrite(sesn, "COMMUNICATE:HEADER OFF", 22, retCount)
    '        stat = viWrite(sesn, "IMAGE:MODE NORMAL", 17, retCount)
    '        stat = viWrite(sesn, "IMAGE:FORMAT PNG", 16, retCount)
    '        '  stat = viWrite(sesn, "IMAGE:TONE REVERSE", 18, retCount)
    '        stat = viWrite(sesn, "IMAGE:BACKGROUND NORMAL", 23, retCount)
    '        stat = viWrite(sesn, "IMAGE:SEND?", 11, retCount)
    '        Call Delay(4)

    '        stat = viReadToFile(sesn, "C:\TEMP\TempScopePlotData.bin", 600000, retCount)

    '        stat = viClose(sesn)
    '        stat = viClose(dfltRM)

    '        Open "C:\TEMP\TempScopePlotData.bin" For Binary As #1
    '  strTempData = Space(LOF(1))
    '  Get #1, , strTempData
    '  Close #1
    '  Kill("C:\TEMP\TempScopePlotData.bin")

    '        strTempData = Mid(strTempData, 11)

    '        Open Environ("TEMP") + "\Picture.png" For Binary As #1 'übergibt Daten vom Buffer in die Datei

    '        For j = 1 To Len(strTempData)
    '            Put #1, j, Asc(Mid(strTempData, j, 1)) 'Convert each char to byte and write it in file
    '        Next j

    '        Close #1

    '  Call Delay(1)
    '    End Sub

    '    Public Function GET_WAVEFORM_DATA(ByVal Channel As Integer) As Single()
    '        'Channel 1 to 8
    '        Dim stat As ViStatus, dfltRM As ViSession, sesn As ViSession
    '        Dim retCount As Long
    '        Const MaxCnt = 62500
    '        Dim ReadBuf As String * MaxCnt
    '  Dim StartPnt As Long, EndPnt As Long, TempEndPnt As Long
    '        Const PntStep As Long = 1000
    '        Dim strStart As String, strEnd As String, strTemp As String, strRead As String
    '        Dim arrReadData() As Single
    '        Dim arrByte() As Byte, Rng As Single, Offs As Single, Pos As Single
    '        Dim i As Long, j As Long, x As Long

    '        StartPnt = 0
    '        x = 0

    '        EndPnt = Me.WAVEFORM_LENGTH - 1
    '        Pos = Me.WAVEFORM_POSITION(Channel)
    '        Rng = Me.WAVEFORM_RANGE(Channel)
    '        Offs = Me.WAVEFORM_OFFSET(Channel)
    '        ReDim arrReadData(0)

    '        'arrReadData(0) = Application.WorksheetFunction.Substitute(Trim(Scope.Label(Channel)), """", "")

    '        stat = viOpenDefaultRM(dfltRM)
    '        stat = viOpen(dfltRM, "GPIB0::" & mAddr & "::INSTR", 0, 0, sesn)
    '        'stat = viOpen(dfltRM, "TCPIP0::" & mIPAddr & "::inst0::INSTR", 0, 0, sesn)
    '        stat = viSetAttribute(sesn, VI_ATTR_TMO_VALUE, 5000)

    '        stat = viWrite(sesn, "WAVEFORM:TRACE " & Channel, 16, retCount)
    '        stat = viWrite(sesn, "WAVEFORM:FORMAT RBYTE", 21, retCount)

    '        For i = StartPnt To EndPnt Step PntStep
    '            TempEndPnt = StartPnt + i + PntStep - 1
    '            strStart = "WAVEFORM:START " & StartPnt + i
    '            strEnd = "WAVEFORM:END " & TempEndPnt
    '            stat = viWrite(sesn, strStart, Len(strStart), retCount)
    '            stat = viWrite(sesn, strEnd, Len(strEnd), retCount)
    '            stat = viWrite(sesn, "WAVEFORM:SEND?", 14, retCount)
    '            stat = viBufRead(sesn, ReadBuf, MaxCnt, retCount)
    '            strTemp = Left(ReadBuf, retCount)
    '            strRead = Right(strTemp, Len(strTemp) - 11)
    '            arrByte = StrConv(strRead, vbFromUnicode)

    '            For j = 0 To UBound(arrByte) - 1
    '                ReDim Preserve arrReadData(x)
    '                arrReadData(x) = (Rng * (arrByte(j) - Pos) / 25) + Offs
    '                x = x + 1
    '            Next j

    '            If TempEndPnt = EndPnt Then GoTo EndReading
    '        Next i

    'EndReading:

    '        stat = viClose(sesn)
    '        stat = viClose(dfltRM)

    '        GET_WAVEFORM_DATA = arrReadData

    '    End Function

#End Region
End Class
