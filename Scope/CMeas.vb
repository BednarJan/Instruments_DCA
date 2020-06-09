Public Class CMeas

    '    Inherits BCVisaDev
    '    Inherits CScope_Yokogawa_DLM4038

    '    Implements IScope

    '    Private _ErrorLogger As CErrorLogger
    '    Private _strvisa_adr As String = String.Empty

    '#Region "Dictionaries"
    '    Private DictMeasFunc As New Dictionary(Of CScope_Helper.EMeasParam, String) From {
    '    {CScope_Helper.EMeasParam.AMPLitude, "AMPL"},
    '    {CScope_Helper.EMeasParam.AVERage, "AVER"},
    '    {CScope_Helper.EMeasParam.AVGFreq, "AVGF"},
    '    {CScope_Helper.EMeasParam.BWIDth, "BWID"},
    '    {CScope_Helper.EMeasParam.DELay, "DEL"},
    '    {CScope_Helper.EMeasParam.DT, "DT"},
    '    {CScope_Helper.EMeasParam.DUTYcycle, "DUTY"},
    '    {CScope_Helper.EMeasParam.ENUMber, "ENUM"},
    '    {CScope_Helper.EMeasParam.FALL, "FALL"},
    '    {CScope_Helper.EMeasParam.FREQuency, "FREQ"},
    '    {CScope_Helper.EMeasParam.HIGH, "HIGH"},
    '    {CScope_Helper.EMeasParam.LOW, "LOW"},
    '    {CScope_Helper.EMeasParam.MAXimum, "MAX"},
    '    {CScope_Helper.EMeasParam.MINimum, "MIN"},
    '    {CScope_Helper.EMeasParam.NOVershoot, "NOV"},
    '    {CScope_Helper.EMeasParam.NWIDth, "NWID"},
    '    {CScope_Helper.EMeasParam.PERiod, "PER"},
    '    {CScope_Helper.EMeasParam.PNUMber, "PNUM"},
    '    {CScope_Helper.EMeasParam.POVershoot, "POV"},
    '    {CScope_Helper.EMeasParam.PTOPeak, "PTOP"},
    '    {CScope_Helper.EMeasParam.PWIDth, "PWID"},
    '    {CScope_Helper.EMeasParam.RISE, "RISE"},
    '    {CScope_Helper.EMeasParam.RMS, "RMS"},
    '    {CScope_Helper.EMeasParam.SDEViation, "SDEV"},
    '    {CScope_Helper.EMeasParam.TY1Integ, "TY1I"},
    '    {CScope_Helper.EMeasParam.TY2Integ, "TY2I"},
    '    {CScope_Helper.EMeasParam.V1, "V1"},
    '    {CScope_Helper.EMeasParam.V2, "V2"}
    '    }
    '#End Region

    '#Region "Constructor"
    '    Public Sub New(strVisa_Adr As String, oErrorLogger As CErrorLogger)
    '        MyBase.New(strVisa_Adr, oErrorLogger)
    '        _ErrorLogger = oErrorLogger
    '        _strvisa_adr = strVisa_Adr
    '    End Sub
    '#End Region

    '    Public Function PODL(ByVal bit As Integer)
    '        Try
    '            MyBase.Send("::PODL" & bit & "?")
    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function


    '    Private Function Meas(ByVal Channel As CScope_Helper.EChannel, ByVal Parameter As CScope_Helper.EMeasParam, Optional ByVal Area As CScope_Helper.EArea = 1) As Single
    '        Chan = 1...8
    '        Parameter = AMPLitude  | AVERage  | AVGFreq   | AVGPeriod | BWIDth
    '                    DELay      | DT       | DUTYcycle | ENUMber   | FALL
    '                    FREQuency  | HIGH     | LOW       | MAXimum   | MINimum
    '                    NOVershoot | NWIDth   | PERiod    | PNUMber   | POVershoot
    '                    PTOPeak    | PWIDth   | RISE      | RMS       | SDEViation
    '                    TY1Integ   | TY2Integ | V1        | V2
    '        Area = 1 | 2
    '        Dim Param As String = String.Empty

    '        Param = DictMeasFunc.Item(Parameter)

    '        Try
    '            MyBase.Send("COMMUNICATE:HEADER OFF")
    '            MyBase.Send("MEASURE:MODE ON")
    '            MyBase.Send(":MEASURE:CHANNEL" & Channel & ":AREA" & Area & ":" & Param & ":STATE ON")
    '            Threading.Thread.Sleep(300)
    '            MyBase.Send(":MEASURE:CHANNEL" & Channel & ":AREA" & Area & ":" & Param & ":VALUE?")

    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function


    '    Public Function MathMax(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Math(Chan, "MAX", Area)
    '    End Function

    '    Public Function MathMin(ByVal Chan As Integer, Optional ByVal Area As Integer = 1) As Single
    '        Return Math(Chan, "MIN", Area)
    '    End Function

    '    Private Function Math(ByVal Chan As Integer, ByVal Parameter As String, Optional ByVal Area As Integer = 1) As Single
    '        Chan = 1...4
    '        Parameter = AMPLitude  | AVERage  | AVGFreq   | AVGPeriod | BWIDth
    '                    DELay      | DT       | DUTYcycle | ENUMber   | FALL
    '                    FREQuency  | HIGH     | LOW       | MAXimum   | MINimum
    '                    NOVershoot | NWIDth   | PERiod    | PNUMber   | POVershoot
    '                    PTOPeak    | PWIDth   | RISE      | RMS       | SDEViation
    '                    TY1Integ   | TY2Integ | V1        | V2
    '        Area = 1 | 2
    '        Try
    '            MyBase.Send("COMMUNICATE:HEADER OFF")
    '            MyBase.Send("MEASURE:MODE ON")
    '            MyBase.Send(":MEASURE:MATH" & Chan & ":AREA" & Area & ":" & Parameter & ":STATE ON")
    '            Threading.Thread.Sleep(300)
    '            MyBase.Send(":MEASURE:MATH" & Chan & ":AREA" & Area & ":" & Parameter & ":VALUE?")

    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function

    '    Public Sub SetTRange(ByVal Range As Integer, ByVal Pos1 As Single, ByVal Pos2 As Single)
    '        Range = 1 | 2
    '        Pos1 = -5... 5
    '        Pos2 = -5... 5
    '        Try
    '            MyBase.Send("COMMUNICAT:HEADER OFF")
    '            MyBase.Send(":MEASURE:TRANGE" & Range & " " & Pos1 & "," & Pos2)

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '            Throw ex
    '        End Try
    '    End Sub

    '    Public Sub SetRefLevel(ByVal Channel As Integer, ByVal Mode As String, ByVal Distal As Single, ByVal Mesial As Single, ByVal Proximal As Single)
    '        Channel = 1... 8
    '        Mode = "UNIT" | "PERCENT"
    '        Distal / Mesial / Proximal = 0... 100% (step 1%) | Voltage level
    '        Try
    '            MyBase.Send("COMMUNICAT:HEADER OFF")
    '            MyBase.Send("MEASURE:CHANNEL" & Channel & ":DPROXIMAL:MODE " & Mode)

    '            If Mode = "UNIT" Then
    '                MyBase.Send(":MEASURE:CHANNEL" & Channel & ":DPROXIMAL:UNIT " & Distal & "," & Mesial & "," & Proximal)
    '            ElseIf Mode = "PERCENT" Then
    '                MyBase.Send(":MEASURE:CHANNEL" & Channel & ":DPROXIMAL:PERCENT " & Distal & "," & Mesial & "," & Proximal)
    '            End If

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '            Throw ex
    '        End Try
    '    End Sub

    '    Public Function Delay(ByVal MeasChan As Integer,
    '                          ByVal MeasSlope As String,
    '                          Optional ByVal Source As String = "TRIGGER",
    '                          Optional ByVal RefChan As Integer = 0,
    '                          Optional ByVal RefSlope As String = "") As Single
    '        Source = "TRACE" | "TRIGGER"
    '        If Source = Trigger then you don't need to define a RefChan and RefSlope
    '        Try
    '            MyBase.Send("MEASURE:MODE ON")
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":DELAY:STATE ON")
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":AREA1:DELAY:MEASURE:SLOPE " & MeasSlope)
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":AREA1:DELAY:REFERENCE:SOURCE " & Source)
    '            If Source = "TRACE" Then
    '                MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":AREA1:DELAY:REFERENCE:TRACE " & RefChan)
    '                MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":AREA1:DELAY:REFERENCE:SLOPE " & RefSlope)
    '            End If
    '            Threading.Thread.Sleep(300)
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":DELAY:VALUE?")

    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function

    '    Public Function DelayTriggerToAnalog(ByVal AnalogChan As Integer,
    '                                            ByVal Slope As String,
    '                                            Optional ByVal Area As Integer = 1) As Single
    '        AnalogChan = 1...8
    '        Slope = "RISE" | "FALL"
    '        Area = 1 | 2
    '        Try
    '            MyBase.Send("MEASURE:MODE ON")
    '            MyBase.Send(":MEASURE:CHANNEL" & AnalogChan & ":DELAY:STATE ON")
    '            MyBase.Send(":MEASURE:CHANNEL" & AnalogChan & ":AREA" & Area & ":DELAY:MEASURE:SLOPE " & Slope)
    '            MyBase.Send(":MEASURE:CHANNEL" & AnalogChan & ":AREA" & Area & ":DELAY:REFERENCE:SOURCE TRIGGER")
    '            Threading.Thread.Sleep(300)
    '            MyBase.Send(":MEASURE:CHANNEL" & AnalogChan & ":DELAY:VALUE?")

    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function

    '    Public Function DelayTriggerToLogic(ByVal LogicChan As Integer,
    '                                           ByVal Slope As String,
    '                                           Optional ByVal Area As Integer = 1) As Single
    '        LogicChan = 0...7
    '        Slope = "RISE" | "FALL"
    '        Area = 1 | 2
    '        Try
    '            MyBase.Send("MEASURE:MODE ON")
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":DELAY:STATE ON")
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":AREA1:DELAY:MEASURE:SLOPE " & Slope)
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":AREA1:DELAY:REFERENCE:SOURCE TRIGGER")
    '            Threading.Thread.Sleep(300)
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":DELAY:VALUE?")

    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function

    '    Public Function DelayAnalogToAnalog(ByVal RefChan As Integer,
    '                                           ByVal RefSlope As String,
    '                                           ByVal MeasChan As Integer,
    '                                           ByVal MeasSlope As String,
    '                                           Optional ByVal Area As Integer = 1) As Single
    '        RefChan = 1...8
    '        RefSlope = "RISE" | "FALL"
    '        MeasChan = 1...8
    '        MeasSlope = "RISE" | "FALL"
    '        Area = 1 | 2
    '        Try
    '            MyBase.Send("MEASURE:MODE ON")
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":DELAY:STATE ON")
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":AREA" & Area & ":DELAY:MEASURE:SLOPE " & MeasSlope)
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:SOURCE TRACE")
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:TRACE " & RefChan)
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:SLOPE " & RefSlope)
    '            Threading.Thread.Sleep(300)
    '            MyBase.Send(":MEASURE:CHANNEL" & MeasChan & ":DELAY:VALUE?")

    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function

    '    Public Function DelayAnalogToLogic(ByVal AnalogChan As Integer,
    '                                          ByVal AnalogSlope As String,
    '                                          ByVal LogicChan As Integer,
    '                                          ByVal LogicSlope As String,
    '                                          Optional ByVal Area As Integer = 1) As Single
    '        AnalogChan = 1...8
    '        AnalogSlope = "RISE" | "FALL"
    '        LogicChan = 0...7
    '        LogicSlope = "RISE" | "FALL"
    '        Area = 1 | 2
    '        Try
    '            MyBase.Send("MEASURE:MODE ON")
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":DELAY:STATE ON")
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":AREA" & Area & ":DELAY:MEASURE:SLOPE " & LogicSlope)
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":AREA" & Area & ":DELAY:REFERENCE:SOURCE TRACE")
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":AREA" & Area & ":DELAY:REFERENCE:TRACE " & AnalogChan)
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":AREA" & Area & ":DELAY:REFERENCE:SLOPE " & AnalogSlope)
    '            Threading.Thread.Sleep(300)
    '            MyBase.Send(":MEASURE:PODL" & LogicChan & ":DELAY:VALUE?")

    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function

    '    Public Function DelayLogicToLogic(ByVal RefChan As Integer,
    '                                          ByVal RefSlope As String,
    '                                          ByVal MeasChan As Integer,
    '                                          ByVal MeasSlope As String,
    '                                          Optional ByVal Area As Integer = 1) As Single
    '        RefChan = 0...7
    '        RefSlope = "RISE" | "FALL"
    '        MeasChan = 0...7
    '        MeasSlope = "RISE" | "FALL"
    '        Area = 1 | 2
    '        Try
    '            MyBase.Send("MEASURE:MODE ON")
    '            MyBase.Send(":MEASURE:PODL" & MeasChan & ":DELAY:STATE ON")
    '            MyBase.Send(":MEASURE:PODL" & MeasChan & ":AREA" & Area & ":DELAY:MEASURE:SLOPE " & MeasSlope)
    '            MyBase.Send(":MEASURE:PODL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:SOURCE TRACE")
    '            MyBase.Send(":MEASURE:PODL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:TRACE PODL" & RefChan)
    '            MyBase.Send(":MEASURE:PODL" & MeasChan & ":AREA" & Area & ":DELAY:REFERENCE:SLOPE " & RefSlope)
    '            Threading.Thread.Sleep(300)
    '            MyBase.Send(":MEASURE:PODL" & MeasChan & ":DELAY:VALUE?")

    '            Return MyBase.ReceiveString()

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strvisa_adr)))
    '            Throw ex
    '        End Try
    '    End Function

    '    Public Sub ClearAll(Optional ByVal Chan As Integer = 0)
    '        Dim i As Integer
    '        Try
    '            If Chan = 0 Then
    '                For i = 1 To 8
    '                    MyBase.Send("MEASURE:CHANNEL" & i & ":AREA1:ALL OFF")
    '                Next i
    '            Else
    '                MyBase.Send("MEASURE:CHANNEL" & Chan & ":AREA1:ALL OFF")
    '            End If

    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
    '            Throw ex
    '        End Try
    '    End Sub
End Class

