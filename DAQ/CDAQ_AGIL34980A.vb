Imports Instruments
Imports Ivi.Visa

Public Class CDAQ_AGIL34980A
    Implements IDevice, IDAQ

    Private Enum MeasFunction
        VoltDC = 1
        VoltAC = 2
        CurrDC = 3
        CurrAC = 4
        Res = 5
        FRes = 6
        Freq = 7
        PERIOD = 8
        TempTC_B = 9
        TempTC_E = 10
        TempTC_J = 11
        TempTC_K = 12
        TempTC_N = 13
        TempTC_R = 14
        TempTC_S = 13
        TempTC_T = 14
    End Enum

    Public Enum OutputState
        OutON = 1
        OutOFF = 2
    End Enum

    Public Enum RelaisState
        Open = 1
        Close = 2
    End Enum

    Public Enum DIOWidth
        wByte = 1
        wWord = 2
        wLWord = 3
    End Enum

    Public Enum DIODrive
        Active = 1
        OCollector = 2
    End Enum

    Public Enum DIOMode
        OFF = 1
        Output = 2
        Input = 3
    End Enum

    Private _ErrorLogger As CErrorLogger = Nothing

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
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

    Public Sub Initialize() Implements IDevice.Initialize
        Dim Channel As Integer = 0

        _Visa.SendString("*CLS;*RST")

    End Sub

    Public Sub SndString(cmdStr As String) Implements IDevice.SendString
        _Visa.SendString(cmdStr)
    End Sub

    Public Function RecieveString() As String Implements IDevice.ReceiveString
        Return _Visa.ReceiveString
    End Function

#End Region

#Region "Interface Methodes IDAQ"
    '    Public Overrides Sub RST() Implements IDAQ.RST
    '        Try
    '            MyBase.RST()
    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, GPIB Addr = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Addr)))
    '            Throw ex
    '        End Try
    '    End Sub

    '    Public Overrides Sub CLS() Implements IDAQ.CLS
    '        Try
    '            MyBase.CLS()
    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, GPIB Addr = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Addr)))
    '            Throw ex
    '        End Try
    '    End Sub

    '    Public Overrides Function IDN() As String Implements IDAQ.IDN
    '        Dim sRetVal As String = String.Empty
    '        Try
    '            'sRetVal = MyBase.IDN()
    '            MyBase.Send("*IDN?")
    '            MyBase.WaitTime(1000)
    '            sRetVal = MyBase.ReceiveString
    '        Catch ex As Exception
    '            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, GPIB Addr = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Addr)))
    '            Throw ex
    '        End Try
    '        Return sRetVal
    '    End Function

    '    Public Function initialize() As Integer Implements IDAQ.initialize
    '        Dim nOK As Integer = 0
    '        Try
    '            RST()
    '        Catch ex As Exception
    '            nOK += 1
    '            Throw ex
    '        End Try
    '        Return nOK
    '    End Function

    Public Function Get_Volt_DC(Chan As Integer) As Single Implements IDAQ.Get_Volt_DC
        Return Me.GetDAQ(Chan, MeasFunction.VoltDC)
    End Function

    Public Function Get_Volt_AC(Chan As Integer) As Single Implements IDAQ.Get_Volt_AC
        Throw New NotImplementedException()
    End Function

    Public Function Get_Sample_Volt_DC(Chan As Integer, Sample As Integer) As Single Implements IDAQ.Get_Sample_Volt_DC
        Throw New NotImplementedException()
    End Function

    Public Function Get_Sample_Volt_Dyn(Chan As Integer, Sample As Integer) As Single Implements IDAQ.Get_Sample_Volt_Dyn
        Throw New NotImplementedException()
    End Function

    Public Function Get_Curr_DC(Chan As Integer) As Single Implements IDAQ.Get_Curr_DC
        Throw New NotImplementedException()
    End Function

    Public Function Get_Curr_AC(Chan As Integer) As Single Implements IDAQ.Get_Curr_AC
        Throw New NotImplementedException()
    End Function

    Public Function Get_Res(Chan As Integer) As Single Implements IDAQ.Get_Res
        Throw New NotImplementedException()
    End Function

    Public Function Get_FRes(Chan As Integer) As Single Implements IDAQ.Get_FRes
        Throw New NotImplementedException()
    End Function

    Public Function Get_FReq(Chan As Integer) As Single Implements IDAQ.Get_FReq
        Throw New NotImplementedException()
    End Function
#End Region

#Region "Special Methodes 34980A"
    Private Function GetDAQ(ByVal Chan As Integer, ByVal Func As MeasFunction) As Single
        Dim CmdString As String = String.Empty
        Dim RetVal As Single = Single.NaN

        Chan = Chan + 1000

        Select Case Func
            Case MeasFunction.VoltDC
                CmdString = "MEAS:VOLT:DC? "
            Case MeasFunction.VoltAC
                CmdString = "MEAS:VOLT:AC? "
            Case MeasFunction.CurrDC 'Channels 21 and 22 on 34901A module
                CmdString = "MEAS:CURR:DC? "
            Case MeasFunction.CurrAC 'Channels 21 and 22 on 34901A module
                CmdString = "MEAS:CURR:AC? "
            Case MeasFunction.Res
                CmdString = "MEAS:RES? "
            Case MeasFunction.FRes
                CmdString = "MEAS:FRES? "
            Case MeasFunction.Freq
                CmdString = "MEAS:FREQ? "
            Case MeasFunction.TempTC_B
                CmdString = "MEAS:TEMP? TC,B,"
            Case MeasFunction.TempTC_E
                CmdString = "MEAS:TEMP? TC,E,"
            Case MeasFunction.TempTC_J
                CmdString = "MEAS:TEMP? TC,J,"
            Case MeasFunction.TempTC_K
                CmdString = "MEAS:TEMP? TC,K,"
            Case MeasFunction.TempTC_N
                CmdString = "MEAS:TEMP? TC,N,"
            Case MeasFunction.TempTC_R
                CmdString = "MEAS:TEMP? TC,R,"
            Case MeasFunction.TempTC_S
                CmdString = "MEAS:TEMP? TC,S,"
            Case MeasFunction.TempTC_T
                CmdString = "MEAS:TEMP? TC,T,"
        End Select

        CmdString = CmdString & "(@" & Chan & ")"

        _Visa.SendString(CmdString)
        Return _Visa.ReceiveValue()

    End Function

    Public Function GetDAQSample(ByVal Chan As Integer, ByVal Sample As Integer) As Single
        Dim RetVal As Single = Single.NaN

        Chan = Chan + 1000

        _Visa.SendString("CONF:VOLT:DC AUTO ,DEF, (@" & Chan & ")")
        _Visa.SendString("ROUTe:SCAN (@)")
        _Visa.SendString("ROUTe:SCAN (@" & Chan & ")")
        _Visa.SendString("SAMPle:COUNt " & Sample)
        _Visa.SendString("INITiate")
        cHelper.Delay(1)
        _Visa.SendString("CALC:AVER:AVER? (@" & Chan & ")")

        Return _Visa.ReceiveValue

    End Function

    Public Function GetDAQSampleDyn(ByVal Chan As Integer, ByVal Sample As Integer) As Single
        Dim RetVal As Single = Single.NaN

        Chan = Chan + 1000

        _Visa.SendString("CONF:VOLT:DC AUTO ,DEF, (@" & Chan & ")")
        _Visa.SendString("ROUTe:SCAN (@)")
        _Visa.SendString("ROUTe:SCAN (@" & Chan & ")")
        _Visa.SendString("SAMPle:COUNt " & Sample)
        _Visa.SendString("SENS:VOLT:APER 0.2,(@" & Chan & ")")
        _Visa.SendString("INITiate")
        cHelper.Delay(5)
        _Visa.SendString("CALC:AVER:AVER? (@" & Chan & ")")
        RetVal = _Visa.ReceiveValue

        _Visa.SendString("ROUT:CHAN:DELAY:AUTO ON, (@" & Chan & ")")

        Return RetVal

    End Function

    Public Sub SetRelais(ByVal Chan As Integer, State As RelaisState)
        Dim CmdString As String

        Chan = Chan + 4000

        If State = RelaisState.Close Then CmdString = "ROUT:CLOS " Else CmdString = "ROUT:OPEN "
        CmdString = CmdString & "(@" & Chan & ")"

        _Visa.SendString(CmdString)

    End Sub

    Public Sub SetAnalogOut(ByVal Chan As Integer, ByVal Volt As Single)
        Dim CmdString As String = String.Empty

        Chan = Chan + 2000

        If Volt > 16 Then Volt = 16
        If Volt < -16 Then Volt = -16

        CmdString = "SOUR:VOLT " & Volt & ",(@" & Chan & ")"
        _Visa.SendString(CmdString)
        CmdString = "OUTP:STAT ON,(@" & Chan & ")"
        _Visa.SendString(CmdString)

    End Sub

    Public Sub SetAnalogOutState(ByVal Chan As Integer, ByVal State As OutputState)
        Dim CmdString As String

        Chan = Chan + 2000

        If State = OutputState.OutON Then CmdString = "OUTP:STAT ON" Else CmdString = "OUTP:STAT OFF"
        CmdString = CmdString & "(@" & Chan & ")"

        _Visa.SendString(CmdString)

    End Sub

    Public Sub SetBitDigitalOut(ByVal Channel As Integer, ByVal BitNo As Byte, ByVal Level As Byte)
        Dim CmdString As String = String.Empty
        Dim RetVal As String = String.Empty
        Dim OldByte As Byte, NewByte As Byte, Mask As Integer

        Channel = 3000 + Channel

        'Read current state
        CmdString = "SOURCE:DIGITAL:DATA:BYTE? (@" & Channel & ")"

        _Visa.SendString(CmdString)
        RetVal = _Visa.ReceiveString

        OldByte = CByte(RetVal)

        'Build byte to send
        If Level = 0 Then 'Set bit to logic low
            Mask = Not (2 ^ BitNo)
            NewByte = OldByte And Mask
        Else              'Set bit to logic high
            Mask = 2 ^ BitNo
            NewByte = OldByte Or Mask
        End If

        'Send modified byte
        CmdString = "SOURCE:DIGITAL:DATA:BYTE " & NewByte & ",(@" & Channel & ")"
        _Visa.SendString(CmdString)

    End Sub

    Public Sub ConfigDIO(ByVal ByteNo As Integer, ByVal WIDT As DIOWidth, ByVal DRIV As DIODrive, ByVal MODE As DIOMode)
        Dim CmdStringWIDT As String = String.Empty
        Dim CmdStringDRIV As String = String.Empty
        Dim CmdStringMODE As String = String.Empty
        Dim CmdStringPOL As String = String.Empty

        ByteNo = 3000 + ByteNo

        Select Case WIDT
            Case DIOWidth.wByte
                CmdStringWIDT = "CONF:DIG:WIDT BYTE, (@" & ByteNo & ")"
            Case DIOWidth.wWord
                CmdStringWIDT = "CONF:DIG:WIDT WORD, (@" & ByteNo & ")"
            Case DIOWidth.wLWord
                CmdStringWIDT = "CONF:DIG:WIDT LWORd, (@" & ByteNo & ")"
        End Select

        Select Case MODE
            Case DIOMode.OFF
                CmdStringMODE = "CONF:DIG:STAT OFF, (@" & ByteNo & ")"
            Case DIOMode.Output
                CmdStringMODE = "CONF:DIG:STAT OUTPut, (@" & ByteNo & ")"
            Case DIOMode.Input
                CmdStringMODE = "CONF:DIG:STAT INPut, (@" & ByteNo & ")"
        End Select

        Select Case DRIV
            Case DIODrive.Active
                CmdStringDRIV = "CONF:DIG:DRIV ACTiv, (@" & ByteNo & ")"
            Case DIODrive.OCollector
                CmdStringDRIV = "CONF:DIG:DRIV OCOLlector, (@" & ByteNo & ")"
        End Select

        CmdStringPOL = "CONF:DIG:POL NORM, (@" & ByteNo & ")"

        _Visa.SendString(CmdStringWIDT)
        _Visa.SendString(CmdStringMODE)
        _Visa.SendString(CmdStringPOL)
        _Visa.SendString(CmdStringDRIV)

    End Sub
#End Region

End Class
