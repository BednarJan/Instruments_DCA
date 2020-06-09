'Class CSourceDC_SORENSEN_60_1
'02.05.2019, J. Bednar  
'Compatible Instruments:
'- Sorensen DLM 60-10(single channel)
Imports Ivi.Visa

Public Class CSourceDC_SORENSEN_60_10
    Implements IDevice
    Implements ISource_DC

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMax As Single = 60 Implements ISource_DC.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 10 Implements ISource_DC.CurrentMax
    Public ReadOnly Property PowerMax As Single = 600 Implements ISource_DC.PowerMax
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
        _Visa.SendString("SYST:LANG TMSL")
        Call cHelper.Delay(1)
        _Visa.SendString("*RST;*CLS")
        _Visa.SendString(":OUTPUT:STATE ON")
        _Visa.SendString("CURRENT 0")
        _Visa.SendString("VOLTAGE 0")
    End Sub
#End Region

#Region "Interface Methodes ISource_DC"

    Public Sub SetOutputON() Implements ISource_DC.SetOutputON
        _Visa.SendString(":OUTPUT:STATE ON")
    End Sub

    Public Sub SetOutputOFF() Implements ISource_DC.SetOutputOFF
        _Visa.SendString(":OUTPUT:STATE OFF")
    End Sub

    Public Sub SetSource(ByVal Voltage As Single, CurrentLim As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetSource

        If Voltage > _VoltageMax Then Voltage = _VoltageMax
        If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        If Voltage = 0 Then
            _Visa.SendString("VOLTAGE " & Voltage)
            _Visa.SendString(":OUTPUT:STATE OFF")
        Else
            _Visa.SendString("SOURCE:CURRENT " & CurrentLim)
            _Visa.SendString("VOLTAGE " & Voltage)
            If SetOutON Then _Visa.SendString(":OUTPUT:STATE ON")
        End If

    End Sub

    Public Sub SetVoltage(Voltage As Single) Implements ISource_DC.SetVoltage

        If Voltage > _VoltageMax Then Voltage = _VoltageMax

        If Voltage = 0 Then
            _Visa.SendString("VOLTAGE " & Voltage)
            _Visa.SendString(":OUTPUT:STATE OFF")
        Else
            _Visa.SendString("VOLTAGE " & Voltage)
        End If
    End Sub

    Public Sub SetCurrentLim(CurrentLim As Single) Implements ISource_DC.SetCurrentLim

        If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        _Visa.SendString("CURRENT " & CurrentLim)
    End Sub

    'Public Function GetVoltage() As Single Implements ISource_DC.GetVolt
    '    Dim RetValue As String = String.Empty
    '        _Visa.SendString("MEAS:VOLT?")
    '        RetValue = _Visa.ReceiveString()

    '        RetValue = RetValue.TrimEnd(" ", "V", vbLf)  'Remove ending before convert to Single
    '        GetVoltage = cHelper.StringToSingle(RetValue)
    'End Function

    'Public Sub ClearProt() Implements ISource_DC.ClearProt
    '    _Visa.SendString("OUTP:PROT:CLEar")
    'End Sub

    'Public Function GetStatus() As Single Implements ISource_DC.GetStatus
    '        _Visa.SendString("STAT:QUES:COND?")                  '"SYST:ERR?")
    '        GetStatus = _Visa.ReceiveString()
    'End Function

#End Region

End Class
