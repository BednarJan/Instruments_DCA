'Class CSourceDC_SORENSEN_60_1
'02.05.2019, J. Bednar  
'12.06.2020, JaBe
'Compatible Instruments:
'- Sorensen DLM 60-10(single channel)
Imports Ivi.Visa

Public Class BCSourceDC
    Implements IDevice
    Implements ISource_DC

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public Property VoltageMax As Single = 60 Implements ISource_DC.VoltageMax
    Public Property CurrentMax As Single = 10 Implements ISource_DC.CurrentMax
    Public Property PowerMax As Single = 600 Implements ISource_DC.PowerMax
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

    Public Overridable Sub Initialize() Implements IDevice.Initialize
        _Visa.SendString("SYST:LANG TMSL")
        Call cHelper.Delay(1)
        _Visa.SendString("*RST;*CLS")
        _Visa.SendString(":OUTPUT:STATE OFF")
        _Visa.SendString("CURRENT 0")
        _Visa.SendString("VOLTAGE 0")
    End Sub
#End Region

#Region "Interface Methodes ISource_DC"

    Public Overridable Sub SetOutputON() Implements ISource_DC.SetOutputON
        _Visa.SendString(":OUTPUT:STATE ON")
    End Sub

    Public Overridable Sub SetOutputOFF() Implements ISource_DC.SetOutputOFF
        _Visa.SendString(":OUTPUT:STATE OFF")
    End Sub

    Public Overridable Sub SetVoltage(ByVal Voltage As Single, CurrentLim As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage

        If Voltage > _VoltageMax Then Voltage = _VoltageMax

        If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        If Voltage = 0 Then
            _Visa.SendString("VOLTAGE " & Voltage)
            _Visa.SendString(":OUTPUT:STATE OFF")
        Else
            _Visa.SendString("SOURCE:CURRENT " & CurrentLim)
            _Visa.SendString("SOURCE:VOLTAGE " & Voltage)
            If SetOutON Then _Visa.SendString(":OUTPUT:STATE ON")
        End If

    End Sub

    Public Overridable Sub SetVoltage(Voltage As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage

        SetVoltage(Voltage, _CurrentMax, SetOutON)

    End Sub

    Public Overridable Sub SetCurrentLim(CurrentLim As Single) Implements ISource_DC.SetCurrentLim

        If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        _Visa.SendString("CURRENT " & CurrentLim)

    End Sub

#End Region

End Class
