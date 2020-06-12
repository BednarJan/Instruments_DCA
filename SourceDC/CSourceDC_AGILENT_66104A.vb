'Class CSourceDC_AGILENT_66104A
'09.04.2019, A. Zahler
'12.06.2020, JaBe
'Compatible Instruments:
'- AGILENT 66000A (8 slot mainframe) & 66104A (power module)
'- Only 1 pc. mainframe 66000A with 1 pc. module 66104A in Uster
'- This equipment will be handled as a single channel source
'- Module selection by secondary GPIB address
'- Secondary GPIB address ignored because module in slot 1 (= GPIB addr sec 0)
Imports Ivi.Visa

Public Class CSourceDC_AGILENT_66104A
    Inherits BCSourceDC
    Implements IDevice
    Implements ISource_DC

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize
        Visa.SendString("*RST;*CLS")
        Visa.SendString("CURR 0")
        Visa.SendString("VOLT 0")
        Visa.SendString(":OUTP:STAT ON")
    End Sub
#End Region

#Region "Interface Methodes ISource_DC"
    Public Overrides Sub SetOutputON() Implements ISource_DC.SetOutputON
        MyBase.SetOutputON()
        '_Visa.SendString("OUTP:STAT ON")
    End Sub

    Public Overrides Sub SetOutputOFF() Implements ISource_DC.SetOutputOFF
        MyBase.SetOutputOFF()
        '_Visa.SendString("OUTP:STAT OFF")
    End Sub

    Public Overrides Sub SetVoltage(ByVal Voltage As Single, CurrentLim As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage
        MyBase.SetVoltage(Voltage, CurrentLim, SetOutON)
        'If Voltage > _VoltageMax Then Voltage = _VoltageMax
        'If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        'If Voltage = 0 Then
        '    _Visa.SendString("VOLT " & Voltage)
        '    _Visa.SendString("OUTP:STAT OFF")
        'Else
        '    _Visa.SendString("CURR " & CurrentLim)
        '    _Visa.SendString("VOLT " & Voltage)
        '    If SetOutON Then _Visa.SendString("OUTP:STAT ON")
        'End If

    End Sub

    Public Overrides Sub SetVoltage(Voltage As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage
        MyBase.SetVoltage(Voltage, SetOutON)
        'If Voltage > _VoltageMax Then Voltage = _VoltageMax

        'If Voltage = 0 Then
        '    _Visa.SendString("VOLT " & Voltage)
        '    _Visa.SendString("OUTP:STAT OFF")
        'Else
        '    _Visa.SendString("VOLT " & Voltage)
        'End If

    End Sub

    Public Overrides Sub SetCurrentLim(CurrentLim As Single) Implements ISource_DC.SetCurrentLim
        MyBase.SetCurrentLim(CurrentLim)
        'If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        '_Visa.SendString("CURR " & CurrentLim)

    End Sub

#End Region

End Class
