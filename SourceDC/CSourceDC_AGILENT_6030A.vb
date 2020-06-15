'Class CSourceDC_Agilent_6030A
'07.04.2019, A. Zahler
'15.06.2020, Jabe
'Compatible Instruments:
'- Agilent/HP 6030A (single channel)
Imports Ivi.Visa

Public Class CSourceDC_AGILENT_6030A
    Inherits BCSourceDC
    Implements IDevice
    Implements ISource_DC


#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 200
        CurrentMax = 17
        PowerMax = 1200
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize
        Visa.SendString("SYST:LANG TMSL")
        Call cHelper.Delay(1)
        Visa.SendString("*RST;*CLS")
        Visa.SendString(":OUTPUT:STATE OFF")
        Visa.SendString("CURRENT 0")
        Visa.SendString("VOLTAGE 0")
    End Sub
#End Region

#Region "Interface Methodes ISource_DC"

    Public Overrides Sub SetOutputON() Implements ISource_DC.SetOutputON
        MyBase.SetOutputON()
        '_Visa.SendString(":OUTPUT:STATE ON")
    End Sub

    Public Overrides Sub SetOutputOFF() Implements ISource_DC.SetOutputOFF
        MyBase.SetOutputOFF()
        '_Visa.SendString(":OUTPUT:STATE OFF")
    End Sub

    Public Overrides Sub SetVoltage(ByVal Voltage As Single, CurrentLim As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage
        MyBase.SetVoltage(Voltage, CurrentLim, SetOutON)
        'If Voltage > _VoltageMax Then Voltage = _VoltageMax
        'If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        'If Voltage = 0 Then
        '    _Visa.SendString("VOLTAGE " & Voltage)
        '    _Visa.SendString(":OUTPUT:STATE OFF")
        'Else
        '    _Visa.SendString("SOURCE:CURRENT " & CurrentLim)
        '    _Visa.SendString("VOLTAGE " & Voltage)
        '    If SetOutON Then _Visa.SendString(":OUTPUT:STATE ON")
        'End If

    End Sub

    Public Overrides Sub SetVoltage(Voltage As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage
        MyBase.SetVoltage(Voltage, SetOutON)
        'If Voltage > _VoltageMax Then Voltage = _VoltageMax

        'If Voltage = 0 Then
        '    _Visa.SendString("VOLTAGE " & Voltage)
        '    _Visa.SendString(":OUTPUT:STATE OFF")
        'Else
        '    _Visa.SendString("VOLTAGE " & Voltage)
        'End If
    End Sub

    Public Overrides Sub SetCurrentLim(CurrentLim As Single) Implements ISource_DC.SetCurrentLim
        MyBase.SetCurrentLim(CurrentLim)

        'If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        '_Visa.SendString("CURRENT " & CurrentLim)
    End Sub

#End Region

End Class
