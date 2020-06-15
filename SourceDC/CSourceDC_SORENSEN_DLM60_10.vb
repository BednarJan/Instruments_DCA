'Class CSourceDC_SORENSEN_60_1
'12.06.2020, J. Bednar  
'Compatible Instruments:
'- Sorensen DLM 60-10(single channel)
Imports Ivi.Visa

Public Class CSourceDC_SORENSEN_DLM60_10
    Inherits BCSourceDC
    Implements IDevice
    Implements ISource_DC


#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

        VoltageMax = 60
        CurrentMax = 10
        PowerMax = 600

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
