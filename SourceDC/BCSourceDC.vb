'Class BCSourceDC common for all DC ources,  inherited fro the BCDevice
'02.05.2019, J. Bednar  
'12.06.2020, JaBe
'Base clase for the DC sources
Imports Ivi.Visa

Public Class BCSourceDC
    Inherits BCDevice
    Implements ISource_DC

#Region "Shorthand Properties"

    Public Property VoltageMax As Single Implements ISource_DC.VoltageMax
    Public Property CurrentMax As Single Implements ISource_DC.CurrentMax
    Public Property PowerMax As Single Implements ISource_DC.PowerMax

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Function IDN() As String Implements IDevice.IDN
        Return MyBase.IDN()
    End Function

    Public Overrides Sub RST() Implements IDevice.RST
        MyBase.RST()
    End Sub

    Public Overrides Sub CLS() Implements IDevice.CLS
        MyBase.CLS()
    End Sub

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

    Public Overridable Sub SetOutputON() Implements ISource_DC.SetOutputON
        Visa.SendString(":OUTPUT:STATE ON")
    End Sub

    Public Overridable Sub SetOutputOFF() Implements ISource_DC.SetOutputOFF
        Visa.SendString(":OUTPUT:STATE OFF")
    End Sub

    Public Overridable Sub SetVoltage(ByVal Voltage As Single, CurrentLim As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage

        If Voltage > _VoltageMax Then Voltage = _VoltageMax

        If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        If Voltage = 0 Then
            Visa.SendString("VOLTAGE " & Voltage)
            Visa.SendString(":OUTPUT:STATE OFF")
        Else
            Visa.SendString("SOURCE:CURRENT " & CurrentLim)
            Visa.SendString("SOURCE:VOLTAGE " & Voltage)
            If SetOutON Then Visa.SendString(":OUTPUT:STATE ON")
        End If

    End Sub

    Public Overridable Sub SetVoltage(Voltage As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage

        SetVoltage(Voltage, _CurrentMax, SetOutON)

    End Sub

    Public Overridable Sub SetVoltPuls(ByVal vPulse As Single, ByVal Width As Single, ByVal vEnd As Single) Implements ISource_DC.SetVoltPuls
        SetVoltage(vPulse, True)
        cHelper.Delay(Width)
        SetVoltage(vEnd, True)
    End Sub


    Public Overridable Sub SetCurrentLim(CurrentLim As Single) Implements ISource_DC.SetCurrentLim

        If CurrentLim > CurrentMax Then CurrentLim = _CurrentMax

        Visa.SendString("CURRENT " & CurrentLim)

    End Sub

#End Region

End Class
