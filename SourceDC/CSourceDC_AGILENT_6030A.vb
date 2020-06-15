'Class CSourceDC_Agilent_6030A
'07.04.2019, A. Zahler
'15.06.2020, Jabe
'Compatible Instruments:
'- Agilent/HP 6030A (single channel)
Imports Ivi.Visa

Public Class CSourceDC_AGILENT_6030A
    Inherits BCSourceDC
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

    Public Overrides Sub Initialize()
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
    End Sub

    Public Overrides Sub SetOutputOFF() Implements ISource_DC.SetOutputOFF
        MyBase.SetOutputOFF()
    End Sub

    Public Overrides Sub SetVoltage(ByVal Voltage As Single, CurrentLim As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage
        MyBase.SetVoltage(Voltage, CurrentLim, SetOutON)
    End Sub

    Public Overrides Sub SetVoltage(Voltage As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetVoltage
        MyBase.SetVoltage(Voltage, SetOutON)
    End Sub

    Public Overrides Sub SetCurrentLim(CurrentLim As Single) Implements ISource_DC.SetCurrentLim
        MyBase.SetCurrentLim(CurrentLim)
    End Sub

#End Region

End Class
