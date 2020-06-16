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
    Implements ISource_DC

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize()
        Visa.SendString("*RST;*CLS")
        Visa.SendString("CURR 0")
        Visa.SendString("VOLT 0")
        Visa.SendString(":OUTP:STAT ON")
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

    Public Overrides Sub SetVoltPuls(ByVal vPulse As Single, ByVal Width As Single, ByVal vEnd As Single) Implements ISource_DC.SetVoltPuls
        MyBase.SetVoltPuls(vPulse, Width, vEnd)
    End Sub

    Public Overrides Sub SetCurrentLim(CurrentLim As Single) Implements ISource_DC.SetCurrentLim
        MyBase.SetCurrentLim(CurrentLim)
    End Sub

#End Region

End Class
