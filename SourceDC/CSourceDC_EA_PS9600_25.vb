'Class CSourceDC_EA_PS9600_25
'07.04.2019, A. Zahler
'12.06.2020, JaBe
'Compatible Instruments:
'- EA-PS 9600-25 (single channel)
Imports Ivi.Visa

Public Class CSourceDC_EA_PS9600_25
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
        MyBase.Initialize()
    End Sub
#End Region

#Region "Interface Methodes ISource_DC"


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
