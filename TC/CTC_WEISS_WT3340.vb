Imports Ivi.Visa
Imports System.Math

Public Class CTC_WEISS_WT3340

    Inherits BCTC_WEISS
    Implements ITC

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

        Name = "Weiss WT3340"
        MinTemp = -75
        MaxTemp = 130

        MinHumidity = 0
        MaxHumidity = 100

        MinPressure = Single.MinValue
        MaxPressure = Single.MaxValue

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Overrides Sub Initialize() Implements IDevice.Initialize

        SetpointTemp = 25
        SetpointHumidity = 0
        SetpointPressure = Single.MinValue

        HeatGrad = 3  ' K/min
        CoolGrad = 3  ' K/min
        HumiGrad = 2.5 ' K/min
        DeHumiGrad = 2.5 ' K/min

        SetGradients()

        TurnOFF()

    End Sub
#End Region

#Region "Interface ITC Methodes"



#End Region

#Region "private methodes and functions"

    Overrides Sub SetTempAndHumidity(ByVal temp As Single, Optional hum As Single = 0, Optional bONOff As Boolean = True)

        MyBase.SetTempAndHumidity(temp, hum, bONOff)

    End Sub


    Overrides Function GetTempAndHumidity(ByRef hum As Single) As Single

        Return MyBase.GetTempAndHumidity(hum)

    End Function


    Overrides Sub SetGradients()

        MyBase.SetGradients()

    End Sub

#End Region


End Class





