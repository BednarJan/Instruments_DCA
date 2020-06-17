'Class CPWAN_YOKO_WT310
'06.01.2019, A. Zahler
'17.06.2020, JaBe
'Compatible Instruments:
'- Yokogawa WT310HC
'- Yokogawa WT310EH
Imports Ivi.Visa

Public Class CPWAN_YOKO_WT310
    Inherits BCPWAN
    Implements IPWAN

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 600
        CurrentMax = 40
        InputElements = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize


    End Sub

#End Region

#Region "Interface Methodes IPWAN"

#End Region

#Region "Public Special Functions WT310"

    Private Sub AutoRangeDelay()
        Dim TimerEnd As Double, TimerCounter As Double
        Dim SumCheckRange As Integer

        'Check during one second if Auto Range is "busy"
TimerStart:
        SumCheckRange = 0
        TimerEnd = Timer + 1
        Do While TimerCounter < TimerEnd
            SumCheckRange = SumCheckRange + CheckRange()
            TimerCounter = Timer
        Loop

        If SumCheckRange > 0 Then GoTo TimerStart
    End Sub

    Private Function CheckRange() As Integer
        Visa.SendString(":INPUT:CRANGE?" & Chr(10))
        Return Visa.ReceiveValue

    End Function
#End Region

End Class
