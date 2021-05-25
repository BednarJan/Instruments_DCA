'CDAQ_AGIL_34970A
'21.05.2021, J. Bednar 
'Compatible Instruments:
'- AGIL_34980A
Imports Ivi.Visa
Public Class CDAQ_AGIL_34970A
    Inherits BCDAQ
    Implements IDAQ

#Region "Constructor"
    Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)
        MyBase.ModuleFacor = 100

    End Sub

#End Region

#Region "Interface Methodes IDAQ"
    Public Overrides Function QueryNumericItems() As Double() Implements IDAQ.QueryNumericItems

        Return MyBase.QueryNumericItems()

    End Function

    Public Overrides Function MeasChannel(Chan As CDAQChannel, Optional Sample As Integer = 1) As Single Implements IDAQ.MeasChannel

        ConfigureChannel(Chan)
        RouteChannel(Chan)
        Visa.SendString("READ?")
        Return Visa.ReceiveValue()

    End Function



#End Region
End Class
