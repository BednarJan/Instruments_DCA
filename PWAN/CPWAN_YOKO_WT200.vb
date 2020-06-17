'Class CPWAN_YOKO_WT200
'17.06.2020, JaBe
'Compatible Instruments:
Imports Ivi.Visa

Public Class CPWAN_YOKO_WT200
    Inherits BCPWAN
    Implements IPWAN

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 600
        CurrentMax = 20
        InputElements = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize


    End Sub

#End Region

#Region "Interface Methodes IPWAN"

#End Region

#Region "Public Special Functions"


#End Region

End Class
