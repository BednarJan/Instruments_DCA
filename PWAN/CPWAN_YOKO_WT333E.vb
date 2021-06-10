'Class CPWAN_YOKO_WT333E
'08.06.2021, JaBe
'Compatible Instruments:
'- Yokogawa WT333E   ............ model with external current-voltage probe sensors
Imports Ivi.Visa

Public Class CPWAN_YOKO_WT333E
    Inherits BCPWAN_YOKO
    Implements IPWAN

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 600
        CurrentMax = 20
        InputElements = 3
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        MyBase.Initialize()

    End Sub

#End Region

#Region "Interface Methodes IPWAN"

#End Region

#Region "Public Special Functions "


#End Region

End Class
