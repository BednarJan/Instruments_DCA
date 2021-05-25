Public Interface IDAQ
    Inherits IDevice

#Region "Properties (Set & Get)"

    Property VoltageMax As Single
    Property CurrentMax As Single

#End Region


#Region "Enum definitions"

    Enum DAQ_Function
        VOLT_DC = 1
        VOLT_AC = 2
        CURRENT_DC = 3
        CURRENT_AC = 4
        RES = 5
        FRES = 6
        FREQ = 7
    End Enum

#End Region

#Region "Methods (Sub & Functions)"

    Property ScanList As List(Of CDAQChannel)

    Function MeasChannel(Chan As CDAQChannel, Optional Sample As Integer = 1) As Single

    Function QueryNumericItems() As Double()

    Sub RouteClose(ByVal nModul As Integer, ByVal Chan As Integer)
    Sub RouteOpen(ByVal nModul As Integer, ByVal Chan As Integer)

#End Region

End Interface
