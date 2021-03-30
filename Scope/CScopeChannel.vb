Public Class CScopeChannel

    Private _ErrorLogger As CErrorLogger

#Region "Shorthand Properties"
    Public Property Nr As Integer
    Public Property Name As String
    Public Property IsOn As Boolean
    Public Property VertVolt As Single
    Public Property DispUnit As String
    Public Property Position As Single
    Public Property Probe As Single
    Public Property Bandwidth As Integer
    Public Property Coupling As Integer
    Public Property InvertOff As Integer
    Public Property Impedance As Integer
    Public Property Offset As Single

#Region "Constructor"
    Public Sub New(ErrorLogger As CErrorLogger)
        _ErrorLogger = ErrorLogger
    End Sub
#End Region


#End Region

End Class
