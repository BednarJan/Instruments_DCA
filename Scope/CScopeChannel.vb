Public Class CScopeChannel

    Private _ErrorLogger As CErrorLogger

#Region "Enum Definnitions"
    Enum ChanBandwidth
        B20 = 20
        B100 = 100
    End Enum

    Enum ChanCoupling
        DC = 1
        AC = 2
        LFReject = 3
        DC_HFReject = 4
        DC_NREJect = 5
        AC_HFReject = 6
        AC_NREJect = 7
    End Enum

    Enum ChanPolarity
        NORMAL = 1
        INVERT = 2
    End Enum

    Enum ChanState
        STATE_OFF = 0
        STATE_ON = 1
    End Enum


    Enum Acquire
        STATE_STOP = 0
        STATE_SINGL = 1
        STATE_RUN = 2
    End Enum

#End Region



#Region "Shorthand Properties"
    Public Property Name As String
    Public Property State As ChanState
    Public Property VertVolt As Single
    Public Property DispUnit As String
    Public Property Position As Single
    Public Property Polarity As ChanPolarity
    Public Property Probe As Single
    Public Property Bandwidth As ChanBandwidth
    Public Property Coupling As ChanCoupling
    Public Property InvertOff As Integer
    Public Property Impedance As Integer
    Public Property Offset As Single

#End Region

#Region "Constructor"
    Public Sub New(ErrorLogger As CErrorLogger)
        _ErrorLogger = ErrorLogger
    End Sub
#End Region



End Class
