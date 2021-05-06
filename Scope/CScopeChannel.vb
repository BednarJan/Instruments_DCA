Public Class CScopeChannel

    Private _ErrorLogger As CErrorLogger

#Region "Enum Definnitions"
    Public Enum ChanBandwidth
        B20 = 20
        B100 = 100
    End Enum

    Public Enum ChanCoupling
        DC = 1
        AC = 2
        LFReject = 3
        DC_HFReject = 4
        DC_NREJect = 5
        AC_HFReject = 6
        AC_NREJect = 7
    End Enum

    Public Enum ChanPolarity
        NORMAL = 1
        INVERT = 2
    End Enum

    Public Enum ChanState
        STATE_OFF = 0
        STATE_ON = 1
    End Enum

    Public Enum Acquire
        STATE_STOP = 0
        STATE_SINGL = 1
        STATE_RUN = 2
    End Enum

    Public Enum InputImpedance
        INPUT_IMPEDANCE_MEG = 0
        INPUT_IMPEDANCE_FIFTY = 1
    End Enum

#End Region



#Region "Shorthand Properties"

    Public Property ChanNr As Integer
    Public Property State As ChanState
    Public Property VertVolt As Single
    Public Property DispUnit As String
    Public Property Position As Single
    Public Property Polarity As ChanPolarity
    Public Property Probe As Single
    Public Property Bandwidth As ChanBandwidth
    Public Property Coupling As ChanCoupling
    Public Property Impedance As Integer
    Public Property Offset As Single

    Public ReadOnly Property Name As String
        Get
            Return "CH" & _ChanNr.ToString
        End Get
    End Property


#End Region

#Region "Constructor"
    Public Sub New(ErrorLogger As CErrorLogger, _Nr As Integer)

        _ErrorLogger = ErrorLogger
        _ChanNr = _Nr

    End Sub
#End Region

    Public Sub InitChannel(_Nr As Integer)

        _ChanNr = _Nr

        _State = ChanState.STATE_ON
        _VertVolt = 1000
        _DispUnit = "V"
        _Position = _Nr - 1
        _Polarity = ChanPolarity.NORMAL
        _Probe = 1
        _Bandwidth = ChanBandwidth.B20
        _Coupling = ChanCoupling.AC
        _Impedance = InputImpedance.INPUT_IMPEDANCE_MEG
        _Offset = 0

    End Sub

End Class
