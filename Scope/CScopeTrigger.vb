Public Class CScopeTrigger
    Private _ErrorLogger As CErrorLogger

#Region "Enum definitions"
    Enum TriggerSlope
        RISE = 1
        FALL = 2
        EITHER = 3
    End Enum

    Enum TriggerSource
        Ch1 = 1
        Ch2 = 2
        Ch3 = 3
        Ch4 = 4
        Ch5 = 5
        Ch6 = 6
        Ch7 = 7
        Ch8 = 8
        Ext = 99
    End Enum


    Enum TriggerMode
        NORMAL = 1
        AUTO = 2
        SIGNL = 3
    End Enum

    Enum TriggerCoupling
        DC = 1
        AC = 2
        LFReject = 3
        DC_HFReject = 4
        DC_NREJect = 5
        AC_HFReject = 6
        AC_NREJect = 7
    End Enum


    Enum Acquire
        STATE_STOP = 0
        STATE_SINGLE = 1
        STATE_RUN = 2
    End Enum

#End Region

#Region "Shorthand Properties"

    Property Mode As TriggerMode
    Property Coupling As TriggerCoupling
    Property Position As Integer
    Property Status As Integer
    Property Level As Single
    Property Source As TriggerSource
    Property Slope As TriggerSlope

#End Region

#Region "Constructor"
    Public Sub New(ErrorLogger As CErrorLogger)
        _ErrorLogger = ErrorLogger


    End Sub
#End Region
End Class
