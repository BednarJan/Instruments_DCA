Public Class CDAQChannel
#Region "Shorthand Properties"

    Property Modul As Integer
    Property Nr As Integer
    Property DAQFn As IDAQ.DAQ_Function
    Property Range As Decimal

#End Region

    Sub New(nModul As Integer, nChan As Integer, nDAQFn As IDAQ.DAQ_Function)
        _Modul = nModul
        _Nr = nChan
        _DAQFn = nDAQFn
        _Range = 0
    End Sub


    Sub New(nModul As Integer, nChan As Integer, nDAQFn As IDAQ.DAQ_Function, nRange As Decimal)
        _Modul = nModul
        _Nr = nChan
        _DAQFn = nDAQFn
        _Range = nRange
    End Sub


End Class
