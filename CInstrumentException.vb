Public Class CInstrumentException
    Inherits System.Exception
#Region "Shorthand Properties"

    Public Property VisaAdr As String
    Public Property Exception As System.Exception

#End Region

    Public Sub New(myEx As Exception, myVisaAdr As String)
        MyBase.New(myEx.Message)
        _Exception = myEx
        _VisaAdr = myVisaAdr
    End Sub

End Class
