Public Class CScopeChannels
    Inherits List(Of CScopeChannel)

    Private _Channel As CScopeChannel

    Sub New()

        MyBase.New

    End Sub

    Public Property Channel(ChanNr As Integer)
        Get
            Dim _CScopeChannel As CScopeChannel = Nothing

            For Each _CScopeChannel In Me

                If _CScopeChannel.Name.LastIndexOf(CStr(ChanNr)) > 0 Then

                    Exit For

                End If

            Next

            Return _CScopeChannel

        End Get
        Set(value)
            Me(ChanNr) = value
        End Set
    End Property

End Class
