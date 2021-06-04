Public Class CNumericNormalItem

    Property Fn As String
    Property Elm As IPWAN.Elements
    Property Itm As Integer
    Property OrderHarm As String


    Sub New(myFn As String, myItm As Integer, myElm As IPWAN.Elements)

        _Fn = myFn
        _Elm = myElm
        _Itm = myItm
        _OrderHarm = String.Empty

    End Sub

    Sub New(myFn As String, myItm As Integer, myElm As IPWAN.Elements, myOrderHarm As String)

        _Fn = myFn
        _Elm = myElm
        _Itm = myItm
        _OrderHarm = myOrderHarm

    End Sub

End Class
