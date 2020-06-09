Imports Ivi.Visa

Public Class CVisaManager

    Public Resources As List(Of String)

    Public Sub New() '(GPIB As CGPIBDevice, ErrorLogger As CErrorLogger)

        Try
            Resources = GlobalResourceManager.Find()
        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try

    End Sub

End Class
