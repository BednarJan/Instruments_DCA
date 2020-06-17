'Class BCDevice 
'02.05.2019, J. Bednar  
'12.06.2020, JaBe
'Base clase for the all devices (instruments)
Imports Ivi.Visa

Public MustInherit Class BCDevice
    Implements IDevice

#Region "Shorthand Properties"

    Friend _ErrorLogger As CErrorLogger
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        Try
            _Visa = New CVisaDevice(Session, ErrorLogger)
            _ErrorLogger = ErrorLogger
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"
    Public Overridable Function IDN() As String Implements IDevice.IDN

        Dim ErrorMessages(1) As String

        _Visa.SendString("*IDN?", ErrorMessages(0))
        Return _Visa.ReceiveString(ErrorMessages(1))

    End Function

    Public Overridable Sub RST() Implements IDevice.RST

        Dim ErrorMessage As String = ""

        _Visa.SendString("*RST", ErrorMessage)

    End Sub

    Public Overridable Sub CLS() Implements IDevice.CLS

        Dim ErrorMessage As String = ""

        _Visa.SendString("*CLS", ErrorMessage)
    End Sub

    Public Overridable Sub Initialize() Implements IDevice.Initialize

        Dim ErrorMessage As String = ""

        _Visa.SendString("*RST", ErrorMessage)

    End Sub

    Public Sub SendString(ByVal myStr As String) Implements IDevice.SendString

        Dim ErrorMessage As String = ""

        _Visa.SendString(myStr, ErrorMessage)
    End Sub

    Public Function ReceiveString() As String Implements IDevice.ReceiveString
        Dim ErrorMessage As String = ""
        Return _Visa.ReceiveString(ErrorMessage)
    End Function


#End Region


End Class

