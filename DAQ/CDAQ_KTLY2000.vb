Imports Instruments
Imports Ivi.Visa

Public Class CDAQ_KTLY2000
    Implements IDevice

    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        _Visa = New CVisaDeviceNI(Session, ErrorLogger)

    End Sub

    Private Sub New()

    End Sub

#End Region

#Region "Basic Device Functions (IDevice)"
    Public Function IDN() As String Implements IDevice.IDN
        Dim ErrorMessages(1) As String

        _Visa.SendString("*IDN?", ErrorMessages(0))
        Return _Visa.ReceiveString(ErrorMessages(1))

    End Function

    Public Sub RST() Implements IDevice.RST
        Dim ErrorMessage As String = ""

        _Visa.SendString("*RST", ErrorMessage)

    End Sub

    Public Sub CLS() Implements IDevice.CLS
        Dim ErrorMessage As String = ""

        _Visa.SendString("*CLS", ErrorMessage)
    End Sub

    Public Sub Initialize() Implements IDevice.Initialize
        RST()
        CLS()
    End Sub
#End Region

    Public Function MeasDC() As Single

        _Visa.SendString(":INIT:CONT OFF")
        _Visa.SendString(":READ?")
        MeasDC = Visa.ReceiveValue()

    End Function

    'Public Function Get_Volt_AC(Chan As Integer) As Single Implements IDAQ.Get_Volt_AC
    '    Throw New NotImplementedException()
    'End Function

    'Public Function Get_Sample_Volt_DC(Chan As Integer, Sample As Integer) As Single Implements IDAQ.Get_Sample_Volt_DC
    '    Throw New NotImplementedException()
    'End Function

    'Public Function Get_Sample_Volt_Dyn(Chan As Integer, Sample As Integer) As Single Implements IDAQ.Get_Sample_Volt_Dyn
    '    Throw New NotImplementedException()
    'End Function

    'Public Function Get_Curr_DC(Chan As Integer) As Single Implements IDAQ.Get_Curr_DC
    '    Throw New NotImplementedException()
    'End Function

    'Public Function Get_Curr_AC(Chan As Integer) As Single Implements IDAQ.Get_Curr_AC
    '    Throw New NotImplementedException()
    'End Function

    'Public Function Get_Res(Chan As Integer) As Single Implements IDAQ.Get_Res
    '    Throw New NotImplementedException()
    'End Function

    'Public Function Get_FRes(Chan As Integer) As Single Implements IDAQ.Get_FRes
    '    Throw New NotImplementedException()
    'End Function

    'Public Function Get_FReq(Chan As Integer) As Single Implements IDAQ.Get_FReq
    '    Throw New NotImplementedException()
    'End Function
End Class
