Imports Ivi.Visa
Imports System.Math

Public Class CTC_WEISS_WK64
    Implements ITC
    Implements IDevice

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"

    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa

    Public ReadOnly Property SetpointTemp As Single Implements ITC.SetpointTemp
    Public ReadOnly Property SetpointHumidity As Single Implements ITC.SetpointHumidity
    Public ReadOnly Property SetpointPressure As Single Implements ITC.SetpointPressure


    Public ReadOnly Property MinTemp As Single Implements ITC.MinTemp
    Public ReadOnly Property MaxTemp As Single Implements ITC.MaxTemp

    Public ReadOnly Property MinHumidity As Single Implements ITC.MinHumidity
    Public ReadOnly Property MaxHumidity As Single Implements ITC.MaxHumidity

    Public ReadOnly Property MinPressure As Single Implements ITC.MinPressure
    Public ReadOnly Property MaxPressure As Single Implements ITC.MaxPressure


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Visa = New CVisaDeviceNI(Session, ErrorLogger)
        _ErrorLogger = ErrorLogger
        _Name = "Weiss WK64"
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"
    Public Function IDN() As String Implements IDevice.IDN

        Return _Name

    End Function

    Public Sub RST() Implements IDevice.RST
        Throw New NotImplementedException()
    End Sub

    Public Sub CLS() Implements IDevice.CLS
        Throw New NotImplementedException()
    End Sub

    Public Sub Initialize() Implements IDevice.Initialize

        _SetpointTemp = 25
        _SetpointHumidity = 0
        _SetpointPressure = 1.5
        _MinTemp = -75
        _MaxTemp = 130

        TurnOFF()

    End Sub
#End Region

#Region "Interface ITC Methodes"
    Public Sub SendCommand(cmdStr As String) Implements ITC.SendCommand
        _Visa.SendString(cmdStr)
    End Sub

    Public Function ReadResponse() As String Implements ITC.ReadResponse
        Return _Visa.ReceiveString
    End Function


    Public Sub SetTemp(val As Single) Implements ITC.SetTemp

        _SetpointTemp = val

        Call SetTempHumidity(_SetpointTemp,, False)
    End Sub

    Public Function GetSetpointTemp() As Single Implements ITC.GetSetpointTemp
        Return _SetpointTemp
    End Function

    Public Sub SetHumidity(val As Single) Implements ITC.SetHumidity

        Throw New NotImplementedException()
    End Sub

    Public Sub SetPumpPressure(val As Single) Implements ITC.SetPumpPressure
        Throw New NotImplementedException()
    End Sub

    Public Sub TurnOFF() Implements ITC.TurnOFF
        SetTempHumidity(_SetpointTemp, , False)
    End Sub

    Public Sub TurnON() Implements ITC.TurnON
        SetTempHumidity(_SetpointTemp, , True)
    End Sub

    Public Function GetInternalTemp() As Single Implements ITC.GetInternalTemp
        Dim actual_humidity As Single
        Return GetTempAndHumidity(actual_humidity)
    End Function

    Public Function GetProcessTemp() As Single Implements ITC.GetProcessTemp
        Dim actual_humidity As Single
        Return GetTempAndHumidity(actual_humidity)
    End Function

    Public Function RegTemp(ByVal val As Single) As Boolean Implements ITC.RegTemp
        Throw New NotImplementedException()
    End Function

    Public Function GetHumidity() As Single Implements ITC.GetHumidity
        'Throw New NotImplementedException()
        Return Single.MinValue
    End Function

    Public Function GetPumpPressure() As Single Implements ITC.GetPumpPressure
        'Throw New NotImplementedException()
        Return Single.MinValue
    End Function

#End Region

#Region "private methodes and functions"

    Private Sub SetTempHumidity(ByVal temp As Single, Optional hum As Single = 0, Optional bONOff As Boolean = True)

        Dim strg As String, cmdStr As String
        Dim sONoFF As String

        If temp < MinTemp Then temp = MinTemp
        If temp > MaxTemp Then temp = MaxTemp

        sONoFF = "0100001000000"
        If bONOff = True Then
            sONoFF = "0101010001000000"
        End If

        If temp < 0 Then
            strg = "00E" & Format(Abs(temp), "0-00.0") & sONoFF & "000"
        Else
            strg = "00E" & Format(temp, "0000.0") & sONoFF & "000"
        End If

        cmdStr = Chr(36) & strg & Chr(13)

        Call _Visa.SendString(cmdStr)

    End Sub


    Private Function GetTempAndHumidity(ByRef hum As Single) As Single

        Dim Buffer As String
        Dim temp As Single = Single.MinValue

        Dim errMsg As String = vbNullString
        Dim sHlp() As String


        For iii As Integer = 1 To 5

            ' TC returns "-075.0 -038.3 0101000000000000Z"

            _Visa.SendString("$00I")
            cHelper.Delay(1)
            Buffer = _Visa.ReceiveString(errMsg)
            sHlp = Split(Buffer, " ")

            If IsNumeric(sHlp(1)) Then
                temp = Convert.ToSingle(sHlp(1))
                Exit For
            End If
        Next iii

        hum = Single.MinValue

        Return temp

    End Function


#End Region




End Class





