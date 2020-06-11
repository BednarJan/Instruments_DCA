Imports Ivi.Visa
Imports System.Math

Public Class BCTC_WEISS
    Implements ITC
    Implements IDevice

    Public Const STX As Byte = 2
    Public Const ETX As Byte = 3

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"

    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa

    Public Property SetpointTemp As Single Implements ITC.SetpointTemp
    Public Property SetpointHumidity As Single Implements ITC.SetpointHumidity
    Public Property SetpointPressure As Single Implements ITC.SetpointPressure


    Public Property MinTemp As Single Implements ITC.MinTemp
    Public Property MaxTemp As Single Implements ITC.MaxTemp

    Public Property MinHumidity As Single Implements ITC.MinHumidity
    Public Property MaxHumidity As Single Implements ITC.MaxHumidity

    Public Property MinPressure As Single Implements ITC.MinPressure
    Public Property MaxPressure As Single Implements ITC.MaxPressure

    Public Property HeatGrad As Single Implements ITC.HeatGrad
    Public Property CoolGrad As Single Implements ITC.CoolGrad

    Public Property HumiGrad As Single Implements ITC.HumiGrad

    Public Property DeHumiGrad As Single Implements ITC.DeHumiGrad


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Visa = New CVisaDevice(Session, ErrorLogger)
        _ErrorLogger = ErrorLogger
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"
    Public Overridable Function IDN() As String Implements IDevice.IDN

        Return _Name

    End Function

    Public Overridable Sub RST() Implements IDevice.RST
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub CLS() Implements IDevice.CLS
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub Initialize() Implements IDevice.Initialize

        SetGradients()
        TurnOFF()

    End Sub
#End Region

#Region "Interface ITC Methodes"
    Public Overridable Sub SendCommand(cmdStr As String) Implements ITC.SendCommand
        _Visa.SendString(cmdStr)
    End Sub

    Public Overridable Function ReadResponse() As String Implements ITC.ReadResponse
        Return _Visa.ReceiveString
    End Function


    Public Overridable Sub SetTemp(val As Single) Implements ITC.SetTemp

        _SetpointTemp = val

        Call SetTempAndHumidity(_SetpointTemp,, True)
    End Sub

    Public Overridable Function GetSetpointTemp() As Single Implements ITC.GetSetpointTemp
        Return _SetpointTemp
    End Function

    Public Overridable Sub SetHumidity(val As Single) Implements ITC.SetHumidity

        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub SetPumpPressure(val As Single) Implements ITC.SetPumpPressure
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub TurnOFF() Implements ITC.TurnOFF
        SetTempAndHumidity(_SetpointTemp, , False)
    End Sub

    Public Overridable Sub TurnON() Implements ITC.TurnON
        SetTempAndHumidity(_SetpointTemp, , True)
    End Sub

    Public Overridable Function GetInternalTemp() As Single Implements ITC.GetInternalTemp
        Dim actual_humidity As Single
        Return GetTempAndHumidity(actual_humidity)
    End Function

    Public Overridable Function GetProcessTemp() As Single Implements ITC.GetProcessTemp
        Dim actual_humidity As Single
        Return GetTempAndHumidity(actual_humidity)
    End Function

    Public Overridable Function RegTemp(ByVal val As Single) As Boolean Implements ITC.RegTemp
        Throw New NotImplementedException()
    End Function

    Public Overridable Function GetHumidity() As Single Implements ITC.GetHumidity
        'Throw New NotImplementedException()
        Return Single.MinValue
    End Function

    Public Overridable Function GetPumpPressure() As Single Implements ITC.GetPumpPressure
        'Throw New NotImplementedException()
        Return Single.MinValue
    End Function

    Public Overridable Sub SetGradients() Implements ITC.SetGradients

    End Sub


#End Region

#Region "private methodes and functions"

    Overridable Sub SetTempAndHumidity(ByVal temp As Single, Optional hum As Single = 0, Optional bONOff As Boolean = True)

    End Sub


    Overridable Function GetTempAndHumidity(ByRef hum As Single) As Single

        Return Single.MinValue

    End Function


    Public Function Calc_Checksum(ByVal myStr)
        Dim l As Integer, i As Integer
        Dim b As Integer, j As Integer, K As Byte
        Dim ch As String

        myStr = Chr(2) + myStr
        l = Len(myStr)
        b = 256

        For i = 1 To l
            ch = Mid(myStr, i, 1)
            j = Asc(ch)
            If j < b Then
                b += -j
            Else
                b += -j
                b += 256
            End If
        Next i

        j = b \ 16
        If j < 10 Then
            j += 48
        Else
            j += 55
        End If

        K = b Mod 16
        If K < 10 Then
            K += 48
        Else
            K += 55
        End If

        Return Chr(j) + Chr(K)

    End Function


#End Region




End Class





