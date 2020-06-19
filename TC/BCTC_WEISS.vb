Imports Ivi.Visa
Imports System.Math

Public Class BCTC_WEISS
    Inherits BCDevice
    Implements ITC

    Public Const STX As Byte = 2
    Public Const ETX As Byte = 3

#Region "Shorthand Properties ITC"

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
        MyBase.New(Session, ErrorLogger)
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"
    Public Overrides Function IDN() As String Implements IDevice.IDN
        Return MyBase.Name
    End Function

    Public Overrides Sub RST() Implements IDevice.RST
        TurnOFF()
    End Sub

    Public Overrides Sub CLS() Implements IDevice.CLS
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        SetGradients()
        TurnOFF()

    End Sub
#End Region


#Region "Interface ITC Methodes"
    Public Overridable Sub SendCommand(cmdStr As String) Implements ITC.SendCommand
        Visa.SendString(cmdStr)
    End Sub

    Public Overridable Function ReadResponse() As String Implements ITC.ReadResponse
        Return Visa.ReceiveString
    End Function


    Public Overridable Sub SetTemp(val As Single) Implements ITC.SetTemp

        _SetpointTemp = val

        Call SetTempAndHumidity(_SetpointTemp,, True)

    End Sub

    Public Overridable Function GetSetpointTemp() As Single Implements ITC.GetSetpointTemp
        Return _SetpointTemp
    End Function

    Public Overridable Sub SetHumidity(val As Single) Implements ITC.SetHumidity

        _SetpointHumidity = val
        Call SetTempAndHumidity(_SetpointTemp, val, True)

    End Sub

    Public Overridable Function GetHumidity() As Single Implements ITC.GetHumidity
        Return _SetpointHumidity
    End Function

    Public Overridable Sub SetPumpPressure(val As Single) Implements ITC.SetPumpPressure
        Throw New NotImplementedException()
    End Sub

    Public Overridable Function GetPumpPressure() As Single Implements ITC.GetPumpPressure
        Throw New NotImplementedException()
    End Function

    Public Overridable Sub TurnOFF() Implements ITC.TurnOFF
        SetTempAndHumidity(_SetpointTemp, _SetpointHumidity, False)
    End Sub

    Public Overridable Function GetInternalTemp() As Single Implements ITC.GetInternalTemp
        Dim actual_humidity As Single
        Return GetTempAndHumidity(actual_humidity)
    End Function

    Public Overridable Function GetProcessTemp() As Single Implements ITC.GetProcessTemp
        Dim actual_humidity As Single
        Return GetTempAndHumidity(actual_humidity)
    End Function

    Public Overridable Function RegTemp(ByVal mySetVal As Single, Optional ByVal accu As Single = 1) As Boolean Implements ITC.RegTemp

        Dim EndTime As DateTime
        Dim actTemp As Single
        Dim NotFinished As Boolean

        NotFinished = True

        Call SetTemp(mySetVal)
        EndTime = Now().AddHours(1)
        cHelper.Delay(3)           ' wait 3 sec

        Do While NotFinished        ' condition for waiting till Chamber will sets a temp
            If Now() < EndTime Then ' timeout condition, if it should take more than 1 hour  

                actTemp = GetProcessTemp()

                If (actTemp < mySetVal + accu) And (actTemp > mySetVal - accu) Then   ' acuarecy +/- C
                    NotFinished = False
                Else
                    cHelper.Delay(5)
                End If
            Else
                NotFinished = True
            End If
        Loop

        Return Not NotFinished

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





