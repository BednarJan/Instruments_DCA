Imports Ivi.Visa
Imports System.Math

Public Class CTC_WEISS_WK64
    Inherits BCTC_WEISS

    Implements ITC
    Implements IDevice

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

        Name = "Weiss WK64"

        MinTemp = -75
        MaxTemp = 130

        MinHumidity = Single.MinValue
        MaxHumidity = Single.MaxValue

        MinPressure = Single.MinValue
        MaxPressure = Single.MaxValue

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Overrides Sub Initialize() Implements IDevice.Initialize

        SetpointTemp = 25
        SetpointHumidity = 0
        SetpointPressure = Single.MinValue

        HeatGrad = 3  ' K/min
        CoolGrad = 3  ' K/min
        HumiGrad = 2.5 ' K/min
        DeHumiGrad = 2.5 ' K/min

        SetGradients()

        TurnOFF()

    End Sub
#End Region

#Region "Interface ITC Methodes"



#End Region

#Region "private methodes and functions"

    Overrides Sub SetTempAndHumidity(ByVal temp As Single, Optional hum As Single = 0, Optional bONOff As Boolean = True)

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

        Visa.SendString(cmdStr)

    End Sub


    Overrides Function GetTempAndHumidity(ByRef hum As Single) As Single

        Dim Buffer As String
        Dim temp As Single = Single.MinValue

        Dim errMsg As String = vbNullString
        Dim sHlp() As String


        For iii As Integer = 1 To 5

            ' TC returns "-075.0 -038.3 0101000000000000Z"

            Visa.SendString("$00I")
            cHelper.Delay(1)
            Buffer = Visa.ReceiveString(errMsg)
            sHlp = Split(Buffer, " ")

            If IsNumeric(sHlp(1)) Then
                temp = Convert.ToSingle(sHlp(1))
                Exit For
            End If
        Next iii

        hum = Single.MinValue

        Return temp

    End Function


    Overrides Sub SetGradients()

    End Sub

#End Region


End Class





