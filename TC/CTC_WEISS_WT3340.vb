Imports Ivi.Visa
Imports System.Math

Public Class CTC_WEISS_WT3340

    Inherits BCTC_WEISS
    Implements ITC

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

        Name = "Weiss WT3340"
        MinTemp = -75
        MaxTemp = 130

        MinHumidity = 0
        MaxHumidity = 100

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

        Dim cmdStr As String
        Dim Checksum As String
        Dim sONoFF As String

        If temp < -75 Then temp = -75
        If temp > 130 Then temp = 130

        sONoFF = "0"
        If bONOff = True Then
            sONoFF = "1"
        End If

        If temp < 0 Then
            cmdStr = "1T" & Format(Abs(temp), "-00.0") & "F" & Format(hum, "00") & "R" & sONoFF & "000000000000000"
        Else
            cmdStr = "1T" & Format(temp, "000.0") & "F" & Format(hum, "000") & "R" & sONoFF & "000000000000000"
        End If

        Checksum = Calc_Checksum(cmdStr)

        cmdStr = Chr(STX) & cmdStr & Checksum & Chr(ETX)

        Visa.SendString(cmdStr)
    End Sub


    Overrides Function GetTempAndHumidity(ByRef hum As Single) As Single

        Dim cmdStr As String, Buffer As String, Checksum As String
        Dim actTemp As Single
        Dim tempIndexStart As Integer
        Dim humiIndexStart As Integer
        Dim humiIndexEnd As Integer
        Dim errMsg As String = vbNullString



        cmdStr = "1?"
        Checksum = Calc_Checksum(cmdStr)

        cmdStr = Chr(STX) & cmdStr & Checksum & Chr(ETX)

        Visa.SendString(cmdStr)

        cHelper.Delay(0.5)

        Buffer = Visa.ReceiveString(errMsg, ETX)

        actTemp = Single.MinValue
        hum = Single.MinValue

        If Not Buffer Is Nothing Then
            tempIndexStart = InStr(Buffer, "T")
            humiIndexStart = InStr(Buffer, "F")
            humiIndexEnd = InStr(Buffer, "P")

            Dim strTemp As String = Mid(Buffer, tempIndexStart + 1, humiIndexStart - tempIndexStart - 1)
            Dim strHum As String = Mid(Buffer, humiIndexStart + 1, humiIndexEnd - humiIndexStart - 1)

            If IsNumeric(strTemp) AndAlso IsNumeric(strHum) Then
                actTemp = Convert.ToSingle(strTemp)
                hum = Convert.ToSingle(strHum)
            End If
        End If

        Return actTemp

    End Function


    Overrides Sub SetGradients()

        Dim cmdStr As String
        Dim Checksum As String = vbNullString

        cmdStr = "1U" & Format(HeatGrad, "0000.0") & " "
        cmdStr &= Format(CoolGrad, "0000.0") & " "
        cmdStr &= Format(HumiGrad, "0000.0") & " "
        cmdStr &= Format(DeHumiGrad, "0000.0")

        Checksum = Calc_Checksum(cmdStr)

        cmdStr = Chr(STX) & cmdStr & Checksum & Chr(ETX)

        Visa.SendString(cmdStr)

    End Sub

#End Region


End Class





