Imports Ivi.Visa
Imports System.Math

Public Class CTC_WEISS_WT3340

    Inherits BCTC_WEISS
    Implements ITC
    Implements IDevice

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)
        Name = "Weiss WT3340"


    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Overrides Sub Initialize() Implements IDevice.Initialize

        SetpointTemp = 25
        SetpointHumidity = 0
        SetpointPressure = Single.MinValue
        MinTemp = -75
        MaxTemp = 130

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

        MyBase.SendCommand(cmdStr)
    End Sub


    Overrides Function GetTempAndHumidity(ByRef hum As Single) As Single

        Dim cmdStr As String, Buffer As String, bRet As Boolean, Checksum As String
        Dim temp As Single
        Dim tempIndexStart As Integer
        Dim tempIndexEnd As Integer
        Dim humiIndexEnd As Integer
        Dim errMsg As String = vbNullString

        cmdStr = "1?"
        Checksum = Calc_Checksum(cmdStr)

        cmdStr = Chr(2) & cmdStr & Checksum

        MyBase.SendCommand("$00I")
        cHelper.Delay(1)
        Buffer = MyBase.ReadResponse(errMsg)

        tempIndexStart = InStr(Buffer, "T")
        tempIndexEnd = InStr(Buffer, "F")
        humiIndexEnd = InStr(Buffer, "P")

        bRet = IsNumeric(Mid(Buffer, tempIndexStart + 1, tempIndexEnd - tempIndexStart - 1))

        temp = Single.MinValue
        hum = Single.MinValue
        If bRet Then
            temp = Convert.ToSingle(Mid(Buffer, tempIndexStart + 1, tempIndexEnd - tempIndexStart - 1))
            hum = Convert.ToSingle(Mid(Buffer, tempIndexEnd + 1, humiIndexEnd - tempIndexEnd - 1))
        End If

        Return temp

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

        MyBase.SendCommand(cmdStr)

    End Sub

#End Region


End Class





