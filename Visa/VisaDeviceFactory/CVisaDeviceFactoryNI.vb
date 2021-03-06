﻿Imports Ivi.Visa
Imports NationalInstruments.Visa

Public Class CVisaDeviceFactoryNI
    Implements IVisaDeviceFactory

    Public Function CreateDevice(ResourceName As String) As IMessageBasedSession Implements IVisaDeviceFactory.CreateDevice

        Dim _Session As Session

        If InStr(ResourceName, "INSTR") Then

            Select Case True
                Case InStr(ResourceName, "GPIB",)
                    _Session = New GpibSession(ResourceName)
                Case InStr(ResourceName, "ASRL")
                    _Session = New SerialSession(ResourceName)
                Case InStr(ResourceName, "USB")
                    _Session = New UsbSession(ResourceName)
                Case InStr(ResourceName, "TCPIP")
                    _Session = New TcpipSession(ResourceName)
                Case Else
                    _Session = Nothing
            End Select
        ElseIf InStr(ResourceName, "SOCKET") Then
            _Session = New TcpipSocket(ResourceName)
        Else
            _Session = Nothing
        End If

        Return _Session

    End Function


    Sub SetSerialPort(sesion As IMessageBasedSession, params As String) Implements IVisaDeviceFactory.SetSerialPort
        Dim serial As Ivi.Visa.ISerialSession = sesion

        Dim hlp() As String = params.Split(",")

        serial.BaudRate = hlp(0)

        Select Case hlp(1)
            Case 0
                serial.Parity = SerialParity.None
            Case 1
                serial.Parity = SerialParity.Odd
            Case 2
                serial.Parity = SerialParity.Even
        End Select

        Select Case hlp(2)
            Case 1
                serial.StopBits = SerialStopBitsMode.One
            Case 1.5
                serial.StopBits = SerialStopBitsMode.OneAndOneHalf
            Case 2
                serial.StopBits = SerialStopBitsMode.Two
        End Select

        serial.DataBits = hlp(3)

        Select Case hlp(4)
            Case "N"
                serial.FlowControl = SerialFlowControlModes.None
            Case "DTR/DSR"
                serial.FlowControl = SerialFlowControlModes.DtrDsr
            Case "RTS/CTS"
                serial.FlowControl = SerialFlowControlModes.RtsCts
            Case "XON/XOFF"
                serial.FlowControl = SerialFlowControlModes.XOnXOff
        End Select

    End Sub

End Class
