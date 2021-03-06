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

End Class
