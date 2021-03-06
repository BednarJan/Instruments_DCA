﻿Imports System.Runtime.Serialization
Imports Ivi.Visa
Imports Keysight.Visa

Public Class CVisaDevice
    Implements IVisaDevice

    Const byteArrSize As Integer = 3000000

#Region "Readonly Properties"
    Public ReadOnly Property ErrorLogger As CErrorLogger Implements IVisaDevice.ErrorLogger
    Public ReadOnly Property Session As IMessageBasedSession Implements IVisaDevice.Session
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Session = Session
        _Session.TimeoutMilliseconds = 5000
        _ErrorLogger = ErrorLogger
    End Sub
#End Region

#Region "IO Device Functions"
    Public Sub SendString(CMD As String, Optional ByRef ErrorMsg As String = "") Implements IVisaDevice.SendString

        If Not Session Is Nothing Then
            Try
                Session.FormattedIO.WriteLine(CMD)
            Catch ex As Exception
                ErrorMsg = ex.Message
                _ErrorLogger.LogException(ex, Session.ResourceName)
                Throw New CInstrumentException(ex, Session.ResourceName)
            End Try
        End If

    End Sub

    Public Function ReceiveString(Optional ByRef ErrorMsg As String = "", Optional termchar As Byte = 10) As String Implements IVisaDevice.ReceiveString

        If Not Session Is Nothing Then
            Try

                'original method Return Session.FormattedIO.ReadLine()

                Session.TerminationCharacterEnabled = True
                Session.TerminationCharacter = termchar
                Session.FormattedIO.DiscardBuffers()

                Return Session.FormattedIO.ReadUntilMatch(Chr(termchar))

            Catch ex As Exception
                ErrorMsg = ex.Message
                _ErrorLogger.LogException(ex, Session.ResourceName)
                Throw New CInstrumentException(ex, Session.ResourceName)
            End Try
        End If

        Return vbNullString

    End Function

    Public Function ReceiveValue(Optional ByRef ErrorMsg As String = "") As Double Implements IVisaDevice.ReceiveValue

        If Not Session Is Nothing Then
            Try
                Return Session.FormattedIO.ReadDouble()
            Catch ex As Exception

                ErrorMsg = ex.Message
                _ErrorLogger.LogException(ex, Session.ResourceName)
                Throw New CInstrumentException(ex, Session.ResourceName)

            End Try
        End If

        Return Double.NaN

    End Function

    'Public Function ReceiveValue(Optional ByRef ErrorMsg As String = "") As Double Implements IVisaDevice.ReceiveValue
    '    Dim str As String = String.Empty
    '    Dim retVal As Double = New Double

    '    If Session IsNot Nothing Then
    '        Try

    '            str = ReceiveString()
    '            retVal = cHelper.StringToDouble(str)

    '        Catch ex As Exception
    '            _ErrorLogger.LogException(ex, Session.ResourceName)
    '        End Try
    '    End If

    '    Return retVal

    'End Function


    'Public Function ReceiveValueList(Optional ByRef ErrorMsg As String = "") As Double() Implements IVisaDevice.ReceiveValueList
    '    If Not Session Is Nothing Then
    '        Try
    '            'Return Session.FormattedIO.ReadLineBinaryBlockOfDouble
    '            Return CType(Session.FormattedIO.ReadListOfDouble, Array)
    '        Catch ex As Exception
    '            _ErrorLogger.LogException(ex, Session.ResourceName)
    '            Return New Double() {}
    '        End Try
    '    End If

    'End Function


    Public Function ReceiveValueList(separ As Char, Optional ByRef ErrorMsg As String = "") As Double() Implements IVisaDevice.ReceiveValueList

        Dim str As String = String.Empty
        Dim retVal() As Double = New Double() {}

        If Session IsNot Nothing Then
            Try

                str = ReceiveString()
                retVal = cHelper.CSVString2DoubleArray(str, separ)

            Catch ex As Exception

                ErrorMsg = ex.Message
                _ErrorLogger.LogException(ex, Session.ResourceName)
                Throw New CInstrumentException(ex, Session.ResourceName)

            End Try

            Return retVal

        End If

    End Function

    Public Sub ReadTextToFile(ByVal HardcopyFullFileName As String, txtEnc As Text.Encoding, Optional fromFirstChar As Char = "", Optional termchar As Byte = 10) Implements IVisaDevice.ReadTextToFile

        Dim _data As Byte()

        Try
            _data = RecBytesRAW(termchar)

            If fromFirstChar <> vbNullChar Then

                Dim srcIndx As Long = Array.IndexOf(Of Byte)(_data, Asc(fromFirstChar))

                If srcIndx > 0 Then

                    Dim dstIndx As Long = 0
                    Dim dstLen As Integer = _data.Length - srcIndx
                    Dim dst(dstLen) As Byte

                    Array.Copy(_data, srcIndx, dst, dstIndx, dstLen)

                    System.IO.File.WriteAllText(HardcopyFullFileName, txtEnc.GetString(_data))

                End If

            Else

                System.IO.File.WriteAllText(HardcopyFullFileName, txtEnc.GetString(_data))

            End If

        Catch ex As Exception

            _ErrorLogger.LogException(ex, Session.ResourceName)
            Throw New CInstrumentException(ex, Session.ResourceName)

        End Try

    End Sub


    Public Sub ReadRAWDataToFile(ByVal HardcopyFullFileName As String, Optional fromFirstChar As Char = "", Optional termchar As Byte = 10) Implements IVisaDevice.ReadRAWDataToFile

        Dim _data As Byte()

        Try
            _data = RecBytesRAW(termchar)

            If fromFirstChar <> vbNullChar Then

                Dim srcIndx As Long = Array.IndexOf(Of Byte)(_data, Asc(fromFirstChar))

                If srcIndx > 0 Then

                    Dim dstIndx As Long = 0
                    Dim dstLen As Integer = _data.Length - srcIndx
                    Dim dst(dstLen) As Byte

                    Array.Copy(_data, srcIndx, dst, dstIndx, dstLen)

                    System.IO.File.WriteAllBytes(HardcopyFullFileName, _data)


                End If

            Else

                System.IO.File.WriteAllBytes(HardcopyFullFileName, _data)

            End If

        Catch ex As Exception

            _ErrorLogger.LogException(ex, Session.ResourceName)
            Throw New CInstrumentException(ex, Session.ResourceName)

        End Try

    End Sub


#End Region

#Region "Session Configuration via User Interface"
    Public Sub ShowConfigUI() Implements IVisaDevice.ShowConfigUI

        Dim _PropertyNames As List(Of String) = Nothing

        'Set up filter to select the 
        Select Case _Session.GetType
            Case GetType(GpibSession)
                _PropertyNames = New List(Of String)(New String() {"TimeoutMilliseconds", "AllowDma"})
            Case GetType(SerialSession)
                _PropertyNames = New List(Of String)(New String() {"BaudRate"})
            Case GetType(GpibSession)
                _PropertyNames = New List(Of String)(New String() {"one", "two", "three"})
            Case GetType(GpibSession)
                _PropertyNames = New List(Of String)(New String() {"one", "two", "three"})
            Case Else
                _PropertyNames = Nothing
        End Select

        'Dim frmConfigUI As New frmSessionConfig(Me.Session, _PropertyNames)

        'frmConfigUI.ShowDialog()

    End Sub
#End Region

#Region "Private Functions"

    Function RecBytesRAW(Optional termchar As Byte = 10) As Byte()

        Dim _data As Byte() = New Byte(byteArrSize) {}

        Try

            If Session IsNot Nothing Then

                Session.TerminationCharacterEnabled = False
                Session.TerminationCharacter = termchar

                Session.FormattedIO.DiscardBuffers()
                _data = Session.RawIO.Read(3000000)

            End If

        Catch ex As Exception

            _ErrorLogger.LogException(ex, Session.ResourceName)
            Throw ex

        End Try

        Return _data

    End Function


#End Region

End Class
