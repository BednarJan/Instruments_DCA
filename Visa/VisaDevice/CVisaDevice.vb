Imports System.Runtime.Serialization
Imports Ivi.Visa
Imports Keysight.Visa

Public Class CVisaDevice
    Implements IVisaDevice

#Region "Readonly Properties"
    Public ReadOnly Property ErrorLogger As CErrorLogger Implements IVisaDevice.ErrorLogger
    Public ReadOnly Property Session As IMessageBasedSession Implements IVisaDevice.Session
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Session = Session
        _ErrorLogger = ErrorLogger
    End Sub
#End Region

#Region "IO Device Functions"
    Public Sub SendString(CMD As String, Optional ByRef ErrorMsg As String = "") Implements IVisaDevice.SendString

        If Not Session Is Nothing Then
            Try
                Session.FormattedIO.WriteLine(CMD)
            Catch ex As Exception
                _ErrorLogger.LogException(ex, Session.ResourceName)
            End Try
        End If

    End Sub

    Public Function ReceiveString(Optional ByRef ErrorMsg As String = "", Optional termchar As Byte = 10) As String Implements IVisaDevice.ReceiveString

        If Not Session Is Nothing Then
            Try

                'original method Return Session.FormattedIO.ReadLine()

                Session.TerminationCharacterEnabled = True
                Session.TerminationCharacter = termchar

                Return Session.FormattedIO.ReadUntilMatch(Chr(termchar))

            Catch ex As Exception
                _ErrorLogger.LogException(ex, Session.ResourceName)
            End Try
        End If

        Return vbNullString

    End Function

    Public Function ReceiveValue(Optional ByRef ErrorMsg As String = "") As Double Implements IVisaDevice.ReceiveValue

        If Not Session Is Nothing Then
            Try
                Return Session.FormattedIO.ReadDouble()
            Catch ex As Exception
                _ErrorLogger.LogException(ex, Session.ResourceName)
            End Try
        End If

        Return Double.NaN

    End Function
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

End Class
