Imports System.ComponentModel

Public Module OfflineHandler

#Region "Interface"
    'Forces all classes with this interface to implement the cOfflineHandler
    Public Interface IOfflineDevice
        Property Offline As cOfflineHandler
    End Interface
#End Region

#Region "Offline Handler Class"
    'This Class creates a Dictionary of all OfflineValues in an IOfflineDevice
    Public Class cOfflineHandler
        Public Property Values As New Dictionary(Of String, cOfflineValue)
    End Class
#End Region

#Region "Offline Value Class"
    'This Class notifies via PropertyChanged Event if its Value has changed. This Event can
    'be handled externally
    Public Class cOfflineValue
        Implements INotifyPropertyChanged

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private _Value As VariantType

        Public Sub New(OfflineSystem As cOfflineHandler, EventName As String)

            OfflineSystem.Values.Add(EventName, Me)

        End Sub

        Public Property Value() As VariantType
            Get
                Value = _Value
            End Get
            Set(ByVal value As VariantType)
                If value <> _Value Then
                    _Value = value
                    RaiseEvent PropertyChanged(Me, Nothing)
                End If
            End Set
        End Property
    End Class
#End Region

End Module

