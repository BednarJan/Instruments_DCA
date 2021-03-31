Public Interface IScope
    Inherits IDevice


#Region "Properties"

    Property HardcopyFullFileName As String
    Property Channels As List(Of CScopeChannel)
    Property TimeBase As Single
    Property Trigger As CScopeTrigger
    Property HardcopyFileFormat As UInteger

#End Region

End Interface

