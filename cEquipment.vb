Public Class cEquipment
    Private mEquipName As String = String.Empty
    Private mEquipAddr As String = String.Empty

    Public Property EquipName As String
        Get
            Return mEquipName
        End Get
        Set(value As String)
            mEquipName = value
        End Set
    End Property

    Public Property EquipAddr As String
        Get
            Return mEquipAddr
        End Get
        Set(value As String)
            mEquipAddr = value
        End Set
    End Property

End Class
