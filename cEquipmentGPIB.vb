Public Class cEquipmentGPIB
    Private mVendorName As String = String.Empty
    Private mEquipName As String = String.Empty
    Private mEquipSubName As String = String.Empty
    Private mAddr As String = String.Empty
    Private sHlp() As String


    Public Sub New(sIDN As String)
        sHlp = sIDN.Split(",")
        mVendorName = sHlp(0)


        Dim myHlp() As String = sHlp(1).Split("-")
        mEquipName = myHlp(0).Replace(" ", "")
        If myHlp.Length = 2 Then
            mEquipSubName = myHlp(1)
        End If

    End Sub


    Public ReadOnly Property VendorName As String
        Get
            Return mVendorName
        End Get
    End Property

    Public ReadOnly Property EquipName As String
        Get
            Return mEquipName
        End Get
    End Property

    Public ReadOnly Property EquipSubName As String
        Get
            mEquipSubName.Replace(" ", vbEmpty)
            Return mEquipSubName
        End Get
    End Property

    Public Property Addr As String
        Get
            Return mAddr
        End Get
        Set(value As String)
            mAddr = value
        End Set
    End Property

    Public ReadOnly Property AddrGPIB As String
        Get
            Return "GPIB::" & mAddr
        End Get
    End Property



End Class
