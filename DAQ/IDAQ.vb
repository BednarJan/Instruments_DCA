Public Interface IDAQ

#Region "Properties (Set & Get)"

    Property VoltageMax As Single
    Property CurrentMax As Single

#End Region


#Region "Enum definitions"

    Enum DAQ_Function
        VOLT_DC = 1
        VOLT_AC = 2
        CURRENT_DC = 3
        CURRENT_AC = 4
        RES = 5
        FRES = 6
        FREQ = 7
    End Enum

#End Region

#Region "Methods (Sub & Functions)"

    Property ScanList As CScanList

    Function Get_Volt_DC(ByVal Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single
    Function Get_Volt_AC(ByVal Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single
    Function Get_Curr_DC(ByVal Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single
    Function Get_Curr_AC(ByVal Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single
    Function Get_Res(ByVal Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single
    Function Get_FRes(ByVal Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single
    Function Get_FReq(ByVal Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single

    Function Get_Sample_Volt_DC(ByVal Chan As Integer, ByVal Sample As Integer) As Single

    'Function ScanChannels(ChanList As Integer(), Funct As IDAQ.DAQ_Function) As Decimal()

#End Region

End Interface
