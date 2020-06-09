Public Interface IDAQ

#Region "Properties (Set & Get)"

#End Region

#Region "Methods (Sub & Functions)"
    Function Get_Volt_DC(ByVal Chan As Integer) As Single
    Function Get_Volt_AC(ByVal Chan As Integer) As Single
    Function Get_Sample_Volt_DC(ByVal Chan As Integer, ByVal Sample As Integer) As Single
    Function Get_Sample_Volt_Dyn(ByVal Chan As Integer, ByVal Sample As Integer) As Single
    Function Get_Curr_DC(ByVal Chan As Integer) As Single
    Function Get_Curr_AC(ByVal Chan As Integer) As Single
    Function Get_Res(ByVal Chan As Integer) As Single
    Function Get_FRes(ByVal Chan As Integer) As Single
    Function Get_FReq(ByVal Chan As Integer) As Single

#End Region

End Interface
