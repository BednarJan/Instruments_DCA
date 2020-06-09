Public Interface IPWAN

#Region "Properties"
    ReadOnly Property VoltageMax As Single
    ReadOnly Property CurrentMax As Single
    ReadOnly Property InputElements As UInteger
#End Region

#Region "Methods (Sub & Functions)"
    Function GetPF(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    Function GetVrms(Optional ByVal WaitOnPWAN As Boolean = False) As Single
    Function GetVthd(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    Function GetIrms(Optional ByVal WaitOnPWAN As Boolean = False) As Single
    Function GetIthd(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    Function GetPreal(Optional ByVal WaitOnPWAN As Boolean = False) As Single
    Function GetPapp(Optional ByVal WaitOnPWAN As Boolean = False) As Single
    Function GetPreact(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    Function GetFreq(Optional ByVal WaitOnPWAN As Boolean = False) As Single

    Sub GetHarmonics_IEC61000_3_2(EquipClass As CPWAN_Helper.EEquipmentClass, ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, Optional ByVal WaitOnPWAN As Boolean = False)
    Sub GetHarmonics_IEC61000_3_12(ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, ByRef THD As Single, ByRef PWHD As Single, Optional ByVal WaitOnPWAN As Boolean = False)
    'Function IntegratePower(durationInSec As Single) As Double
    'Sub Init_Integration(durationInSec As Single)
    'Sub Get_Integrations_Data(ByRef measData() As Double)
    'Sub Set_Volt_Range(sRange As String)
    'Sub Set_Curr_Range(sRange As String)

#End Region
End Interface