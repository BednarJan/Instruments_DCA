Public Interface ITC


#Region "Properties"

    Property SetpointTemp As Single
    Property SetpointHumidity As Single
    Property SetpointPressure As Single

    ReadOnly Property MinTemp As Single
    ReadOnly Property MaxTemp As Single

    ReadOnly Property MinHumidity As Single
    ReadOnly Property MaxHumidity As Single

    ReadOnly Property MinPressure As Single
    ReadOnly Property MaxPressure As Single

    Property HeatGrad As Single
    Property CoolGrad As Single

    Property HumiGrad As Single
    Property DeHumiGrad As Single


#End Region


#Region "Methods (Sub & Functions)"

    Sub SendCommand(cmdStr As String)
    Function ReadResponse() As String

    Sub SetTemp(ByVal val As Single)

    Function GetSetpointTemp() As Single
    Function GetInternalTemp() As Single
    Function GetProcessTemp() As Single

    Function RegTemp(ByVal val As Single, Optional ByVal accu As Single = 1) As Boolean

    Sub SetHumidity(ByVal val As Single)
    Function GetHumidity() As Single

    Sub SetPumpPressure(ByVal val As Single)
    Function GetPumpPressure() As Single

    Sub TurnOFF()

    Sub SetGradients()

#End Region

End Interface
