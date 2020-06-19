Imports Ivi.Visa

Public Interface IVisaDevice

    'Properties
    ReadOnly Property ErrorLogger As CErrorLogger
    ReadOnly Property Session As IMessageBasedSession

    'Communication Methods
    Sub SendString(CMD As String, Optional ByRef ErrorMsg As String = "")
    Function ReceiveString(Optional ByRef ErrorMsg As String = "", Optional termchar As Byte = 10) As String
    Function ReceiveValue(Optional ByRef ErrorMsg As String = "") As Double

    'Session Config (User Interface)
    Sub ShowConfigUI()

End Interface
