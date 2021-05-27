Imports Ivi.Visa

Public Interface IVisaDevice

    'Properties
    ReadOnly Property ErrorLogger As CErrorLogger
    ReadOnly Property Session As IMessageBasedSession

    'Communication Methods
    Sub SendString(CMD As String, Optional ByRef ErrorMsg As String = "")
    Function ReceiveString(Optional ByRef ErrorMsg As String = "", Optional termchar As Byte = 10) As String
    Function ReceiveValue(Optional ByRef ErrorMsg As String = "") As Double

    Function ReceiveValueList(separ As Char, Optional ByRef ErrorMsg As String = "") As Double()

    Sub ReadStringToFileRAW(ByVal HardcopyFullFileName As String, Optional fromFirstChar As Char = "", Optional termchar As Byte = 10)

    'Session Config (User Interface)
    Sub ShowConfigUI()

End Interface
