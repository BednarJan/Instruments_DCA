#Region "Enums"
Public Enum ESource
    TRACE
    TRIGGER
End Enum
Public Enum ESlope
    RISE
    FALL
End Enum
Public Enum ESignal
    CHANNEL
    MATH
    LOGIC
End Enum
#End Region

Public Interface IScope

#Region "Properties (Set & Get)"
    ReadOnly Property VoltageMax As Single
    Property File2SavePicture As String
#End Region

#Region "Standard GPIB Methods"
    Sub RST()
    Sub CLS()
    Function IDN() As String
#End Region

#Region "Methods (Sub & Functions)"
    Sub Initialize()
    Function Meas_Delay(ByVal MeasNr As Integer, ByVal Source1 As Integer, ByVal edge1 As String, ByVal Source2 As Integer, ByVal Edge2 As String) As Single
    Function Meas_Freq(ByVal MeasNr As Integer, ByVal Source As Integer) As Single
    Function Meas_Pk2Pk(ByVal MeasNr As Integer, ByVal Source As Integer) As Single
    Function Meas_RMS(ByVal MeasNr As Integer, ByVal Source As Integer) As Single
    Function Meas_Povershoot(ByVal MeasNr As Integer, ByVal Source As Integer) As Single
    Function Meas_Novershoot(ByVal MeasNr As Integer, ByVal Source As Integer) As Single
    Function Meas_Imax(ByVal MeasNr As Integer, ByVal Source As Integer) As Single
    Function Meas_Imin(ByVal MeasNr As Integer, ByVal Source As Integer) As Single
    Sub Print_Display_to_File()

#End Region
End Interface
