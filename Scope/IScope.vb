Public Interface IScope
    Inherits IDevice


#Region "Properties"

    Property Channel(_ChanNr As Integer) As CScopeChannel
    Property HardcopyFullFileName As String
    Property Channels As CScopeChannels
    Property TimeBase As Single
    Property Trigger As CScopeTrigger
    ReadOnly Property HardcopyFileFormat As String


#End Region

#Region "Methods (Sub & Functions)"


    Sub CaptureScreen2File()
    Sub InitChannel(Nr As Integer)
    Sub InitChannel(Chan As CScopeChannel)

    Sub InitChannels()
    Sub SetChannel(Nr As Integer)
    Sub SetChannel(Chan As CScopeChannel)
    Sub SetChannels()
    Sub SetHorizontal()
    Sub SetTrigger()
    Sub Acquire(acqState As CScopeTrigger.Acquire)
    Sub ClearScreen()
    Sub LoadReferenceCurve(sFileName As String, nRef As Integer)
    Sub RefCurveOn(refNr As Integer)
    Sub RefCurveOff(refNr As Integer)
    Sub SetColorScheme(colScheme As String)

    Function MeasDelay(ByVal MeasNr As Integer, ByVal Source1 As String, ByVal slope1 As Integer, ByVal Source2 As String, ByVal slope2 As Integer) As Single
    Function MeasDelay(ByVal MeasNr As Integer, ByVal Chan1 As CScopeChannel, ByVal slope1 As Integer, ByVal Chan2 As CScopeChannel, ByVal slope2 As Integer) As Single

    Function MeasEdge(ByVal MeasNr As Integer, ByVal SOURCE As String, ByVal lowRefLevel As Integer, ByVal highRefLevel As Integer, ByVal slope As Integer) As Single
    Function MeasEdge(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel, ByVal lowRefLevel As Integer, ByVal highRefLevel As Integer, ByVal slope As Integer) As Single


    Function MeasFreq(ByVal MeasNr As Integer, ByVal sSource1 As String) As Single
    Function MeasFreq(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel) As Single

    Function MeasPK2PK(ByVal MeasNr As Integer, ByVal Source1 As String) As Single
    Function MeasPK2PK(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel) As Single

    Function MeasRMS(ByVal MeasNr As Integer, ByVal sSource1 As String, Optional waitTime As Integer = 1) As Single
    Function MeasRMS(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel, Optional waitTime As Integer = 1) As Single

    Function MeasPOVERSHOOT(ByVal MeasNr As Integer, ByVal sSource1 As String) As Single
    Function MeasPOVERSHOOT(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel) As Single

    Function MeasNOVERSHOOT(ByVal MeasNr As Integer, ByVal sSource1 As String) As Single
    Function MeasNOVERSHOOT(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel) As Single

    Function MeasIMAX(ByVal MeasNr As Integer, ByVal sSource1 As String) As Single
    Function MeasIMAX(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel) As Single

    Function MeasIMIN(ByVal MeasNr As Integer, ByVal sSource1 As String) As Single
    Function MeasIMIN(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel) As Single

    Function MeasHIGH(ByVal MeasNr As Integer, ByVal sSource1 As String) As Single
    Function MeasHIGH(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel) As Single

    Function MeasLOW(ByVal MeasNr As Integer, ByVal sSource1 As String) As Single
    Function MeasLOW(ByVal MeasNr As Integer, ByVal Chan As CScopeChannel) As Single

#End Region


End Interface

