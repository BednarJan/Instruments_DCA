Imports Ivi.Visa
Public Class BCScopeTEK
    Inherits BCScope
    Implements IScope


    Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger, Optional myCountOfChannels As Integer = 4)

        MyBase.New(Session, ErrorLogger, myCountOfChannels)

    End Sub

#Region "Interface Methodes IScope"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        MyBase.Initialize()

        Dim strDate As String = "'" & Format(Now(), "YYYY") & "-"
        strDate &= Format(Now(), "MM") & "-"
        strDate &= Format(Now(), "DD") & "'"

        Dim strTime As String = "'" & Format(Now(), "HH:NN:SS") & "'"



        Visa.SendString("COMMUNICATE:REMOTE ON")
        Visa.SendString("DATE " & strDate)
        Visa.SendString("TIME " & strTime)

    End Sub

#End Region

#Region "Interface Methodes IScope"

    Public Overrides Sub PrintDisplay2File() Implements IScope.PrintDisplay2File

        Call Visa.SendString("HARDCOPY START")
        Call cHelper.Delay(2)
        Call Visa.ReadStringToFileRAW(MyBase.HardcopyFullFileName)

    End Sub


    Public Overrides Sub SetChannels() Implements IScope.SetChannels
        MyBase.SetChannels()
    End Sub

    Public Overrides Sub InitChannels() Implements IScope.InitChannels
        MyBase.InitChannels()
    End Sub

    Public Overrides Sub SetHorizontal() Implements IScope.SetHorizontal
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub SetTrigger() Implements IScope.SetTrigger
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub Acquire(acqState As Integer) Implements IScope.Acquire
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub ClearScreen() Implements IScope.ClearScreen
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub LoadReferenceCurve(sFileName As String, nRef As Integer) Implements IScope.LoadReferenceCurve
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub RefCurveOn(refNr As Integer) Implements IScope.RefCurveOn
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub RefCurveOff(refNr As Integer) Implements IScope.RefCurveOff
        Throw New NotImplementedException()
    End Sub

    Public Overrides Function MeasDelay(MeasNr As Integer, Source1 As String, slope1 As Integer, Source2 As String, slope2 As Integer) As Single Implements IScope.MeasDelay
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasEdge(MeasNr As Integer, SOURCE As String, lowRefLevel As Integer, highRefLevel As Integer, slope As Integer) As Single Implements IScope.MeasEdge
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasFreq(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasFreq
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasPK2PK(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasPK2PK
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasRMS(MeasNr As Integer, sSource1 As String, Optional waitTime As Integer = 1) As Single Implements IScope.MeasRMS
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasPOVERSHOOT(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasPOVERSHOOT
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasNOVERSHOOT(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasNOVERSHOOT
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasIMAX(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasIMAX
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasIMIN(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasIMIN
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasHIGH(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasHIGH
        Throw New NotImplementedException()
    End Function

    Public Overrides Function MeasLOW(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasLOW
        Throw New NotImplementedException()
    End Function

#End Region

End Class
