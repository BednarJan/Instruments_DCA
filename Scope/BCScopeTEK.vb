Imports Ivi.Visa
Public Class BCScopeTEK
    Inherits BCScope
    Implements IScope


    Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger, Optional myCountOfChannels As Integer = 4)

        MyBase.New(Session, ErrorLogger, myCountOfChannels)

    End Sub

#Region "Interface Methodes IDevice"

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

    Public Overrides Sub CaptureScreen2File() Implements IScope.CaptureScreen2File

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

    Public Overrides Sub SetChannel(Chan As CScopeChannel) Implements IScope.SetChannel

        MyBase.SetChannel(Chan)

    End Sub

    Public Overrides Sub SetChannel(Nr As Integer) Implements IScope.SetChannel

        MyBase.SetChannel(Nr)

    End Sub

    Public Overrides Sub SetHorizontal() Implements IScope.SetHorizontal
        Call Visa.SendString("HORizontal:RESOlution high")
        Call Visa.SendString("HORizontal:Delay:State off")
        Call Visa.SendString("HORizontal:SCALe " & FormatNumber(MyBase.TimeBase))
    End Sub

    Public Overrides Sub SetTrigger() Implements IScope.SetTrigger
        Call Visa.SendString("MEASUrement:MEAS1:STATE OFF")
        Call Visa.SendString("MEASUrement:MEAS2:STATE OFF")
        Call Visa.SendString("MEASUrement:MEAS3:STATE OFF")
        Call Visa.SendString("MEASUrement:MEAS4:STATE OFF")

        Call Visa.SendString("Trigger:A:MODe " & GetTriggerMode(MyBase.Trigger.Mode))

        Call Visa.SendString("Trigger:A:EDGe:COUPling " & GetTriggerCoupling(MyBase.Trigger.Coupling))

        Select Case MyBase.Trigger.Source
            Case CScopeTrigger.TriggerSource.Ext

                Call Visa.SendString("Trigger:A:EDGe:SOUrce EXT10")
                Call Visa.SendString("Trigger:A:Level TTL")

            Case Else

                'Call Visa.SendString("Trigger:A:EDGe:SOUrce CH" & MyBase.Trigger.TriggerSource)
                Call Visa.SendString("Trigger:A:EDGe:SOUrce " & UCase(MyBase.Trigger.Source.ToString))
                Call Visa.SendString("Trigger:A:Level " & FormatNumber(MyBase.Trigger.Level))
        End Select

        Call cHelper.Delay(1)
        Call Visa.SendString("Trigger:A:EDGe:SLOpe " & GetSlope(MyBase.Trigger.Slope))

        Call Visa.SendString("HORizontal:TRIGger:Position " & FormatNumber(MyBase.Trigger.Position))

    End Sub

    Public Overrides Sub Acquire(acqState As CScopeTrigger.Acquire) Implements IScope.Acquire
        Call Visa.SendString("ACQuire:MODe SAMple")
        Call Visa.SendString("ACQuire:STOPAfter SEQuence")
        Select Case acqState
            Case MyBase.Trigger.Acquire.STATE_RUN
                Call Visa.SendString("ACQuire:STATE RUN")
            Case Else
                Call Visa.SendString("ACQuire:STATE " & acqState)
        End Select

        cHelper.Delay(0.5)
    End Sub

    Public Overrides Sub ClearScreen() Implements IScope.ClearScreen
        MyBase.ClearScreen()
    End Sub

    Public Overrides Sub LoadReferenceCurve(sFileName As String, nRef As Integer) Implements IScope.LoadReferenceCurve
        Dim CmdString As String

        CmdString = "RECAll:WAWEform " & Chr(34) & sFileName & Chr(34) & ",REF" & nRef

        Call Visa.SendString(CmdString)
    End Sub

    Public Overrides Sub RefCurveOn(refNr As Integer) Implements IScope.RefCurveOn

        Call Visa.SendString("SELect:REF" & refNr & " On")

    End Sub

    Public Overrides Sub RefCurveOff(refNr As Integer) Implements IScope.RefCurveOff

        Call Visa.SendString("SELect:REF" & refNr & " Off")

    End Sub

    Public Overrides Function MeasDelay(MeasNr As Integer, Source1 As String, slope1 As Integer, Source2 As String, slope2 As Integer) As Single Implements IScope.MeasDelay
        Dim edge1 As String
        Dim edge2 As String


        Call Visa.SendString("MEASUrement:METHod MINMax")
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":TYPe DELAY")


        If InStr(Source1, "Ref") Then
            Call Visa.SendString("SELect:" & Source1 & " On")
        End If

        edge1 = GetSlope(slope1)
        edge2 = GetSlope(slope2)

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":SOURCE1 " & Source1)

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":SOURCE2 " & Source2)
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":DEL:EDGE1 " & edge1)
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":DEL:EDGE2 " & edge2)
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":DEL:DIRECTION FORW")
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":INDI:STATE " & MeasNr)
        Call Visa.SendString("MEASUrement:REFLevel:PERCent:MID 45")
        Call Visa.SendString("MEASUrement:REFLevel:PERCent:MID2 90")
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":STATE on")
        Call Visa.SendString("MEASUrement:INDICators:State Meas" & MeasNr)


        Call cHelper.Delay(3)

        Return MeasDecimalValue(MeasNr)

    End Function

    Public Overrides Function MeasEdge(MeasNr As Integer, Source As String, LowRefLevel As Integer, HighRefLevel As Integer, Slope As Integer) As Single Implements IScope.MeasEdge

        Dim edge As String = GetSlope(Slope)

        Call Visa.SendString("MEASUrement:REFlevel:METHod PERCent")
        Call Visa.SendString("MEASUrement:REFLevel:PERCent:LOW " & FormatNumber(LowRefLevel))
        Call Visa.SendString("MEASUrement:REFLevel:PERCent:HIGH " & FormatNumber(HighRefLevel))
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":SOURCE " & Source)
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":STATE ON")
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":TYPe " & edge)

        Call cHelper.Delay(3)

        Return MeasDecimalValue(MeasNr)

    End Function

    Public Overrides Function MeasFreq(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasFreq

        Return MeasSource(MeasNr, Source1, "FREQ")

    End Function

    Public Overrides Function MeasPK2PK(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasPK2PK

        Return MeasSource(MeasNr, Source1, "PK2PK")

    End Function

    Public Overrides Function MeasRMS(MeasNr As Integer, Source1 As String, Optional waitTime As Integer = 1) As Single Implements IScope.MeasRMS

        Return MeasSource(MeasNr, Source1, "RMS")

    End Function

    Public Overrides Function MeasPOVERSHOOT(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasPOVERSHOOT

        Return MeasSource(MeasNr, Source1, "POV")

    End Function

    Public Overrides Function MeasNOVERSHOOT(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasNOVERSHOOT

        Return MeasSource(MeasNr, Source1, "NOV")

    End Function

    Public Overrides Function MeasIMAX(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasIMAX

        Return MeasSource(MeasNr, Source1, "MAXI")

    End Function

    Public Overrides Function MeasIMIN(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasIMIN

        Return MeasSource(MeasNr, Source1, "MINI")

    End Function

    Public Overrides Function MeasHIGH(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasHIGH

        Return MeasSource(MeasNr, Source1, "HIGH")

    End Function

    Public Overrides Function MeasLOW(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasLOW

        Return MeasSource(MeasNr, Source1, "LOW")

    End Function


#End Region

#Region "Help Meas Functions"

    Overrides Function MeasSource(MeasNr As Integer, Source As String, MeasType As String) As Single

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":SOURCE " & Source)
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":STATE ON")
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":TYPe " & MeasType)

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":VALue?")

        Return CDec(Visa.ReceiveValue)

    End Function


    Overrides Function MeasDecimalValue(MeasNr As Integer) As Decimal


        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":VALue?")

        Return CDec(Visa.ReceiveValue)

    End Function

    Overrides Function GetSlope(nSlope As Integer) As String
        Dim sRet As String = String.Empty

        Select Case nSlope
            Case CScopeTrigger.TriggerSlope.RISE
                sRet = "rise"
            Case CScopeTrigger.TriggerSlope.FALL
                sRet = "fall"
        End Select

        Return sRet
    End Function


#End Region








End Class
