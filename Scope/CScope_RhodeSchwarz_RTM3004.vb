Imports Ivi.Visa
Public Class CScope_RhodeSchwarz_RTM3004

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

        ClearScreen()

        Visa.SendString("MENU OFF")
        Visa.SendString("HCOPy:FORMat " & HardcopyFileFormat)
        Visa.SendString("HCOPy:COLor:SCHeme INVerted")

        Visa.SendString("SYSTEM:DATE " & strDate)
        Visa.SendString("SYSTEM:TIME " & strTime)

    End Sub

#End Region

#Region "Interface Methodes IScope"

    Public Overrides Sub CaptureScreen2File() Implements IScope.CaptureScreen2File

        Call Visa.SendString("HCOPy:DATA?")
        Call cHelper.Delay(2)
        Call Visa.ReadStringToFileRAW(HardcopyFullFileName, "‰")

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

        Dim _Chan As CScopeChannel = Me.Channel(Nr)

        If _Chan IsNot Nothing Then

            With _Chan

                If .State = CScopeChannel.ChanState.STATE_ON Then

                    Call Visa.SendString("CHANnel" & .ChanNr & " :STATe ON")

                    Select Case .Bandwidth
                        Case CScopeChannel.ChanBandwidth.B20
                            Call Visa.SendString("CHANnel" & .ChanNr & ":Bandwidth B20")
                        Case CScopeChannel.ChanBandwidth.B100
                            Call Visa.SendString("CHANnel" & .ChanNr & ":Bandwidth FULL")
                    End Select

                    Select Case .Coupling

                        Case CScopeChannel.ChanCoupling.DC
                            Call Visa.SendString("CHANnel" & .ChanNr & ":Coupling DCLimit")
                        Case CScopeChannel.ChanCoupling.AC
                            Call Visa.SendString("CHANnel" & .ChanNr & ":Coupling ACLimit")
                        Case CScopeChannel.ChanCoupling.GND
                            Call Visa.SendString("CHANnel" & .ChanNr & ":Coupling GND")

                    End Select

                    Select Case .Impedance
                        Case CScopeChannel.InputImpedance.INPUT_IMPEDANCE_MEG
                            Call Visa.SendString("PROBE" & .ChanNr & ":SETUP:IMP 1MOHm")
                        Case CScopeChannel.InputImpedance.INPUT_IMPEDANCE_FIFTY
                            Call Visa.SendString("PROBE" & .ChanNr & ":SETUP:IMP 50OHm")
                    End Select

                    Select Case .Polarity
                        Case CScopeChannel.ChanPolarity.NORMAL
                            Call Visa.SendString("CHANnel" & .ChanNr & ":POLarity NORMal")
                        Case CScopeChannel.ChanPolarity.INVERT
                            Call Visa.SendString("CHANnel" & .ChanNr & ":POLarity INVerted")
                    End Select

                    Call Visa.SendString("CHANnel" & .ChanNr & ":POSition " & FormatNumber(.Position, 2))

                    Select Case .DispUnit

                        Case "V"

                            Call Visa.SendString("PROBe" & .ChanNr & ":SETup:ATTenuation:UNIT V")
                            Call Visa.SendString("PROBe" & .ChanNr & ":SETup:ATTenuation:MANual " & FormatNumber(.Probe, 3))

                        Case "A"

                            Call Visa.SendString("PROBe" & .ChanNr & ":SETup:ATTenuation:UNIT A")
                            Call Visa.SendString("PROBe" & .ChanNr & ":SETup:ATTenuation:MANual " & 1000 / FormatNumber(.Probe, 3))

                    End Select

                    Call Visa.SendString("CHANnel" & .ChanNr & ":SCALe " & FormatNumber(.VertVolt, 3))
                    Call Visa.SendString("CHANnel" & .ChanNr & ":OFFSet " & FormatNumber(.Offset))


                    Call Visa.SendString("CHANnel" & .ChanNr & ":STATe OFF")

                End If


            End With

        End If

    End Sub

    Public Overrides Sub SetHorizontal() Implements IScope.SetHorizontal

        Call Visa.SendString("TIMebase:REFerence 8.33")
        Call Visa.SendString("TIMebase:POSition 0")
        Call Visa.SendString("TIMebase:SCALe " & FormatNumber(TimeBase, 6) & "")

    End Sub

    Public Overrides Sub SetTrigger() Implements IScope.SetTrigger

        For i As Integer = 1 To 8

            Call Visa.SendString("MEASurement" & i & ":ENABle OFF")

        Next i

        Call Visa.SendString("Trigger:A:MODe " & GetTriggerMode(Trigger.Mode))
        Call Visa.SendString("Trigger:A:TYPe EDGE")
        Call Visa.SendString("Trigger:A:EDGe:COUPling " & GetTriggerCoupling(Trigger.Coupling))

        Select Case Trigger.Coupling

            Case CScopeTrigger.TriggerCoupling.DC, CScopeTrigger.TriggerCoupling.AC, CScopeTrigger.TriggerCoupling.LFReject

                Call Visa.SendString("TRIGger:A:EDGE:FILTer:HFReject OFF")
                Call Visa.SendString("TRIGger:A:EDGE:FILTer:NREJect OFF")

            Case CScopeTrigger.TriggerCoupling.DC_HFReject, CScopeTrigger.TriggerCoupling.AC_HFReject

                Call Visa.SendString("TRIGger:A:EDGE:FILTer:NREJect OFF")
                Call Visa.SendString("TRIGger:A:EDGE:FILTer:HFReject ON")

            Case CScopeTrigger.TriggerCoupling.DC_NREJect, CScopeTrigger.TriggerCoupling.AC_NREJect

                Call Visa.SendString("TRIGger:A:EDGE:FILTer:HFReject OFF")
                Call Visa.SendString("TRIGger:A:EDGE:FILTer:NREJect ON")

        End Select

        Select Case Trigger.Source
            Case CScopeTrigger.TriggerSource.Ext, CScopeTrigger.TriggerSource.Aux

                Call Visa.SendString("TRIGger:A:SOURce EXTernanalog")
                Call Visa.SendString("Trigger:A:Level TTL")

            Case Else

                Call Visa.SendString("TRIGger:A:SOURce " & UCase(Trigger.Source.ToString))
                Call Visa.SendString("TRIGger:A:Level" & UCase(Trigger.Source.ToString) & ":VALue " & FormatNumber(Trigger.Level))

        End Select

        Call Visa.SendString("Trigger:A:EDGe:SLOpe " & GetSlope(Trigger.Slope))

        Call Visa.SendString("TIMebase:POSition " & FormatNumber(Trigger.Position))

    End Sub

    Public Overrides Sub Acquire(acqState As CScopeTrigger.Acquire) Implements IScope.Acquire

        Call Visa.SendString("ACQuire:POINts:AUTomatic ON")

        Select Case acqState

            Case Trigger.Acquire.STATE_RUN
                Call Visa.SendString("ACQuire:STATE RUN")
            Case Trigger.Acquire.STATE_STOP
                Call Visa.SendString("ACQuire:STATE STOP")
            Case Trigger.Acquire.STATE_SINGLE
                Call Visa.SendString("ACQuire:STATE SINGLE")

        End Select

        cHelper.Delay(0.5)
    End Sub

    Public Overrides Sub ClearScreen() Implements IScope.ClearScreen

        Call Visa.SendString("DISPlay:CLEar:SCReen")

    End Sub

    Public Overrides Sub LoadReferenceCurve(sFileName As String, nRef As Integer) Implements IScope.LoadReferenceCurve
        Dim CmdString As String

        CmdString = "REFCurve" & nRef & ":LOAD " & Chr(34) & sFileName & Chr(34)

        Call Visa.SendString(CmdString)
    End Sub

    Public Overrides Sub RefCurveOn(refNr As Integer) Implements IScope.RefCurveOn

        Call Visa.SendString("REFCURVE" & refNr & ":STATE On")

    End Sub

    Public Overrides Sub RefCurveOff(refNr As Integer) Implements IScope.RefCurveOff

        Call Visa.SendString("REFCURVE" & refNr & ":STATE Off")

    End Sub

    Public Overrides Function MeasDelay(MeasNr As Integer, Source1 As String, slope1 As Integer, Source2 As String, slope2 As Integer) As Single Implements IScope.MeasDelay

        Dim mySlope1 As String = GetSlope(slope1)
        Dim mySlope2 As String = GetSlope(slope2)

        Dim mySource1 As String = GetMeasSource(Source1)
        Dim mySource2 As String = GetMeasSource(Source2)

        Call Visa.SendString("MEASUrement" & MeasNr & ":MAIN DELay")
        Call Visa.SendString("MEASUrement" & MeasNr & ":SOURce " & mySource1 & "," & mySource2)
        Call Visa.SendString("MEASUrement" & MeasNr & ":DELay:SLOPe " & mySlope1 & "," & mySlope2)
        Call cHelper.Delay(1)

        Return MeasDecimalValue(MeasNr)

    End Function

    Public Overrides Function MeasEdge(MeasNr As Integer, Source As String, LowRefLevel As Integer, HighRefLevel As Integer, Slope As Integer) As Single Implements IScope.MeasEdge

        Dim sSlope As String = GetSlope(Slope)

        Call Visa.SendString("REFLevel:RELative:MODE USER")
        Call Visa.SendString("REFLevel:RELative:LOWer " & FormatNumber(LowRefLevel))
        Call Visa.SendString("REFLevel:RELative:MIDDle 50")
        Call Visa.SendString("REFLevel:RELative:UPPer " & FormatNumber(HighRefLevel))
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":MAIN " & sSlope)
        Call Visa.SendString("MEASurement" & MeasNr & ":SOURce " & Source)

        Call cHelper.Delay(1)

        Return MeasDecimalValue(MeasNr)

    End Function

    Public Overrides Function MeasFreq(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasFreq

        Return MeasSource(MeasNr, Source1, "FREQ")

    End Function

    Public Overrides Function MeasPK2PK(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasPK2PK

        Return MeasSource(MeasNr, Source1, "PEAK")

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

        Return MeasSource(MeasNr, Source1, "UPEakvalue")

    End Function

    Public Overrides Function MeasIMIN(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasIMIN

        Return MeasSource(MeasNr, Source1, "LPEakvalue")

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

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":MAIN " & UCase(MeasType))
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":SOURCE " & Source)

        Call cHelper.Delay(1)

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":RESult?")

        Return CDec(Visa.ReceiveValue)

    End Function



    Overrides Function MeasDecimalValue(MeasNr As Integer) As Decimal

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":RESult?")

        Return cHelper.StringToDecimal(Visa.ReceiveValue)

    End Function


    Overrides Function GetSlope(nSlope As Integer) As String
        Dim sRet As String = String.Empty

        Select Case nSlope
            Case CScopeTrigger.TriggerSlope.RISE
                sRet = "RTIMe"
            Case CScopeTrigger.TriggerSlope.FALL
                sRet = "FTIMe"
            Case CScopeTrigger.TriggerSlope.EITHER
                sRet = "RTIMe"
        End Select

        Return sRet
    End Function


    Private Function GetMeasSource(sSource As String) As String
        Dim mySource As String
        mySource = sSource
        If InStr(UCase(sSource), "REF") Then
            mySource = "RE" & Right(sSource, 1)
        End If
        GetMeasSource = mySource
    End Function


#End Region

End Class
