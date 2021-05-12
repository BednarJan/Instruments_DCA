Imports Ivi.Visa

Public Class CScope_Tektronix_MSO3014B
    Inherits BCScopeTEK
    Implements IScope

    Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger, 4)

    End Sub

#Region "Interface Methodes IDevice"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        MyBase.Initialize()

    End Sub

#End Region

#Region "Interface Methodes IScope"

    Public Overrides Sub CaptureScreen2File() Implements IScope.CaptureScreen2File

        MyBase.CaptureScreen2File()

    End Sub

    Public Overrides Sub SetChannel(Chan As CScopeChannel) Implements IScope.SetChannel

        SetChannel(Chan.ChanNr)

    End Sub

    Public Overrides Sub SetChannel(Nr As Integer) Implements IScope.SetChannel

        Dim _Chan As CScopeChannel = Me.Channel(Nr)

        If _Chan IsNot Nothing Then

            With _Chan

                If .State = CScopeChannel.ChanState.STATE_ON Then

                    Select Case .Bandwidth
                        Case CScopeChannel.ChanBandwidth.B20
                            Call Visa.SendString(.Name & ":Bandwidth TWEnty")
                        Case CScopeChannel.ChanBandwidth.B100
                            Call Visa.SendString(.Name & ":Bandwidth HUNdred")
                    End Select

                    Select Case .Coupling
                        Case CScopeChannel.ChanCoupling.DC
                            Call Visa.SendString(.Name & ":Coupling DC")
                        Case CScopeChannel.ChanCoupling.AC
                            Call Visa.SendString(.Name & ":Coupling AC")
                    End Select

                    Select Case .Impedance
                        Case CScopeChannel.InputImpedance.INPUT_IMPEDANCE_MEG
                            Call Visa.SendString(.Name & ":IMPEDANCE MEG")
                        Case CScopeChannel.InputImpedance.INPUT_IMPEDANCE_FIFTY
                            Call Visa.SendString(.Name & ":IMPEDANCE FIFTY")
                    End Select

                    Select Case .Polarity
                        Case CScopeChannel.ChanPolarity.NORMAL
                            Call Visa.SendString(.Name & ":INVert Off")
                        Case CScopeChannel.ChanPolarity.INVERT
                            Call Visa.SendString(.Name & ":INVert On")
                    End Select

                    Call Visa.SendString(.Name & ":POSition " & FormatNumber(.Position, 2))
                    Call Visa.SendString(.Name & ":YUnit " & Chr(34) & .DispUnit & Chr(34))

                    Select Case .DispUnit
                        Case "V"
                            Call Visa.SendString(.Name & ":PRObe:GAIN " & FormatNumber(1 / .Probe, 3))
                        Case "A"
                            Call Visa.SendString(.Name & ":PRObe:GAIN " & FormatNumber(.Probe / 1000, 3))
                    End Select

                    Call Visa.SendString(.Name & ":VOLts " & FormatNumber(.VertVolt))
                    Call Visa.SendString(.Name & ":OFFSet " & FormatNumber(.Offset))

                    Call Visa.SendString("SELect:" & .Name & " ON")
                Else
                    Call Visa.SendString("SELect:" & .Name & " OFF")

                End If

            End With

        End If

    End Sub



    Public Overrides Sub InitChannels() Implements IScope.InitChannels
        MyBase.InitChannels()
    End Sub

    Public Overrides Sub SetHorizontal() Implements IScope.SetHorizontal
        MyBase.SetHorizontal()
    End Sub

    Public Overrides Sub SetTrigger() Implements IScope.SetTrigger
        Call Visa.SendString("MEASUrement:MEAS1:STATE OFF")
        Call Visa.SendString("MEASUrement:MEAS2:STATE OFF")
        Call Visa.SendString("MEASUrement:MEAS3:STATE OFF")
        Call Visa.SendString("MEASUrement:MEAS4:STATE OFF")

        Call Visa.SendString("Trigger:A:MODe " & GetTriggerMode(MyBase.Trigger.Mode))

        Call Visa.SendString("Trigger:A:EDGe:COUPling " & GetTriggerCoupling(MyBase.Trigger.Coupling))

        Select Case MyBase.Trigger.Source
            Case CScopeTrigger.TriggerSource.Ext, CScopeTrigger.TriggerSource.Aux

                Call Visa.SendString("Trigger:A:EDGe:SOUrce AUX")
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
        MyBase.Acquire(acqState)
    End Sub

    Public Overrides Sub ClearScreen() Implements IScope.ClearScreen
        MyBase.ClearScreen()
    End Sub

    Public Overrides Sub LoadReferenceCurve(sFileName As String, nRef As Integer) Implements IScope.LoadReferenceCurve
        MyBase.LoadReferenceCurve(sFileName, nRef)
    End Sub

    Public Overrides Sub RefCurveOn(refNr As Integer) Implements IScope.RefCurveOn
        MyBase.RefCurveOn(refNr)
    End Sub

    Public Overrides Sub RefCurveOff(refNr As Integer) Implements IScope.RefCurveOff
        MyBase.RefCurveOff(refNr)
    End Sub

    Public Overrides Function MeasDelay(MeasNr As Integer, Source1 As String, slope1 As Integer, Source2 As String, slope2 As Integer) As Single Implements IScope.MeasDelay

        Return MyBase.MeasDelay(MeasNr, Source1, slope1, Source2, slope2)

    End Function

    Public Overrides Function MeasEdge(MeasNr As Integer, Source As String, LowRefLevel As Integer, HighRefLevel As Integer, Slope As Integer) As Single Implements IScope.MeasEdge

        Return MyBase.MeasEdge(MeasNr, Source, LowRefLevel, HighRefLevel, Slope)

    End Function

    Public Overrides Function MeasFreq(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasFreq

        Return MyBase.MeasFreq(MeasNr, Source1)

    End Function

    Public Overrides Function MeasPK2PK(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasPK2PK

        Return MyBase.MeasPK2PK(MeasNr, Source1)

    End Function

    Public Overrides Function MeasRMS(MeasNr As Integer, Source1 As String, Optional waitTime As Integer = 1) As Single Implements IScope.MeasRMS

        Return MyBase.MeasRMS(MeasNr, Source1, waitTime)

    End Function

    Public Overrides Function MeasPOVERSHOOT(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasPOVERSHOOT

        Return MyBase.MeasPOVERSHOOT(MeasNr, Source1)

    End Function

    Public Overrides Function MeasNOVERSHOOT(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasNOVERSHOOT

        Return MyBase.MeasNOVERSHOOT(MeasNr, Source1)

    End Function

    Public Overrides Function MeasIMAX(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasIMAX

        Return MyBase.MeasIMAX(MeasNr, Source1)

    End Function

    Public Overrides Function MeasIMIN(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasIMIN

        Return MyBase.MeasIMIN(MeasNr, Source1)

    End Function

    Public Overrides Function MeasHIGH(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasHIGH

        Return MyBase.MeasHIGH(MeasNr, Source1)

    End Function

    Public Overrides Function MeasLOW(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasLOW

        Return MyBase.MeasLOW(MeasNr, Source1)

    End Function


#End Region


End Class
