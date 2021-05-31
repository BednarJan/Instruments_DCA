Imports Ivi.Visa
Public Class BCScope
    Inherits BCDevice
    Implements IScope

    Private _HardcopyFileFormat As String = "PNG"
    Private _HardcopyFullFileName As String

#Region "Shorthand Properties"

    Property CountOfChannels As Integer

    Public Property Channel(_ChanNr As Integer) As CScopeChannel Implements IScope.Channel
        Get
            Return _Channels.Channel(_ChanNr)
        End Get
        Set(value As CScopeChannel)
            _Channels.Channel(_ChanNr) = value
        End Set
    End Property

    Public Overridable Property HardcopyFullFileName As String Implements IScope.HardcopyFullFileName
        Get
            Return _HardcopyFullFileName
        End Get
        Set(value As String)

            _HardcopyFullFileName = value

            Dim fi As New System.IO.FileInfo(value)
            _HardcopyFileFormat = UCase(fi.Extension)
            _HardcopyFileFormat = _HardcopyFileFormat.Replace(".", "")

            Visa.SendString("HCOPy:FORMat " & _HardcopyFileFormat)

        End Set
    End Property


    Public Property Channels As CScopeChannels Implements IScope.Channels

    Public Property TimeBase As Single Implements IScope.TimeBase

    Public Property Trigger As CScopeTrigger Implements IScope.Trigger

    Public ReadOnly Property HardcopyFileFormat As String Implements IScope.HardcopyFileFormat
        Get

            If Not String.IsNullOrEmpty(_HardcopyFullFileName) Then

                Dim fi As New IO.FileInfo(_HardcopyFullFileName)

                Return fi.Extension
            Else

                Return _HardcopyFileFormat

            End If

        End Get
    End Property

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger, Optional myCountOfChannels As Integer = 4)

        MyBase.New(Session, ErrorLogger)

        _Channels = New CScopeChannels

        'Default settings

        _CountOfChannels = myCountOfChannels

        CreateChannels()

        'Create Trigger with default settings

        _Trigger = New CScopeTrigger(ErrorLogger) With {
         .Slope = CScopeTrigger.TriggerSlope.RISE,
        .Source = CScopeTrigger.TriggerSource.Ch1,
        .Coupling = CScopeTrigger.TriggerCoupling.AC_HFReject,
        .Mode = CScopeTrigger.TriggerMode.SIGNL
        }

    End Sub

#End Region


#Region "Basic Device Functions (IDevice)"

    Public Overrides Function IDN() As String Implements IDevice.IDN
        Return MyBase.IDN()
    End Function

    Public Overrides Sub RST() Implements IDevice.RST
        MyBase.RST()
    End Sub

    Public Overrides Sub CLS() Implements IDevice.CLS
        MyBase.CLS()
    End Sub

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        Visa.SendString("*RST")
        Visa.SendString("*CLS")

        InitChannels()

        SetChannels()

    End Sub

#End Region

#Region "Interface Methodes IScope"

    Public Overridable Sub SetChannel(Chan As CScopeChannel) Implements IScope.SetChannel

        SetChannel(Chan.ChanNr)

    End Sub

    Public Overridable Sub SetChannel(Nr As Integer) Implements IScope.SetChannel

        Dim _Chan As CScopeChannel = Me.Channel(Nr)

        If _Chan IsNot Nothing Then

            With _Chan

                If .State = CScopeChannel.ChanState.STATE_ON Then

                    Call Visa.SendString("SELect:" & .Name & " ON")

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
                            Call Visa.SendString(.Name & ":PRObe " & FormatNumber(.Probe, 2))
                        Case "A"
                            Call Visa.SendString(.Name & ":PRObe " & FormatNumber(1000 / .Probe, 2))
                    End Select

                    Call Visa.SendString(.Name & ":VOLts " & FormatNumber(.VertVolt))
                    Call Visa.SendString(.Name & ":OFFSet " & FormatNumber(.Offset))

                Else
                    Call Visa.SendString("SELect:" & .Name & " OFF")

                End If

            End With

        End If

    End Sub

    Public Overridable Sub SetChannels() Implements IScope.SetChannels

        For Each _chan As CScopeChannel In _Channels

            SetChannel(_chan)

        Next

    End Sub

    Public Overridable Sub InitChannel(Nr As Integer) Implements IScope.InitChannel

        Me.Channel(Nr).InitChannel(Nr)

    End Sub

    Public Overridable Sub InitChannel(Chan As CScopeChannel) Implements IScope.InitChannel

        InitChannel(Chan.ChanNr)

    End Sub


    '''Initialization the scope channels 

    Public Overridable Sub InitChannels() Implements IScope.InitChannels

        'Call Visa.SendString("*RST")
        'Call cHelper.Delay(1)

        Call Visa.SendString("CLEARMenu")
        Call Visa.SendString("Cursor:function OFF")

        For Each _chan As CScopeChannel In _Channels

            InitChannel(_chan)

        Next

    End Sub

    Public Overridable Sub CaptureScreen2File() Implements IScope.CaptureScreen2File
        Throw New NotImplementedException()
    End Sub


    Public Overridable Sub SetHorizontal() Implements IScope.SetHorizontal
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub SetTrigger() Implements IScope.SetTrigger
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub Acquire(acqState As CScopeTrigger.Acquire) Implements IScope.Acquire
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub ClearScreen() Implements IScope.ClearScreen
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub LoadReferenceCurve(sFileName As String, nRef As Integer) Implements IScope.LoadReferenceCurve
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub RefCurveOn(refNr As Integer) Implements IScope.RefCurveOn
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub RefCurveOff(refNr As Integer) Implements IScope.RefCurveOff
        Throw New NotImplementedException()
    End Sub

    Public Overridable Function MeasDelay(MeasNr As Integer, Source1 As String, slope1 As Integer, Source2 As String, slope2 As Integer) As Single Implements IScope.MeasDelay
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasEdge(MeasNr As Integer, SOURCE As String, lowRefLevel As Integer, highRefLevel As Integer, slope As Integer) As Single Implements IScope.MeasEdge
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasFreq(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasFreq
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasPK2PK(MeasNr As Integer, Source1 As String) As Single Implements IScope.MeasPK2PK
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasRMS(MeasNr As Integer, sSource1 As String, Optional waitTime As Integer = 1) As Single Implements IScope.MeasRMS
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasPOVERSHOOT(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasPOVERSHOOT
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasNOVERSHOOT(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasNOVERSHOOT
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasIMAX(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasIMAX
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasIMIN(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasIMIN
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasHIGH(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasHIGH
        Throw New NotImplementedException()
    End Function

    Public Overridable Function MeasLOW(MeasNr As Integer, sSource1 As String) As Single Implements IScope.MeasLOW
        Throw New NotImplementedException()
    End Function

    Public Overridable Sub SetColorScheme(colScheme As String) Implements IScope.SetColorScheme
        Throw New NotImplementedException()
    End Sub
#End Region


#Region "Help functions"

    Private Sub CreateChannels()
        'Create List of 4 channels with default settings
        Dim myChan As CScopeChannel
        For i As Integer = 0 To _CountOfChannels - 1

            myChan = New CScopeChannel(_ErrorLogger, i + 1)

            _Channels.Insert(i, myChan)

        Next
    End Sub

    Public Function GetTriggerMode(nTriggerMode As Integer) As String
        Dim sRet As String = String.Empty

        Select Case nTriggerMode
            Case CScopeTrigger.TriggerMode.AUTO
                sRet = "auto"
            Case CScopeTrigger.TriggerMode.NORMAL
                sRet = "normal"
            Case CScopeTrigger.TriggerMode.SIGNL
                sRet = "signl"
        End Select

        Return sRet

    End Function

    Public Function GetTriggerCoupling(nCoupling As Integer) As String
        Dim sRet As String = String.Empty

        Select Case nCoupling
            Case CScopeTrigger.TriggerCoupling.AC
                sRet = "AC"
            Case CScopeTrigger.TriggerCoupling.DC
                sRet = "DC"
            Case CScopeTrigger.TriggerCoupling.AC_HFReject, CScopeTrigger.TriggerCoupling.DC_HFReject
                sRet = "HFRej"
            Case CScopeTrigger.TriggerCoupling.LFReject
                sRet = "LFRej"
            Case CScopeTrigger.TriggerCoupling.AC_NREJect, CScopeTrigger.TriggerCoupling.DC_NREJect
                sRet = "NOISErej"
        End Select

        Return sRet

    End Function

    Overridable Function GetSlope(nSlope As Integer) As String

        Throw New NotImplementedException

    End Function

    Overridable Function MeasSource(MeasNr As Integer, Source1 As String, cmd As String) As Single

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":SOURCE " & Source1)
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":TYPe " & UCase(cmd))
        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":STATE on")

        Call cHelper.Delay(1)

        Return MeasDecimalValue(MeasNr)

    End Function

    Overridable Function MeasDecimalValue(MeasNr As Integer) As Decimal

        Call Visa.SendString("MEASUrement:MEAS" & MeasNr & ":VALue?")

        Return CDec(Visa.ReceiveValue)

    End Function

#End Region


End Class
