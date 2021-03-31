Imports Ivi.Visa
Public Class BCScope
    Inherits BCDevice
    Implements IScope


#Region "Shorthand Properties"

    Property CountOfChannels As Integer
    Property HardcopyFullFileName As String Implements IScope.HardcopyFullFileName
    Property Channels As List(Of CScopeChannel) Implements IScope.Channels
    Property TimeBase As Single Implements IScope.TimeBase
    Property Trigger As CScopeTrigger Implements IScope.Trigger
    Property HardcopyFileFormat As UInteger Implements IScope.HardcopyFileFormat

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger, nChanCount As Integer)

        MyBase.New(Session, ErrorLogger)

        Dim myChan As CScopeChannel

        _CountOfChannels = nChanCount

        'Default settings

        'Create List of 4 channels with default settings
        For i As Integer = 1 To CountOfChannels

            myChan = New CScopeChannel(ErrorLogger) With {
            .Name = "CH" & CStr(i),
            .State = CScopeChannel.ChanState.STATE_OFF,
            .Bandwidth = CScopeChannel.ChanBandwidth.B20,
            .Coupling = CScopeChannel.ChanCoupling.DC,
            .Polarity = CScopeChannel.ChanPolarity.NORMAL,
            .Position = 0,
            .Offset = 0
            }

            _Channels.Insert(i, myChan)
        Next

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

        Visa.SendString("*RST;*CLS" & Chr(10))

    End Sub
#End Region

#Region "Interface Methodes IScope"




#End Region


End Class
