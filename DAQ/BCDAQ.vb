Imports Ivi.Visa

Public Class BCDAQ
    Inherits BCDevice
    Implements IDAQ

#Region "Shorthand Properties"

    Public Property VoltageMax As Single Implements IDAQ.VoltageMax
    Public Property CurrentMax As Single Implements IDAQ.CurrentMax
    Public Property ScanList As List(Of CDAQChannel) Implements IDAQ.ScanList

    Public ReadOnly Property FunctionList As SortedList(Of String, String)
    Overridable Property ModuleFacor As Integer

    Public Property MeasModul As Integer

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

        _FunctionList = CreateFunctionList()

        _ScanList = New List(Of CDAQChannel)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Function IDN() As String Implements IDevice.IDN
        MyBase.IDN()
    End Function

    Public Overrides Sub RST() Implements IDevice.RST
        MyBase.RST()
    End Sub

    Public Overrides Sub CLS() Implements IDevice.CLS
        MyBase.CLS()
    End Sub

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        RST()
        CLS()

    End Sub
#End Region

#Region "Interface Methodes IDAQ"

    Public Overridable Function MeasChannel(Chan As CDAQChannel, Optional Sample As Integer = 1) As Single Implements IDAQ.MeasChannel

        Throw New NotImplementedException

    End Function


    Public Overridable Function QueryNumericItems() As Double() Implements IDAQ.QueryNumericItems

        If _ScanList.Count > 0 AndAlso GetRouteScanSize() <> _ScanList.Count Then

            ConfigureChannels()
            RouteChannels()

        End If

        Call Visa.SendString("READ?")

        'cHelper.Delay(2)

        Return Visa.ReceiveValueList(",")

    End Function

    Public Sub RouteClose(ByVal nModul As Integer, ByVal Chan As Integer) Implements IDAQ.RouteClose
        Call SetRelais(nModul, Chan, 1)
    End Sub

    Public Sub RouteOPEN(ByVal nModul As Integer, ByVal Chan As Integer) Implements IDAQ.RouteOpen
        Call SetRelais(nModul, Chan, 0)
    End Sub

    Public Overridable Sub ConfigureChannels()

        For Each Chan As CDAQChannel In _ScanList

            ConfigureChannel(Chan)

        Next

    End Sub



    Public Overridable Sub RouteChannel(Chan As CDAQChannel)

        Dim Cmd As String = "ROUTE:SCAN "
        Cmd &= "(@" & CStr(Chan.Modul * ModuleFacor + Chan.Nr) & ")"
        Call Visa.SendString(Cmd)

    End Sub


    Public Overridable Sub RouteChannels()

        Dim CmdString As String = "ROUTE:SCAN "

        CmdString &= GetScanList()

        Call Visa.SendString(CmdString)

    End Sub




#End Region


#Region "Help functions"

    Overridable Sub SetRelais(ByVal nModul As Integer, ByVal Chan As Integer, State As Boolean)
        Dim CmdString As String

        If State Then CmdString = "ROUT:CLOS " Else CmdString = "ROUT:OPEN "

        CmdString = CmdString & "(@" & (nModul * ModuleFacor + Chan).ToString & ")"

        Visa.SendString(CmdString)


    End Sub

    Overridable Sub ConfigureChannel(Chan As CDAQChannel)

        Dim Cmd As String = "CONF:"
        Cmd &= GetFunction(Chan.DAQFn) & " "

        If Chan.Range = 0 Then
            Cmd &= "AUTO"
        Else
            Cmd &= Chan.Range.ToString
        End If

        Cmd &= ", "
        Cmd &= "(@" & CStr(Chan.Modul * ModuleFacor + Chan.Nr) & ")"

        Call Visa.SendString(Cmd)

    End Sub


    Overridable Function CreateFunctionList() As SortedList(Of String, String)
        Dim fsl As New SortedList(Of String, String) From {
            {IDAQ.DAQ_Function.VOLT_DC.ToString, "VOLT:DC"},
            {IDAQ.DAQ_Function.VOLT_AC.ToString, "VOLT:AC"},
            {IDAQ.DAQ_Function.CURRENT_DC.ToString, "CURR:DC"},
            {IDAQ.DAQ_Function.CURRENT_AC.ToString, "CURR:AC"},
            {IDAQ.DAQ_Function.RES.ToString, "RES"},
            {IDAQ.DAQ_Function.FRES.ToString, "FRES"},
            {IDAQ.DAQ_Function.FREQ.ToString, "FREQ"}
        }

        Return fsl

    End Function


    Overridable Function GetFunction(nFn As IDAQ.DAQ_Function) As String

        Dim sRet As String = String.Empty

        If FunctionList.ContainsKey(nFn.ToString) Then

            sRet = FunctionList.Item(nFn.ToString).ToString

        End If

        Return sRet
    End Function

    Overridable Function GetFunctionIndex(nFn As IDAQ.DAQ_Function, nElm As Integer) As Integer

        Return CInt(nFn)

    End Function

    Overridable Function GetChannelIndex(ByVal Chan As CDAQChannel) As Integer

        Dim nRet As Integer = Integer.MinValue

        For i As Integer = 0 To _ScanList.Count - 1

            If CInt(_ScanList(i).Nr) = Chan.Nr Then

                nRet = i
                Exit For

            End If

        Next

        Return nRet

    End Function

    Overridable Function GetScanChanne() As String

    End Function


    Overridable Function GetScanList() As String

        Dim sRet As String = "(@"

        If _ScanList IsNot Nothing Then

            For i As Integer = 0 To _ScanList.Count - 1

                Dim _myChan As CDAQChannel = _ScanList(i)

                sRet &= (_myChan.Modul * ModuleFacor + _myChan.Nr).ToString
                If (_ScanList.Count > 1) And (i < _ScanList.Count - 1) Then
                    'more items
                    sRet &= ","
                End If

            Next
        End If
        Return sRet & ")"
    End Function

    Private Sub AddChane2ScanList(Chan As CDAQChannel)

        If Not IsInScanList(Chan) AndAlso Chan IsNot Nothing Then
            _ScanList.Add(Chan)
        End If

    End Sub


    Private Function IsInScanList(Chan As CDAQChannel) As Boolean

        Dim bRet As Boolean = False

        For Each itm As CDAQChannel In _ScanList

            If itm.Nr = Chan.Nr Then
                bRet = True
            End If

        Next

        Return bRet

    End Function

    Overridable Function GetRouteScanSize() As Integer

        Visa.SendString("ROUTE:SCAN:SIZE?")

        Return CInt(Visa.ReceiveValue)

    End Function

    Private Function CompareScanListASC(x As Integer, y As Integer) As Integer
        Return y.CompareTo(x)
    End Function


#End Region


End Class
