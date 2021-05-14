Imports Ivi.Visa

Public Class BCDAQ
    Inherits BCDevice
    Implements IDAQ

#Region "Shorthand Properties"
    Public Property VoltageMax As Single Implements IDAQ.VoltageMax
    Public Property CurrentMax As Single Implements IDAQ.CurrentMax
    Public Property ScanList As CScanList Implements IDAQ.ScanList

    Public ReadOnly Property FunctionList As SortedList

    Public Property MeasModul As Integer

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

        _FunctionList = CreateFunctionList()

        _ScanList = New CScanList

        _MeasModul = 1

    End Sub
#End Region


#Region "Interface Methodes IDAQ"

    Public Function Get_Volt_DC(Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single Implements IDAQ.Get_Volt_DC

        Return MeasChannel(Chan, IDAQ.DAQ_Function.VOLT_DC, sRange, _MeasModul)

    End Function

    Public Function Get_Volt_AC(Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single Implements IDAQ.Get_Volt_AC

        Return MeasChannel(Chan, IDAQ.DAQ_Function.VOLT_AC, sRange, _MeasModul)

    End Function

    Public Function Get_Curr_DC(Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single Implements IDAQ.Get_Curr_DC

        Return MeasChannel(Chan, IDAQ.DAQ_Function.CURRENT_DC, sRange, _MeasModul)

    End Function

    Public Function Get_Curr_AC(Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single Implements IDAQ.Get_Curr_AC

        Return MeasChannel(Chan, IDAQ.DAQ_Function.CURRENT_AC, sRange, _MeasModul)

    End Function

    Public Function Get_Res(Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single Implements IDAQ.Get_Res

        Return MeasChannel(Chan, IDAQ.DAQ_Function.RES, sRange, _MeasModul)

    End Function

    Public Function Get_FRes(Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single Implements IDAQ.Get_FRes

        Return MeasChannel(Chan, IDAQ.DAQ_Function.FRES, sRange, _MeasModul)

    End Function

    Public Function Get_FReq(Chan As Integer, Optional ByVal sRange As String = "AUTO") As Single Implements IDAQ.Get_FReq

        Return MeasChannel(Chan, IDAQ.DAQ_Function.FREQ, sRange, _MeasModul)

    End Function

    Public Function Get_Sample_Volt_DC(Chan As Integer, Sample As Integer) As Single Implements IDAQ.Get_Sample_Volt_DC

        Return GetDAQSample(Chan, IDAQ.DAQ_Function.VOLT_DC, Sample)

    End Function


#End Region


#Region "Help functions"

    Private Function QueryValueList(ByVal cmdStr As String) As Single()

        Throw New NotImplementedException

    End Function


    Overridable Function CreateFunctionList() As SortedList
        Dim fsl As New SortedList From {
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

        Dim myKeys As ICollection = FunctionList.Keys

        For i As Integer = 0 To FunctionList.Count - 1
            If FunctionList.ContainsKey(nFn.ToString) Then
                sRet = FunctionList.GetByIndex(i).ToString
            End If
        Next

        Return sRet
    End Function

    Overridable Function GetFunctionIndex(nFn As IDAQ.DAQ_Function, nElm As Integer) As Integer

        Dim nRet As Integer = 0

        For i As Integer = 0 To FunctionList.Count - 1

            Dim fn As KeyValuePair(Of String, String) = FunctionList(i)
            If fn.Key = nFn.ToString Then
                nRet = (nElm - 1) * FunctionList.Count + i + 1
            End If

        Next

        Return nRet

    End Function

    Private Function MeasChannel(Chan As Integer, nFn As IDAQ.DAQ_Function, sRange As String, Optional nModul As Integer = 1) As Decimal

        Dim CmdString As String
        Dim channel As Integer = Chan + nModul * 1000

        CmdString = "CONF:" & GetFunction(nFn) & " " & sRange & ","
        CmdString &= GetScanList(Chan)

        Call Visa.SendString(CmdString)
        cHelper.Delay(1)

        Return MeasDecimalValue()

    End Function

    Overridable Function MeasDecimalValue() As Decimal

        Call Visa.SendString("READ?")

        Return CDec(Visa.ReceiveValue)

    End Function

    Public Function GetDAQSample(ByVal Chan As Integer, nFn As IDAQ.DAQ_Function, ByVal Sample As Integer) As Single

        Dim ChList As String = GetScanList(Chan)

        Dim CmdString As String = "CONF:"
        CmdString &= GetFunction(nFn)
        CmdString &= "AUTO ,DEF, "
        CmdString &= ChList
        Visa.SendString(CmdString)

        Visa.SendString("ROUTe:SCAN (@)")
        Visa.SendString("ROUTe:SCAN " & ChList)
        Visa.SendString("SAMPle:COUNt " & Sample)
        Visa.SendString("INITiate")
        cHelper.Delay(1)
        Visa.SendString("CALC:AVER:AVER? " & ChList)

        Return Visa.ReceiveValue

    End Function

    Overridable Function GetScanList(Optional nModul As Integer = 1) As String

        Const sStartStr As String = "(@"
        Const sEndStr As String = ")"

        Dim nDAQChan As Integer = 1000 * nModul
        Dim sRet As String = sStartStr

        If _ScanList IsNot Nothing Then

            _ScanList.Sort(AddressOf CompareScanListASC)

            For i As Integer = 0 To _ScanList.Count - 1

                sRet &= (CInt(_ScanList(i)) + nDAQChan).ToString
                If (_ScanList.Count > 1) And (i < _ScanList.Count) Then
                    'more items
                    sRet &= ","
                End If

            Next
        End If
        Return sRet & sEndStr
    End Function

    Overridable Function GetScanList(Chan As Integer, Optional nModul As Integer = 1) As String

        _ScanList.Add(Chan)

        Return GetScanList(nModul)

    End Function

    Private Function CompareScanListASC(x As Integer, y As Integer) As Integer
        Return y.CompareTo(x)
    End Function


#End Region


End Class
