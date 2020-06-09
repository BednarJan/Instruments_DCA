
Public Class CPSU_SorensenSGI20X500D
    Inherits BCVisaDev
    Implements ISource_DC

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Private ReadOnly Property VoltageMax As Single = 20 Implements ISource_DC.VoltageMax
    Private ReadOnly Property CurrentMax As Single = 500 Implements ISource_DC.CurrentMax
#End Region

#Region "Constructor"
    Public Sub New(strVisa_Adr As String, oErrorLogger As CErrorLogger)
        MyBase.New(strVisa_Adr, oErrorLogger)
        _ErrorLogger = oErrorLogger
        _strVisa_Adr = strVisa_Adr
    End Sub
#End Region

#Region "Interface Methodes IPSU_DC"

    Public Function Initialize() As Integer Implements ISource_DC.Initialize
        Dim nOK As Integer = 0
        Try
            MyBase.Send("SYST:LANG TMSL")
            MyBase.Send("*RST;*CLS:")
            SetOFF()
        Catch ex As Exception
            nOK += 1
        End Try
        Return nOK
    End Function


    Public Sub SetVolt(nVolt As Single) Implements ISource_DC.SetVolt
        Call SetSource(nVolt, _CurrentMax)
    End Sub

    Public Sub SetCurLim(CurLim As Single) Implements ISource_DC.SetCurLim
        Try
            Dim sCmd As String = String.Empty
            If CurLim > _CurrentMax Then CurLim = _CurrentMax
            sCmd = "SOURCE:CURRENT " & CurLim.ToString
            MyBase.send(sCmd)
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try
    End Sub

    Public Sub SetSource(Volt As Single, CurLim As Single, Optional SetOutON As Boolean = True) Implements ISource_DC.SetSource
        Try
            If Not SetOutON Then Call SetOFF()
            Dim sCmd As String = String.Empty
            If Volt > _VoltageMax Then Volt = _VoltageMax
            If CurLim > _CurrentMax Then CurLim = _CurrentMax
            sCmd = "SOURCE:CURRENT " & CurLim.ToString
            sCmd &= ";"
            sCmd &= "VOLT " & Volt.ToString
            sCmd &= ";"
            MyBase.Send(sCmd)
                If SetOutON Then Call SetON()

        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try
    End Sub

    Public Function GetVolt() As Single Implements ISource_DC.GetVolt
        Dim sRetval As Single = Single.MinValue
        Dim sHlp As String = String.Empty
        Try
            MyBase.Send("MEAS:VOLT ? ")
            sHlp = ReceiveString()
            If IsNumeric(sHlp) Then
                sRetval = CSng(sHlp)
            End If
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try
        Return sRetval
    End Function


    Public Sub SetOFF() Implements ISource_DC.SetOFF
        Try
            MyBase.Send("OUTPUT:STATE OFF")
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try
    End Sub

    Public Sub SetON() Implements ISource_DC.SetON
        Try
            MyBase.Send("OUTPUT:STATE ON")
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try
    End Sub

    Public Overridable Sub iPSU_DC_CLS() Implements ISource_DC.CLS
        Try
            MyBase.CLS()
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try

    End Sub

    Public Overridable Sub iPSU_DC_RST() Implements ISource_DC.RST
        Try
            MyBase.RST()
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try

    End Sub

    Public Overridable Function iPSU_DC_IDN() As String Implements ISource_DC.IDN
        Dim sRetVal As String = String.Empty
        Try
            sRetVal = MyBase.IDN()
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try
        Return sRetVal
    End Function
#End Region

End Class
