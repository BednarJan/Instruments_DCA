Imports Instruments

Public Class CPWAN_Yokogawa_WT200_DCA
    Inherits BCVisaDev
    Implements IPWAN

    Private Enum eDisplay
        a = 1
        B = 2
        c = 3
    End Enum

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public ReadOnly Property VoltageMax As Single = 300 Implements IPWAN.VoltageMax
    Public ReadOnly Property CurrentMax As Single = 15 Implements IPWAN.CurrentMax
#End Region

#Region "Constructor"
    Sub New(strVisa_Adr As String, oErrorLogger As CErrorLogger)
        MyBase.New(strVisa_Adr, oErrorLogger)
        _ErrorLogger = oErrorLogger
        _strVisa_Adr = strVisa_Adr

    End Sub
#End Region

#Region "Interface IPWAN Methodes"
    Function Initialize() As Integer Implements IPWAN.Initialize
        Dim nOK As Integer = 0
        Try
            RST()
            CLS()
        Catch ex As Exception
            nOK += 1
            Throw ex
        End Try
        Return nOK
    End Function

    Public Overrides Sub CLS() Implements IPWAN.CLS
        Try
            MyBase.CLS()
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try
    End Sub

    Public Overrides Sub RST() Implements IPWAN.RST
        Try
            MyBase.RST()
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try

    End Sub

    Public Overrides Function IDN() As String Implements IPWAN.IDN
        Dim sRetVal As String = String.Empty
        Try
            sRetVal = MyBase.IDN()
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try
        Return sRetVal
    End Function

    Public Function PowerFactor(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.PowerFactor
        Throw New NotImplementedException()
    End Function

    Public Function GetVrms(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetVrms
        Throw New NotImplementedException()
    End Function

    Public Function GetVthd(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetVthd
        Throw New NotImplementedException()
    End Function

    Public Function GetIrms(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetIrms
        Throw New NotImplementedException()
    End Function

    Public Function GetIthd(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetIthd
        Throw New NotImplementedException()
    End Function

    Public Function GetFreq(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetFreq
        Throw New NotImplementedException()
    End Function

    Public Sub GetHarmonics(clss As Integer, ByRef HarmonicMeas() As Single, ByRef HarmonicLimit() As Single, ByRef ActivePower As Single) Implements IPWAN.GetHarmonics
        Throw New NotImplementedException()
    End Sub

    Public Sub Init_Integration(durationInSec As Single) Implements IPWAN.Init_Integration

        Dim hh As Integer
        Dim mm As Integer
        Dim ss As Integer

        Try

            CHelper.Seconds_2_hh_mm_ss(durationInSec, hh, mm, ss)

            MyBase.Send("DISPLAY" & CStr(eDisplay.a) & ":FUNCtion " & "Time")
            MyBase.Send("DISPLAY" & CStr(eDisplay.B) & ":FUNCtion " & "A")
            MyBase.Send("DISPLAY" & CStr(eDisplay.c) & ":FUNCtion " & "Wh")

            MyBase.Send("INTEGrate:RESet")
            MyBase.Send("INTEGrate:MODE NORMAL")
            MyBase.Send("INTEGRATE:TYPE STANdard")

            MyBase.Send("INTEGrate:TIMer " & hh.ToString.PadLeft(2, "0"c) & "," & mm.ToString.PadLeft(2, "0"c) & "," & ss.ToString.PadLeft(2, "0"c))

        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try

    End Sub

    Public Sub Set_Volt_Range(sRange As String) Implements IPWAN.Set_Volt_Range
        Throw New NotImplementedException()
    End Sub

    Public Sub Set_Cur_Range(sRange As String) Implements IPWAN.Set_Cur_Range
        Throw New NotImplementedException()
    End Sub

#End Region

#Region "Public Special Functions WT200"
    Public Function IntegratePower(durationInSec As Single) As Double Implements IPWAN.IntegratePower
        Dim dRet As Double
        Dim measData() As Double = Nothing
        Try
            Init_Integration(durationInSec)

            WaitTime(2000)

            MyBase.Send("INTEGrate:STARt")

            WaitTime(durationInSec * 1000 + 5)

            Get_Integrations_Data(measData)

            WaitTime(2000)

            dRet = 0
            If durationInSec > 0 Then
                dRet = (measData(1) * 3600) / durationInSec
            End If

            IntegratePower = dRet
        Catch ex As Exception
            _ErrorLogger.LogInfo(String.Format("Error in {0} function, {1}, VISA-string = {2}", System.Reflection.MethodBase.GetCurrentMethod.Name(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name, CStr(_strVisa_Adr)))
            Throw ex
        End Try

    End Function

    Private Sub Get_Integrations_Data(ByRef measData() As Double) Implements IPWAN.Get_Integrations_Data
        'measData(0) ........... (P)ower  - active
        'measData(1) ........... Energy in Wh
        'measData(2) ........... integrations Time (h)
        'measData(3) ........... integrations Time (hr)
        'measData(4) ........... integrations Time (min)
        'measData(5) ........... integrations Time (sec)
        Dim Buffer As String = String.Empty
        Dim measVals() As String
        Dim iii As Integer


        Call MyBase.Send("MEASure:NORMal:Item:PRESet INTEGrate")
        Call MyBase.Send("Measure:Value?")
        Buffer = MyBase.ReceiveString
        measVals = Split(Buffer, ",")

        ReDim measData(UBound(measVals))
        For iii = 0 To UBound(measVals)
            If Not IsNumeric(measVals(iii)) Then
                measVals(iii) = Replace(measVals(iii), ".", ",")
            End If

            measData(iii) = CDbl(measVals(iii))
        Next iii

    End Sub
#End Region

End Class
