Imports Ivi.Visa
Imports Ivi.Visa.FormattedIO

Public Class BCScope
    Inherits BCVisaDev
    Implements IScope

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty
    Private _strFile2SavePicture As String = String.Empty
    'Private _FormattedIO As FormattedIO.MessageBasedFormattedIO = Nothing

    Public ReadOnly Property VoltageMax As Single Implements IScope.VoltageMax
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Property File2SavePicture As String Implements IScope.File2SavePicture
        Get
            Return _strFile2SavePicture
        End Get
        Set(value As String)
            _strFile2SavePicture = value
        End Set
    End Property

    Public Sub New(strVisa_Adr As String, oErrorLogger As CErrorLogger)
        MyBase.New(strVisa_Adr, oErrorLogger)
        _ErrorLogger = oErrorLogger
        _strVisa_Adr = strVisa_Adr
    End Sub

    Private Sub iScope_RST() Implements IScope.RST
        MyBase.RST()
    End Sub

    Private Sub iScope_CLS() Implements IScope.CLS
        MyBase.CLS()
    End Sub

    Private Function iScope_IDN() As String Implements IScope.IDN
        Return MyBase.IDN()
    End Function

    'Public Function Initialize() As Integer Implements IScope.Initialize
    '    Dim nOK As Integer = 0
    '    Try
    '        MyBase.RST()
    '        MyBase.CLS()
    '    Catch ex As Exception
    '        nOK += 1
    '    End Try
    '    Return nOK
    'End Function


    Public Overridable Function Meas_Delay(MeasNr As Integer, Source1 As Integer, edge1 As String, Source2 As Integer, Edge2 As String) As Single Implements IScope.Meas_Delay
        Dim sRet As Single = Single.MinValue
        Try
            MyBase.Send("MEASUrement:METHod MINMax")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":TYPe DELAY")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE1 CH" & Source1)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE2 CH" & Source2)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":DEL:EDGE1 " & edge1)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":DEL:EDGE2 " & Edge2)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":DEL:DIRECTION FORW")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":INDI:STATE " & MeasNr)
            MyBase.Send("MEASUrement:REFLevel:PERCent:MID 90")
            MyBase.Send("MEASUrement:REFLevel:PERCent:MID2 90")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":STATE on")
            MyBase.Send("MEASUrement:INDICators:State Meas" & MeasNr)

            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":VALue?")

            sRet = MyBase.ReceiveString
        Catch ex As Exception
            _ErrorLogger.LogException(ex, _strVisa_Adr)
        End Try
        Return sRet
    End Function

    Public Overridable Function Meas_Freq(MeasNr As Integer, Source As Integer) As Single Implements IScope.Meas_Freq
        Dim sRet As Single = Single.MinValue
        Try
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE CH" & Source)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":TYPe Frequency")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":STATE ON")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":VALue?")
            sRet = MyBase.ReceiveString
        Catch ex As Exception
            _ErrorLogger.LogException(ex, _strVisa_Adr)
        End Try
        Return sRet
    End Function

    Public Overridable Function Meas_Pk2Pk(MeasNr As Integer, Source As Integer) As Single Implements IScope.Meas_Pk2Pk
        Dim sRet As Single = Single.MinValue
        Try
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE CH" & Source)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":STATE on")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":TYPe PK2pk")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":VALue?")
            sRet = MyBase.ReceiveString
        Catch ex As Exception
            _ErrorLogger.LogException(ex, _strVisa_Adr)
        End Try
        Return sRet
    End Function

    Public Overridable Function Meas_RMS(MeasNr As Integer, Source As Integer) As Single Implements IScope.Meas_RMS
        Dim sRet As Single = Single.MinValue
        Try
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE CH" & Source)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":STATE on")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":TYPe RMS")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":VALue?")
            sRet = MyBase.ReceiveString
        Catch ex As Exception
            _ErrorLogger.LogException(ex, _strVisa_Adr)
        End Try
        Return sRet
    End Function

    Public Overridable Function Meas_Povershoot(MeasNr As Integer, Source As Integer) As Single Implements IScope.Meas_Povershoot
        Dim sRet As Single = Single.MinValue
        Try
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE CH" & Source)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":STATE on")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":TYPe POV")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":VALue?")
            sRet = MyBase.ReceiveString
        Catch ex As Exception
            _ErrorLogger.LogException(ex, _strVisa_Adr)
        End Try
        Return sRet
    End Function

    Public Overridable Function Meas_Novershoot(MeasNr As Integer, Source As Integer) As Single Implements IScope.Meas_Novershoot
        Dim sRet As Single = Single.MinValue
        Try
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE CH" & Source)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":STATE on")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":TYPe NOV")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":VALue?")
            sRet = MyBase.ReceiveString
        Catch ex As Exception
            _ErrorLogger.LogException(ex, _strVisa_Adr)
        End Try
        Return sRet
    End Function

    Public Overridable Function Meas_Imax(MeasNr As Integer, Source As Integer) As Single Implements IScope.Meas_Imax
        Dim sRet As Single = Single.MinValue
        Try
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE CH" & Source)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":STATE on")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":TYPe MAX")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":VALue?")
            sRet = MyBase.ReceiveString
        Catch ex As Exception
            _ErrorLogger.LogException(ex, _strVisa_Adr)
        End Try
        Return sRet
    End Function

    Public Overridable Function Meas_Imin(MeasNr As Integer, Source As Integer) As Single Implements IScope.Meas_Imin
        Dim sRet As Single = Single.MinValue
        Try
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":SOURCE CH" & Source)
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":STATE on")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":TYPe MINI")
            MyBase.Send("MEASUrement:MEAS" & MeasNr & ":VALue?")
            sRet = MyBase.ReceiveString
        Catch ex As Exception
            _ErrorLogger.LogException(ex, _strVisa_Adr)
        End Try
        Return sRet
    End Function

    Public Overridable Sub Print_Display_to_File() Implements IScope.Print_Display_to_File


    End Sub


End Class
