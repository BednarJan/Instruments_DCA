Imports System.IO
Imports System.Configuration
Imports System.Diagnostics
Imports System.Reflection

Public Class CErrorLogger
    Private _LogFileName As String = String.Empty
    Private _LastError As String

    Public Property TraceLevels As UShort = 2

#Region "Constructor"
    Public Sub New(sFileName As String)

        _LogFileName = sFileName

    End Sub
#End Region

#Region "Private Methodes"

    Private Function CheckAndCreateIfLogFileNotExists(ByVal myFileName As String) As FileStream

        Dim fs As FileStream

        'Check for existence of logger file    
        If Not Directory.Exists(Path.GetDirectoryName(myFileName)) Then
            Directory.CreateDirectory(Path.GetDirectoryName(myFileName))
        End If

        If Not File.Exists(myFileName) Then
            fs = File.Create(myFileName)
        Else
            fs = New FileStream(myFileName, FileMode.Append, FileAccess.Write)
        End If

        Return fs

    End Function


    Private Sub Info(ByVal info As String)

        'Check for existence of logger file    
        Try

            Dim fs As FileStream = CheckAndCreateIfLogFileNotExists(_LogFileName)
            Dim sw As StreamWriter = New StreamWriter(fs)

            sw.WriteLine(vbCrLf & "--- " + vbCrLf + DateTime.Now + " " + info.ToString)
            sw.Close()
            fs.Close()

        Catch ex As Exception
            LogInfo(ex)
        End Try
    End Sub
#End Region

#Region "Public Methodes"
    Public Sub LogInfo(ByVal ex As Exception)

        Dim LogMessage As String = ""

        Try
            'Writes error information to the log file including name of the file, line number & error message description               
            Dim MainTrace As New StackTrace(True)

            For i As UShort = 1 To _TraceLevels
                Dim Level As UShort = MainTrace.FrameCount - 1 - i
                If Not i < 0 Then LogMessage &= MainTrace.GetFrame(Level).GetMethod.ReflectedType.Name & ": " & MainTrace.GetFrame(Level).GetMethod.Name & vbCrLf
            Next

            'Dim fileNames As String = trace.GetFrame((trace.FrameCount - 1)).GetFileName()
            Info(LogMessage & vbCrLf & "Error Message: " & ex.Message)
        Catch genEx As Exception
            Info(ex.Message)
        End Try
    End Sub

    Public Sub LogInfo(ByVal Message As String)
        Try
            'Write string message to the log file       
            Info("AVT Message: " + Message)
        Catch genEx As Exception
            Info(genEx.Message)
        End Try
    End Sub

    Public Sub LogException(ByVal ex As Exception, _strVisa_Adr As String)
        Dim sLogStr As String
        Dim stackTrace As StackTrace = New StackTrace()
        Dim fullMethodName As String
        Dim myMethod As Reflection.MethodBase
        Try
            myMethod = stackTrace.GetFrame(1).GetMethod()
            fullMethodName = myMethod.ReflectedType.FullName & "." & myMethod.Name

            sLogStr = String.Format("Error in {0} function, VISA-Addr = '{1}'", fullMethodName, _strVisa_Adr)
            sLogStr &= vbCrLf
            sLogStr &= ex.Message
            LogInfo(sLogStr)
        Catch genEx As Exception
            Info(ex.Message)
        End Try
    End Sub
#End Region

End Class

