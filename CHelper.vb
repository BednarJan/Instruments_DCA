Imports System.Globalization
Imports System.Reflection

Public Class cHelper

    Public Shared Function GetEquipmentClasses() As ArrayList
        Dim alRet As ArrayList = New ArrayList

        'Dim myType As Type = GetType(BCVisaDev)
        'Dim assem As Assembly = Assembly.GetAssembly(myType)
        'Dim classes() As Type = assem.GetExportedTypes()

        'For i As Integer = 0 To classes.Length - 1
        '    If classes(i).IsClass AndAlso (classes(i).Name.Contains("cLoad") OrElse classes(i).Name.Contains("cDAQ") OrElse classes(i).Name.Contains("cPSU") OrElse classes(i).Name.Contains("cScope")) Then
        '        alRet.Add(classes(i))
        '    End If
        'Next
        Return alRet
    End Function

    Public Shared Function StringToSingle(ByVal str) As Single
        Dim Separator As String = String.Empty

        Separator = CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator

        str = str.Replace("."c, Separator).Replace(","c, Separator)

        If IsNumeric(str) Then
            Return Convert.ToSingle(str)
        Else
            Return Single.NaN
        End If

    End Function

    Public Shared Sub Seconds_2_hh_mm_ss(ByVal iSecond As Long, ByRef hh As Integer, ByRef mm As Integer, ByRef ss As Integer)
        Dim iSpan As TimeSpan = TimeSpan.FromSeconds(iSecond)
        hh = iSpan.Hours
        mm = iSpan.Minutes
        ss = iSpan.Seconds
    End Sub

#Region "Timing Functions"
    Public Shared Sub Delay(ByVal Seconds As Single)
        System.Threading.Thread.Sleep(Seconds * 1000)
    End Sub
#End Region

End Class
