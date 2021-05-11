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

        Try
            Return Single.Parse(str, CultureInfo.InvariantCulture)
        Catch ex As Exception
            Return Single.MinValue
        End Try

    End Function

    Public Shared Function StringToDouble(ByVal str) As Double
        Dim Separator As String = String.Empty

        Separator = CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator

        str = str.Replace("."c, Separator).Replace(","c, Separator)

        Try
            Return Double.Parse(str, CultureInfo.InvariantCulture)
        Catch ex As Exception
            Return Double.MinValue
        End Try

    End Function

    Public Shared Function StringToDecimal(ByVal str) As Decimal

        Return CDec(StringToDouble(str))

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

    Public Shared Function HexToBytes(ByVal HexString As String) As Byte()
        Dim retBytes As Byte() = BitConverter.GetBytes(Long.Parse(HexString, NumberStyles.AllowHexSpecifier))
        Return retBytes
    End Function

    Public Shared Function String2ByteArraySpecial(ByVal myString As String) As Byte()
        Dim iLen As Integer
        Dim lIndx As Integer
        Dim retBytes As Byte() = New Byte() {}

        If myString <> vbNullString Then

            iLen = Len(myString) - 1

            For lIndx = 0 To iLen
                retBytes(lIndx) = Asc(Mid$(myString, lIndx + 1, 1))
            Next lIndx

        End If

        Return retBytes
    End Function


End Class
