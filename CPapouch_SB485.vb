Imports Ivi.Visa
Imports System.Math
Public Class CPapouch_SB485
     Inherits BCDevice
    Implements IDevice

#Region "Shorthand Properties"

#End Region

#Region "Constructor"

    Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger)

    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"
    Public Overrides Function IDN() As String Implements IDevice.IDN
        Return MyBase.Name
    End Function

    Public Overrides Sub RST() Implements IDevice.RST
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub CLS() Implements IDevice.CLS
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub Initialize() Implements IDevice.Initialize

    End Sub

#End Region

#Region "RS 485 Functions"

    Public Sub RS485_Write(ByRef TxData() As Byte)

        Dim TxBuffer As String
        Dim TxLength As Integer
        Dim i As Integer

        Try
            TxBuffer = ""
            TxLength = UBound(TxData) + 1

            For i = 0 To TxLength - 1
                TxBuffer &= Chr(TxData(i))
            Next i

            Call MyBase.SendString(TxBuffer)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub RS485_Read(ByRef RxData() As Byte,
                            Optional ByVal bWithHex2ByteConversion As Boolean = True)

        Dim sRet As String

        Try
            sRet = MyBase.ReceiveString()

            If bWithHex2ByteConversion Then

                RxData = cHelper.HexToBytes(sRet)

            Else

                RxData = cHelper.String2ByteArraySpecial(sRet)

            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    Public Function RS485_WriteAndRead(TxData() As Byte) As Byte()

        Dim RxData As Byte() = New Byte() {}

        RS485_Write(TxData)
        cHelper.Delay(0.2)
        RS485_Read(RxData, False)

        Return RxData

    End Function


#End Region


End Class
