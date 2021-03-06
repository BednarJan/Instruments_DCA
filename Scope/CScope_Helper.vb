﻿Public Class CScope_Helper

#Region "Enums"
    Public Enum EChannel
    CH1 = 1
    CH2 = 2
    CH3 = 3
    CH4 = 4
    CH5 = 5
    CH6 = 6
    CH7 = 7
    CH8 = 8
End Enum

Public Enum EMeasParam
    AMPLitude
    AVERage
    AVGFreq
    AVGPeriod
    BWIDth
    DELay
    DT
    DUTYcycle
    ENUMber
    FALL
    FREQuency
    HIGH
    LOW
    MAXimum
    MINimum
    NOVershoot
    NWIDth
    PERiod
    PNUMber
    POVershoot
    PTOPeak
    PWIDth
    RISE
    RMS
    SDEViation
    TY1Integ
    TY2Integ
    V1
    V2
End Enum

Public Enum EArea
    A1 = 1
    A2 = 2
End Enum

#End Region

    Public Shared Function GetUniquePictureFileName(ByVal TestResultsPath As String,
                                             TestName As String,
                                             fileExt As String) As String

        Dim sRet As String = TestResultsPath & "\ScopePlots\"

        If Not System.IO.Directory.Exists(sRet) Then
            System.IO.Directory.CreateDirectory(sRet)
        End If

        sRet &= TestName
        sRet &= Format(Now(), "ddMMyyyyhhmmss")
        sRet &= "." & fileExt

        Return sRet

    End Function

End Class
