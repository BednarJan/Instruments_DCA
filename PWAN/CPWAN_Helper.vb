Public Class CPWAN_Helper

#Region "Enums"
    Public Enum EEquipmentClass
        ClassA = 1
        ClassD = 2
    End Enum

    'Only for Yokogawa PWANs
    Public Enum EPWAN_Attributes
        U = 1
        I = 2
        P = 3
        S = 4
        Q = 5
        PF = 6
        PHI = 7
        FU = 8
        FI = 9
        THDU = 10
        THDI = 11
        Items = 16
    End Enum
#End Region

#Region "Methods"
    Shared Sub GetHarmLimit_IEC61000_3_2(EquipClass As CPWAN_Helper.EEquipmentClass, ByVal Preal As Single, ByRef HarmLimit() As Single)
        'Calculates harmonic limits depending from measured power and equipment class according to IEC61000-3-2
        'HarmLimit(0) = Max. limit in amps for 1. harmonic
        'HarmLimit(39) = Max. limit in amps for 40. harmonic
        Dim HarmOrder As Integer, LimitClassA As Single, LimitClassD As Single

        For HarmOrder = 1 To 40
            Select Case HarmOrder
                Case 1
                    LimitClassA = 0
                    LimitClassD = 0
                Case 2
                    LimitClassA = 1.08
                    LimitClassD = 0    'odd harmonics only
                Case 3
                    LimitClassA = 2.3
                    LimitClassD = IIf(Preal * 0.0034 < LimitClassA, Preal * 0.0034, LimitClassA)
                Case 4
                    LimitClassA = 0.43
                    LimitClassD = 0    'odd harmonics only
                Case 5
                    LimitClassA = 1.4
                    LimitClassD = IIf(Preal * 0.0019 < LimitClassA, Preal * 0.0019, LimitClassA)   'Limit 5th harmonic
                Case 6
                    LimitClassA = 0.3
                    LimitClassD = 0    'odd harmonics only
                Case 7
                    LimitClassA = 0.77
                    LimitClassD = IIf(Preal * 0.001 < LimitClassA, Preal * 0.001, LimitClassA)     'Limit 7th harmonic
                Case 8
                    LimitClassA = 0.23 * (8 / HarmOrder)
                    LimitClassD = 0    'odd harmonics only
                Case 9
                    LimitClassA = 0.4
                    LimitClassD = IIf(Preal * 0.0005 < LimitClassA, Preal * 0.0005, LimitClassA)   'Limit 9th harmonic
                Case 10
                    LimitClassA = 0.23 * (8 / HarmOrder)
                    LimitClassD = 0    'odd harmonics only
                Case 11
                    LimitClassA = 0.33
                    LimitClassD = IIf(Preal * 0.00035 < LimitClassA, Preal * 0.00035, LimitClassA) 'Limit 11th harmonic
                Case 12
                    LimitClassA = 0.23 * (8 / HarmOrder)
                    LimitClassD = 0    'odd harmonics only
                Case 13
                    LimitClassA = 0.21
                    LimitClassD = IIf(Preal * (0.00385 / HarmOrder) < LimitClassA, Preal * (0.00385 / HarmOrder), LimitClassA) 'Limit 13th harmonic
                Case 14
                    LimitClassA = 0.23 * (8 / HarmOrder)
                    LimitClassD = 0    'odd harmonics only
                Case 15 To 40
                    '*** odd harmonics from 15 to 39 ***
                    If HarmOrder Mod 2 = 1 Then
                        LimitClassA = 0.15 * (15 / HarmOrder)
                        LimitClassD = IIf(Preal * (0.00385 / HarmOrder) < LimitClassA, Preal * (0.00385 / HarmOrder), LimitClassA) 'Limit 15th to 39th harmonic
                        '*** even harmonics from 16 to 40 ***
                    Else
                        LimitClassA = 0.23 * (8 / HarmOrder)
                        LimitClassD = 0  'odd harmonics only
                    End If
            End Select
            HarmLimit(HarmOrder - 1) = IIf(EquipClass = CPWAN_Helper.EEquipmentClass.ClassA, LimitClassA, LimitClassD)
        Next HarmOrder
    End Sub

    Shared Sub GetHarmLimit_IEC61000_3_12(ByRef HarmLimit() As Single)
        'Calculates harmonic limits according to IEC61000-3-12
        'HarmLimit(0) = Max. limit in amps for 1. harmonic
        'HarmLimit(39) = Max. limit in amps for 40. harmonic
        Dim HarmOrder As Integer

        For HarmOrder = 1 To 40
            Select Case HarmOrder
                Case 2
                    HarmLimit(HarmOrder) = 16 / HarmOrder
                Case 3
                    HarmLimit(HarmOrder) = 21.6
                Case 4
                    HarmLimit(HarmOrder) = 16 / HarmOrder
                Case 5
                    HarmLimit(HarmOrder) = 10.7
                Case 6
                    HarmLimit(HarmOrder) = 16 / HarmOrder
                Case 7
                    HarmLimit(HarmOrder) = 7.2
                Case 8
                    HarmLimit(HarmOrder) = 16 / HarmOrder
                Case 9
                    HarmLimit(HarmOrder) = 3.8
                Case 10
                    HarmLimit(HarmOrder) = 16 / HarmOrder
                Case 11
                    HarmLimit(HarmOrder) = 3.1
                Case 12
                    HarmLimit(HarmOrder) = 16 / HarmOrder
                Case 13
                    HarmLimit(HarmOrder) = 2
                Case Else
                    HarmLimit(HarmOrder) = 0
            End Select
        Next HarmOrder
    End Sub
#End Region
End Class
