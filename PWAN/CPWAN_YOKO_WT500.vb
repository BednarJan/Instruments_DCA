'Class CPWAN_YOKO_WT500
'06.01.2019, A. Zahler
'Compatible Instruments:
'- Yokogawa WT500 with 3 input elements
Imports Ivi.Visa

Public Class CPWAN_YOKO_WT500
    Implements IDevice
    Implements IPWAN

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMaxRms As Single = 1000 Implements IPWAN.VoltageMax
    Public ReadOnly Property CurrentMaxRms As Single = 40 Implements IPWAN.CurrentMax
    Public ReadOnly Property InputElements As UInteger = 3 Implements IPWAN.InputElements
    Public Property ChannelNo As Byte = 1 'Channel 0 = All modules synchronized
#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        _Visa = New CVisaDeviceNI(Session, ErrorLogger)
        _ErrorLogger = ErrorLogger
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"
    Public Function IDN() As String Implements IDevice.IDN
        Dim ErrorMessages(1) As String

        _Visa.SendString("*IDN?", ErrorMessages(0))
        Return _Visa.ReceiveString(ErrorMessages(1))

    End Function

    Public Sub RST() Implements IDevice.RST
        Dim ErrorMessage As String = ""

        _Visa.SendString("*RST", ErrorMessage)
    End Sub

    Public Sub CLS() Implements IDevice.CLS
        Dim ErrorMessage As String = ""

        _Visa.SendString("*CLS", ErrorMessage)
    End Sub

    Public Sub Initialize() Implements IDevice.Initialize
        Dim Element As UInteger, Item As Integer, ItemDef() As String = {"U", "I", "P", "S", "Q", "LAMB", "PHI", "FU", "FI", "UTHD", "ITHD", "NONE", "NONE", "NONE", "NONE", "NONE"}

        _Visa.SendString("*RST;*CLS" & Chr(10))
        _Visa.SendString(":INPUT:VOLTAGE:AUTO:ALL ON" & Chr(10))
        _Visa.SendString(":INPUT:CURRENT:AUTO:ALL ON" & Chr(10))
        _Visa.SendString(":HARMONICS:THD TOTAL" & Chr(10))
        _Visa.SendString(":NUMERIC:FORMAT ASCII" & Chr(10))
        _Visa.SendString(":RATE 500MS" & Chr(10))
        _Visa.SendString(":DISPLAY:NUMERIC:NORMAL:FORMAT VAL16" & Chr(10))
        _Visa.SendString(":NUMERIC:PRESET 2" & Chr(10))

        For Element = 1 To InputElements
            For Item = 1 To CPWAN_Helper.EPWAN_Attributes.Items
                _Visa.SendString(":NUMERIC:ITEM" & Item + (Element - 1) * CPWAN_Helper.EPWAN_Attributes.Items & " " & ItemDef(Item - 1) & Chr(10))
            Next Item
        Next Element
    End Sub
#End Region

#Region "Interface Methodes IPWAN"

    Public Function GetPF(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPF
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.PF & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetVrms(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetVrms
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.U & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetVthd(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetVthd
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.THDU & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetIrms(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetIrms
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.I & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetIthd(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetIthd
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.THDI & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPreal(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPreal
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.P & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPapp(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPapp
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.S & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPreact(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPreact
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.Q & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Function GetFreq(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetFreq
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.FU & Chr(10))
        Return _Visa.ReceiveValue

    End Function

    Public Sub GetHarmonics_IEC61000_3_2(EquipClass As CPWAN_Helper.EEquipmentClass, ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, Optional WaitOnPWAN As Boolean = False) Implements IPWAN.GetHarmonics_IEC61000_3_2
        Dim HarmOrder As Integer, harmstring As String(), Preal As Single

        _Visa.SendString(":HARMONICS:PLLSOURCE U" & _ChannelNo & Chr(10))
        _Visa.SendString(":HARMONICS:ORDER 1,40" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:PRESET 2" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:NUMBER 1" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:VALUE? " & CPWAN_Helper.EPWAN_Attributes.I + (_ChannelNo - 1) * 5 & Chr(10))

        'PWAN delivers string with 40 comma separated values. Split converts into 40 substrings
        harmstring = Split(_Visa.ReceiveString, ",")

        For HarmOrder = 1 To 40
            HarmMeas(HarmOrder - 1) = Val(harmstring(HarmOrder - 1)) 'Converts string to numeric
        Next HarmOrder

        'Read power and calculate limits
        _Visa.SendString(":NUMERIC:VALUE? " & (_ChannelNo - 1) * CPWAN_Helper.EPWAN_Attributes.Items + CPWAN_Helper.EPWAN_Attributes.P & Chr(10))
        Preal = _Visa.ReceiveValue

        CPWAN_Helper.GetHarmLimit_IEC61000_3_2(EquipClass, Preal, HarmLimit)

    End Sub

    Public Sub GetHarmonics_IEC61000_3_12(ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, ByRef THD As Single, ByRef PWHD As Single, Optional WaitOnPWAN As Boolean = False) Implements IPWAN.GetHarmonics_IEC61000_3_12
        Dim HarmMeasCurr(39) As Single, HarmString As String(), Sum As Single

        _Visa.SendString(":HARMONICS:PLLSOURCE U" & _ChannelNo & Chr(10))
        _Visa.SendString(":HARMONICS:ORDER 1,40" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:PRESET 2" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:NUMBER 1" & Chr(10))
        _Visa.SendString(":NUMERIC:LIST:VALUE?" & CPWAN_Helper.EPWAN_Attributes.I + (_ChannelNo - 1) * 5 & Chr(10))

        'PWAN delivers string with 40 comma separated values. Split converts into 40 substrings
        HarmString = Split(_Visa.ReceiveString, ",")

        For HarmOrder = 1 To 40
            HarmMeasCurr(HarmOrder - 1) = Val(HarmString(HarmOrder - 1)) 'Converts string to numeric
            HarmMeas(HarmOrder - 1) = HarmMeasCurr(HarmOrder - 1) / HarmMeasCurr(0) * 100 'Calculate perecentage of fundamental current
        Next HarmOrder

        'Calculate THD (2 to 40 order)
        Sum = 0

        For HarmOrder = 2 To 40
            Sum += (HarmMeas(HarmOrder - 1) / HarmMeas(0)) ^ 2
        Next HarmOrder

        THD = 100 * Math.Sqrt(Sum)

        'Calculate PWHD (14 to 40 order)
        Sum = 0

        For HarmOrder = 14 To 40
            Sum += HarmOrder * (HarmMeas(HarmOrder - 1) / HarmMeas(0)) ^ 2
        Next HarmOrder

        PWHD = 100 * Math.Sqrt(Sum)

        CPWAN_Helper.GetHarmLimit_IEC61000_3_12(HarmLimit)

    End Sub

#End Region

#Region "Public Special Functions WT500"
    'Required before energy or flicker measurement over a defined time period
    Sub IntegrationReset()
        _Visa.SendString(":INTEGRATE:RESET")

    End Sub
#End Region

End Class
