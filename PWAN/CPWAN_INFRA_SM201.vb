'Class CPWAN_INFRA_SM201
'06.01.2019, A. Zahler
'Compatible Instruments:
'- Infratek SM201
Imports Ivi.Visa

Public Class CPWAN_INFRA_SM201
    Implements IDevice
    Implements IPWAN

    Private _ErrorLogger As CErrorLogger
    Private _strVisa_Adr As String = String.Empty

#Region "Shorthand Properties"
    Public Property Name As String Implements IDevice.Name
    Public Property Visa As IVisaDevice Implements IDevice.Visa
    Public ReadOnly Property VoltageMaxRms As Single = 600 Implements IPWAN.VoltageMax
    Public ReadOnly Property CurrentMaxRms As Single = 30 Implements IPWAN.CurrentMax
    Public ReadOnly Property InputElements As UInteger = 1 Implements IPWAN.InputElements
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
        _Visa.SendString("*RST;*CLS")
        _Visa.SendString("ACQUIRE:INPUT IN30")
        _Visa.SendString("ACQUIRE:RANGE:VOLTAGE AUTO")
        _Visa.SendString("ACQUIRE:RANGE:CURRENT AUTO")
        _Visa.SendString("ACQUIRE:APERTURE 500M")
        _Visa.SendString("ACQUIRE:SYNCHRO VOLTAGE") 'Command not in manual!
        '_Visa.SendString("FORMAT:START 1") '-> Defined in GetHarmonics
        '_Visa.SendString("FORMAT:END 13")  '-> Defined in GetHarmonics
    End Sub
#End Region

#Region "Interface Methodes IPWAN"
    Public Function GetPF(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPF
        _Visa.SendString("POWER:FACTOR?")
        Return _Visa.ReceiveValue

    End Function

    Public Function GetVrms(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetVrms
        _Visa.SendString("VOLTAGE:RMS?")
        Return _Visa.ReceiveValue

    End Function

    Public Function GetVthd(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetVthd
        _Visa.SendString("VOLTAGE:THD?")
        Return _Visa.ReceiveValue

    End Function

    Public Function GetIrms(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetIrms
        _Visa.SendString("CURRENT:RMS?")
        Return _Visa.ReceiveValue

    End Function

    Public Function GetIthd(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetIthd
        _Visa.SendString("CURRENT:THD?")
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPreal(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPreal
        _Visa.SendString("POWER:ACTIVE?")
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPapp(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPapp
        _Visa.SendString("POWER:APPARENT?")
        Return _Visa.ReceiveValue

    End Function

    Public Function GetPreact(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetPreact
        _Visa.SendString("POWER:REACTIVE?")
        Return _Visa.ReceiveValue

    End Function

    Public Function GetFreq(Optional WaitOnPWAN As Boolean = False) As Single Implements IPWAN.GetFreq
        _Visa.SendString("FREQUENCY?")
        Return _Visa.ReceiveValue

    End Function

    Public Sub GetHarmonics_IEC61000_3_2(EquipClass As CPWAN_Helper.EEquipmentClass, ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, Optional WaitOnPWAN As Boolean = False) Implements IPWAN.GetHarmonics_IEC61000_3_2
        Dim HarmOrder As Integer, Preal As Single

        _Visa.SendString("FORMAT:START 1")
        _Visa.SendString("FORMAT:END 40")
        _Visa.SendString("CURRENT:FFT?")

        For HarmOrder = 1 To 40
            HarmMeas(HarmOrder - 1) = _Visa.ReceiveValue
        Next HarmOrder

        'Read power and calculate limits
        _Visa.SendString("POWER:ACTIVE?")
        Preal = _Visa.ReceiveValue

        CPWAN_Helper.GetHarmLimit_IEC61000_3_2(EquipClass, Preal, HarmLimit)

    End Sub

    Public Sub GetHarmonics_IEC61000_3_12(ByRef HarmMeas() As Single, ByRef HarmLimit() As Single, ByRef THD As Single, ByRef PWHD As Single, Optional WaitOnPWAN As Boolean = False) Implements IPWAN.GetHarmonics_IEC61000_3_12
        Dim HarmMeasCurr(39) As Single, Sum As Single

        _Visa.SendString("FORMAT:START 1")
        _Visa.SendString("FORMAT:END 40")
        _Visa.SendString("CURRENT:FFT?")

        For HarmOrder = 1 To 40
            HarmMeasCurr(HarmOrder - 1) = _Visa.ReceiveValue
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

#Region "Public Special Functions SM201"
    'Required before energy or flicker measurement over a defined time period
    Sub IntegrationReset()
        _Visa.SendString("ENERGY:RESET")
    End Sub

    Sub SetFlickerMode()
        _Visa.SendString("ACQUIRE:MEASUREMO FLICKER")
    End Sub

    'Short term flicker
    Private Function GetPST(Optional WaitOnPWAN As Boolean = False) As Single
        _Visa.SendString("VOLTAGE:PST?")
        Return _Visa.ReceiveValue

    End Function

    'Long term flicker
    Private Function GetPLT(Optional WaitOnPWAN As Boolean = False) As Single
        _Visa.SendString("VOLTAGE:PLT?")
        Return _Visa.ReceiveValue

    End Function
#End Region

End Class
