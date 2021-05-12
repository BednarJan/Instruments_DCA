'Class CPWAN_YOKO_WT200
'17.06.2020, JaBe
'Compatible Instruments:
Imports Ivi.Visa

Public Class CPWAN_YOKO_WT200
    Inherits BCPWAN
    Implements IPWAN

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 600
        CurrentMax = 20
        InputElements = 1
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        Visa.SendString("*RST;*CLS" & Chr(10))
        Visa.SendString(":NUMERIC:FORMAT ASCII" & Chr(10))
        Visa.SendString(":DISPLAY:NUMERIC:NORMAL:FORMAT VAL16" & Chr(10))
        Visa.SendString(":INPUT:VOLTAGE:AUTO ON" & Chr(10))
        Visa.SendString(":INPUT:CURRENT:AUTO ON" & Chr(10))
        Visa.SendString(":HARMONICS:THD TOT" & Chr(10))

        Visa.SendString(":DISPLAY:NORMAL:ITEM1 U" & Chr(10))
        Visa.SendString(":DISPLAY:NORMAL:ITEM2 I" & Chr(10))
        Visa.SendString(":DISPLAY:NORMAL:ITEM3 P" & Chr(10))

    End Sub

#End Region

#Region "Interface Methodes IPWAN"

    Public Overrides Sub SetWiring(iWir As IPWAN.Wiring) Implements IPWAN.SetWiring

        Throw New NotImplementedException

    End Sub


#End Region

#Region "Public Special Functions"


#End Region

End Class
