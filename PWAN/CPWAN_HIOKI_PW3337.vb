'Class CPWAN_HIOKI_PW3337
'02.05.2019, J. Bednar 
'17.06.2020, J. Bednar
'Compatible Instruments:
'- HIOKI PW3337
Imports Ivi.Visa

Public Class CPWAN_HIOKI_PW3337
    Inherits BCPWAN
    Implements IPWAN

#Region "Shorthand Properties"

#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 1000
        CurrentMax = 65
        InputElements = 3
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"


    Public Overrides Sub Initialize() Implements IDevice.Initialize
        Dim ItemDef() As String = {"U", "I", "P", "S", "Q", "LAMB", "PHI", "FU", "FI", "UTHD", "ITHD", "NONE", "NONE", "NONE", "NONE", "NONE"}

        Visa.SendString("*RST;*CLS" & Chr(10))
        Visa.SendString("HEADER OFF" & Chr(10))

        For Element As Integer = 1 To InputElements
            For Item As Integer = 1 To IPWAN.PA_Attributes.Items
                Visa.SendString(":NUMERIC:ITEM" & Item + (Element - 1) * CPWAN_Helper.EPWAN_Attributes.Items & " " & ItemDef(Item - 1) & Chr(10))
            Next Item
        Next Element

    End Sub
#End Region

#Region "Interface Methodes IPWAN"


#End Region

#Region "Public Special Functions HIOKI PW3337"

#End Region

End Class
