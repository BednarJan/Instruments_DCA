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

    Overrides Function CreateFunctionList() As SortedList

        Visa.SendString(":NUMERIC:NORMAL:PRESET 3" & Chr(10))

        'Fix item order within the preset pattern 3 

        Dim fsl As New SortedList
        fsl.Add("Voltage", "U")
        fsl.Add("Current", "I")
        fsl.Add("ActivePower", "P")
        fsl.Add("ApparentPower", "S")
        fsl.Add("ReactivPower", "Q")
        fsl.Add("PF", "LAMBda")
        fsl.Add("PhaseDiff", "PHI")
        fsl.Add("FrequencyU", "FU")
        fsl.Add("FrequencyI", "FI")
        fsl.Add("VoltPeakPlus", "UPP")
        fsl.Add("VoltPeakMinus", "UMP")
        fsl.Add("CurrentPeakPlus", "IPP")
        fsl.Add("CurrentPeakMinus", "IMP")
        fsl.Add("PowerPeakPlus", "PPP")
        fsl.Add("PowerPeakMinus", "PMP")
        Return fsl

    End Function

    Overrides Sub PresetPattern()
        Dim fsl As SortedList = CreateFunctionList()

        For elm As Integer = 1 To InputElements

            For itm As Integer = 0 To fsl.Count - 1

                Dim fn As KeyValuePair(Of String, String) = fsl(itm)

                SetNumericItem(fn.Value, elm, (elm - 1) * fsl.Count + itm)

            Next
        Next
    End Sub



#End Region

End Class
