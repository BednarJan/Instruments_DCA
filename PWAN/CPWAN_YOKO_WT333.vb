'Class CPWAN_YOKO_WT333
'17.06.2020, JaBe
'Compatible Instruments:
'- Yokogawa WT333E
Imports Ivi.Visa

Public Class CPWAN_YOKO_WT333
    Inherits BCPWAN_YOKO
    Implements IPWAN

#Region "Shorthand Properties"


#End Region

#Region "Constructor"
    Public Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)
        MyBase.New(Session, ErrorLogger)
        VoltageMax = 600
        CurrentMax = 20
        InputElements = 3
    End Sub
#End Region

#Region "Basic Device Functions (IDevice)"

    Public Overrides Sub Initialize() Implements IDevice.Initialize

        MyBase.Initialize()

    End Sub

#End Region

#Region "Interface Methodes IPWAN"


#End Region

#Region "Public Special Functions "

    Public Overrides Sub PresetCurrentProbe(sRatioInMamps As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentProbe

        Throw New NotImplementedException

    End Sub

    Public Overrides Sub PresetCurrentTransformer(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentTransformer

        Throw New NotImplementedException

    End Sub

    Public Overrides Sub PresetVoltDivider(sRatio As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetVoltDivider

        Throw New NotImplementedException

    End Sub

    Public Overrides Sub PresetCurrentShunt(resMiliOhms As Single, sRange As Single, Optional elm As IPWAN.Elements = IPWAN.Elements.Element1) Implements IPWAN.PresetCurrentShunt

        Throw New NotImplementedException

    End Sub


#End Region

End Class
