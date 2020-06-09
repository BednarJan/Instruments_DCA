Public Interface ITempChamber

#Region "Methods (Sub & Functions)"
    Sub SetTemp(nTemp As Single)
    Function GetTemp(ByRef nPresetTemp As Single) As Single
    Sub TurnOff()
    Function RegTemp(nTemp As Single, nDiff As Single) As Boolean
#End Region

End Interface
