Public Interface IDevice
    Property Name As String
    ReadOnly Property Visa As IVisaDevice  'IO Visa Interface

#Region "Basic Device Functions"
    Function IDN() As String
    Sub RST()
    Sub CLS()
    Sub Initialize()
#End Region

End Interface
