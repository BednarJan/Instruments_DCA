Imports Ivi.Visa

Public Interface IVisaDeviceFactory
    Function CreateDevice(ResourceName As String) As IMessageBasedSession
    Sub SetSerialPort(sesion As IMessageBasedSession, params As String)
    'Function GetDefaultRessourceManager() As Ivi.Visa.IResourceManager
End Interface



