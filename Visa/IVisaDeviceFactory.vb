Imports Ivi.Visa

Public Interface IVisaDeviceFactory
    Function CreateDevice(ResourceName As String) As IMessageBasedSession
End Interface



