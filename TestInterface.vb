﻿Imports System.Configuration
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Xml.Serialization
Imports Ivi.Visa


Module TestInterface

    Public Sub main()

        Const VISAAddrTC As String = "ASRL7::INSTR"

        Dim DeviceFactory As New CVisaDeviceFactory
        Dim ErrorLogger As New CErrorLogger("C:\ErrorLog\ErrorLog.txt")
        Dim ResMngr As New CVisaManager()
        Dim MySession As IMessageBasedSession


        'Dim DCSource As New CSourceDC_SORENSEN_60_10(DeviceFactory.CreateDevice(ResMngr.Resources.Item(7)), ErrorLogger)
        'DCSource.Name = "Sorensen DLM 60-10"


        MySession = DeviceFactory.CreateDevice(VISAAddrTC)

        'SerialParity port settings in the order: baudrate, SerialParity, SerialStopBitsMode, Databits, FlowControlModes  

        DeviceFactory.SetSerialPort(MySession, "9600,0,1,8,N")

        Dim myTC As ITC = New CTC_WEISS_WT3340(MySession, ErrorLogger)

        Dim myPSU As ISource_AC = New CSourceDC_AGILENT_E4356A(MySession, ErrorLogger)

        Dim myPSU_AC As ISource_AC = New CSourceAC_AMETEK_BPS_30(MySession, ErrorLogger)

        'myPSU_AC.SetPhaseMode(ISource_AC.EPhaseMode.ThreePhase)
        'myPSU_AC.SetRange(ISource_AC.ERange.LOW)

        'myTC.HeatGrad = 1
        'myTC.CoolGrad = 1
        'myTC.SetGradients()

        'myTC.SetpointTemp = 22
        'myTC.TurnON()
        'myTC.TurnOFF()

        'Dim PWAN As New CPWAN_HIOKI_PW3337(DeviceFactory.CreateDevice(ResMngr.Resources.Item(10)), ErrorLogger)
        'PWAN.Name = "HIOKI PW3337"

        'Dim frmControl As New frmDeviceControl(DCSource.Name, DCSource)

        'frmControl.ShowDialog()

        'DMM.Visa.ShowConfigUI()

        'For i As Integer = 1 To 5000
        '    Dim Volt As Single = DMM.MeasDC
        '    cHelper.Delay(0.001)
        '    Console.WriteLine(Volt.ToString)
        'Next i

    End Sub


    Public Sub Serialize(ByVal obj As Object, FileName As String)

        Dim objStreamWriter As New StreamWriter(FileName & ".xml")


        Using memoryStream As MemoryStream = New MemoryStream()

            Using reader As StreamReader = New StreamReader(memoryStream)
                Dim serializer As DataContractSerializer = New DataContractSerializer(obj.[GetType]())
                serializer.WriteObject(memoryStream, obj)
                memoryStream.Position = 0

                objStreamWriter.Write(reader.ReadToEnd())
                objStreamWriter.Close() 'Close File

            End Using
        End Using
    End Sub

    'Public Shared Function Deserialize(ByVal xml As String, ByVal toType As Type) As Object
    '    Using stream As Stream = New MemoryStream()
    '        Dim data As Byte() = System.Text.Encoding.UTF8.GetBytes(xml)
    '        stream.Write(data, 0, data.Length)
    '        stream.Position = 0
    '        Dim deserializer As DataContractSerializer = New DataContractSerializer(toType)
    '        Return deserializer.ReadObject(stream)
    '    End Using
    'End Function




    'Sub Serialize(Obj As Object, FileName As String)

    '    'Serialize student object to an XML file, via the use of StreamWriter
    '    Dim objStreamWriter As New StreamWriter(FileName & ".xml")

    '    Dim xsSerialize As New XmlSerializer(Obj.GetType) 'Determine what object types are present

    '    xsSerialize.Serialize(objStreamWriter, Obj) 'Save

    '    objStreamWriter.Close() 'Close File

    'End Sub

    'Function Deserialize(FileName As String) As Object

    '    'Deserialize XML file to a new object.
    '    Dim objStreamReader As New StreamReader(FileName & ".xml") 'Read File

    '    Dim obj As New Object() 'Instantiate new Object

    '    Dim xsDeserialize As New XmlSerializer(obj.GetType) 'Get Info present

    '    obj = xsDeserialize.Deserialize(objStreamReader) 'Deserialize / Open

    '    objStreamReader.Close() 'Close Reader

    'End Function

End Module
