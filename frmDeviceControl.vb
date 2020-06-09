Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms

Public Class frmDeviceControl
    Private DeviceMethods As CObjectMethods = Nothing

#Region "Constructor"
    Public Sub New(ByVal DeviceName As String, ByVal Device As Object)
        ' This call is required by the designer.
        InitializeComponent()
        ControlBox = False
        Text = ""
        lbTitle.Text = DeviceName
        DeviceMethods = New CObjectMethods(Device)
    End Sub
#End Region

#Region "Private Classes"
    Private Class CObjectMethods
        Private MyObject As Object = Nothing
        Private MyType As Type = Nothing
        Private MyInfo() As MethodInfo = Nothing
        Public Property Methods As Dictionary(Of String, CMethodParameters) = Nothing


        Public Sub New(ByVal obj As Object)

            _Methods = New Dictionary(Of String, CMethodParameters)

            MyObject = obj
            MyType = MyObject.GetType
            MyInfo = MyType.GetMethods(BindingFlags.DeclaredOnly Or BindingFlags.Public Or BindingFlags.Instance)

            Dim MethodOverload As Integer = 1

            For i As Integer = 0 To MyInfo.Length - 1
                If Not MyInfo(i).IsSpecialName Then
                    If Not _Methods.ContainsKey(MyInfo(i).Name) Then
                        _Methods.Add(MyInfo(i).Name, New CMethodParameters(MyInfo(i).Name, MyInfo(i).GetParameters))
                    Else
                        MethodOverload += 1
                        _Methods.Add(MyInfo(i).Name & "#" & MethodOverload, New CMethodParameters(MyInfo(i).Name, MyInfo(i).GetParameters))
                    End If
                End If
            Next i

        End Sub

        Public Sub RunMethod(ByVal MethodName As String, MethodParameters() As Object, ParameterType() As Type)

            Try
                Dim MyMethod As MethodInfo = MyType.GetMethod(MethodName, ParameterType)

                If Not MyMethod.ReturnType = GetType(Void) Then
                    MsgBox("Return Value: " & MyMethod.Invoke(MyObject, MethodParameters).ToString, MsgBoxStyle.OkOnly, MethodName)
                Else
                    MyMethod.Invoke(MyObject, MethodParameters)
                End If

            Catch
                MsgBox("Error in Method-Call of: " & MyObject.ToString & "." & MethodName)
            End Try

        End Sub

    End Class

    Private Class CMethodParameters
        Public ReadOnly Property Count As Integer
        Public ReadOnly Property MethodName As String
        Public ReadOnly Property Parameters As ParameterInfo() = Nothing
        Public ReadOnly Property ParameterTypes As Type() = Nothing
        Public ReadOnly Property ParameterControls As Control(,) = Nothing
        Public ReadOnly Property ParameterPanel As Panel = Nothing

        Public Sub New(ByVal MethodName As String, ByVal Parameters As ParameterInfo())
            _MethodName = MethodName
            _Count = Parameters.Length - 1
            _Parameters = Parameters
            _ParameterControls = getControls(Parameters)

            ReDim _ParameterTypes(_Count)
            For i As Integer = 0 To _Count
                _ParameterTypes(i) = _Parameters(i).ParameterType
            Next

        End Sub

        Private Function getControls(ByVal params() As ParameterInfo) As Control(,)
            Dim InputControls As Control(,)

            'new Panel for Method Parameters
            _ParameterPanel = New Panel
            _ParameterPanel.Dock = DockStyle.Fill
            _ParameterPanel.Padding = New Padding(10, 5, 10, 5)
            _ParameterPanel.AutoSize = True
            _ParameterPanel.AutoScroll = True
            _ParameterPanel.Visible = False

            ReDim InputControls(1, _Count)

            For i As Integer = 0 To _Count

                Dim paramLabel As New Label
                paramLabel.Text = params(i).Name & " [" & params(i).ParameterType.ToString & "]"
                paramLabel.Dock = DockStyle.Top
                paramLabel.TextAlign = ContentAlignment.BottomLeft
                paramLabel.Font = New Font("Century Gothic", 8)
                InputControls(0, i) = paramLabel

                Select Case params(i).ParameterType
                    Case GetType(Single), GetType(String), GetType(Integer)
                        InputControls(1, i) = New TextBox
                    Case Else
                        If params(i).ParameterType.IsEnum Then
                            Dim EnumBox As New ComboBox
                            EnumBox.DataSource = params(i).ParameterType.GetEnumValues
                            InputControls(1, i) = EnumBox
                        ElseIf params(i).ParameterType.IsArray Then
                            InputControls(1, i) = New TextBox
                            paramLabel.Text += " Array Example: x;y;z"
                        Else
                            InputControls(1, i) = New TextBox
                        End If
                End Select

                _ParameterPanel.Controls.Add(InputControls(0, i)) 'add Label with Parametername to ParameterPanel
                InputControls(0, i).Dock = DockStyle.Top
                _ParameterPanel.Controls.SetChildIndex(InputControls(0, i), 0)

                _ParameterPanel.Controls.Add(InputControls(1, i)) 'add Input Control to ParameterPanel
                InputControls(1, i).Dock = DockStyle.Top
                _ParameterPanel.Controls.SetChildIndex(InputControls(1, i), 0)
                InputControls(1, i).Name = params(i).Name
            Next

            Return InputControls
        End Function

    End Class
#End Region

#Region "Basic Methods: Load, Resize, Drag&Drop, Close"
    Public Const WM_NCLBUTTONDOWN As Integer = &HA1
    Public Const HT_CAPTION As Integer = &H2

    <System.Runtime.InteropServices.DllImportAttribute("user32.dll")>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function

    <System.Runtime.InteropServices.DllImportAttribute("user32.dll")>
    Public Shared Function ReleaseCapture() As Boolean
    End Function

    Private Sub lbTitle_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lbTitle.MouseDown
        ReleaseCapture()
        SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0)
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        End
    End Sub

#End Region

#Region "Event Handlers"
    Private Sub frmDeviceControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For Each kvp As KeyValuePair(Of String, CMethodParameters) In DeviceMethods.Methods
            pnlParameters.Controls.Add(kvp.Value.ParameterPanel)
            cboxMethods.Items.Add(kvp.Key)
        Next
    End Sub

    Private Sub cboxMethods_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxMethods.SelectedIndexChanged
        UpdateParameterPanel()
    End Sub

    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        If cboxMethods.SelectedItem Is Nothing Then
            MsgBox("Please Select a Method")
        Else
            Dim Params As CMethodParameters = DeviceMethods.Methods.Item(cboxMethods.SelectedItem)

            If Not Params Is Nothing Then
                Dim ParameterValues(Params.Count) As Object
                Dim i As Integer = 0

                For Each param As ParameterInfo In Params.Parameters

                    Dim InputCtrl As Control = Params.ParameterPanel.Controls(param.Name)
                    If Not InputCtrl Is Nothing Then

                        Dim obj = String2DataType(InputCtrl.Text, param.ParameterType)

                        If Not obj Is Nothing Then
                            ParameterValues(i) = obj
                            i += 1
                        Else
                            MsgBox("DataType Error: Input of «" & param.Name & "» has to be of type " & param.ParameterType.ToString)
                            Exit Sub
                        End If

                    End If
                Next

                DeviceMethods.RunMethod(Params.MethodName, ParameterValues, Params.ParameterTypes)
            End If
        End If
    End Sub

#End Region

#Region "Private Methodes"
    Private Sub UpdateParameterPanel()

        For Each kvp As KeyValuePair(Of String, CMethodParameters) In DeviceMethods.Methods
            kvp.Value.ParameterPanel.Visible = False
        Next

        With DeviceMethods.Methods.Item(cboxMethods.SelectedItem).ParameterPanel
            .Visible = True
            Me.Height = pnlTopBar.Height + pnlMethods.Height + .PreferredSize.Height + 10
        End With
    End Sub

    Private Function String2DataType(ByVal UserInput As String, ByVal mType As Type) As Object
        'return true if Ok | false if wrong type
        Try
            If mType.IsEnum Then
                Return [Enum].Parse(mType, UserInput)

            ElseIf mType.IsArray Then
                Dim strArray() As String = UserInput.Split(";")
                Dim arrObj = Array.CreateInstance(mType.GetElementType, strArray.Length)

                For i As Integer = 0 To strArray.Length - 1
                    arrObj(i) = Convert.ChangeType(strArray(i), mType.GetElementType)
                Next i

                Return arrObj

            Else
                Dim obj As Object = Convert.ChangeType(UserInput, mType)
                Return obj
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Function GetArray(Of T)(ByVal ParamArray values() As T) As T()
        Return values
    End Function

#End Region

End Class