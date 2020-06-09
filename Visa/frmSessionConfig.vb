Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms
Imports Ivi.Visa

Public Class frmSessionConfig
    Private _Session As IMessageBasedSession
    Private _SessionProperties As CObjectProperties = Nothing

#Region "Constructor"
    Public Sub New(ByRef Session As IMessageBasedSession, Optional PropertyNames As List(Of String) = Nothing)
        ' This call is required by the designer.
        InitializeComponent()
        ControlBox = False
        Text = ""
        lbTitle.Text = Session.ResourceName

        _Session = Session
        _SessionProperties = New CObjectProperties(_Session, PropertyNames)
        pnlProperties.Controls.Add(_SessionProperties.PropertiesPanel)

        With _SessionProperties.PropertiesPanel
            Me.Height = pnlTopBar.Height + pnlMethods.Height + .PreferredSize.Height + 10
        End With

    End Sub
#End Region

#Region "Private Classes"
    Private Class CObjectProperties
        Private MyObject As Object = Nothing
        Private MyType As Type = Nothing
        Public ReadOnly Properties() As PropertyInfo = Nothing
        Public ReadOnly Property PropertiesPanel As Panel = Nothing

        Public Sub New(ByVal obj As Object, PropertyNames As List(Of String))

            MyObject = obj
            MyType = MyObject.GetType
            Properties = MyType.GetProperties()

            'Apply filter to MyInfo() according to customizable List of PropertyNames
            If Not PropertyNames Is Nothing Then
                Dim MyInfoFiltered = From p As PropertyInfo In Properties Where PropertyNames.Contains(p.Name) Select p
                Properties = MyInfoFiltered.ToArray
            End If

            _PropertiesPanel = getControls(Properties)

        End Sub

        Private Function getControls(ByVal Properties() As PropertyInfo, Optional ByVal HideReadOnlyProperties As Boolean = True) As Panel
            Dim _Panel As New Panel
            Dim InputControls As Control(,)

            'new Panel for Property Controls
            _Panel.Dock = DockStyle.Fill
            _Panel.Padding = New Padding(10, 5, 10, 5)
            _Panel.AutoSize = True
            _Panel.AutoScroll = True
            _Panel.Visible = True

            ReDim InputControls(1, Properties.Count)

            For i As Integer = 0 To Properties.Count - 1

                If HideReadOnlyProperties And Properties(i).CanWrite Then

                    Dim paramLabel As New Label
                    paramLabel.Text = Properties(i).Name & " [" & Properties(i).PropertyType.ToString & "]"
                    paramLabel.Dock = DockStyle.Top
                    paramLabel.TextAlign = ContentAlignment.BottomLeft
                    paramLabel.Font = New Font("Century Gothic", 8)
                    InputControls(0, i) = paramLabel

                    Select Case Properties(i).PropertyType
                        Case GetType(Single), GetType(String), GetType(Integer)
                            InputControls(1, i) = New TextBox
                        Case GetType(Boolean)
                            Dim chkBox As New CheckBox
                            chkBox.Font = New Font("Century Gothic", 8)
                            chkBox.Text = Properties(i).Name
                            paramLabel.Hide()
                            InputControls(1, i) = chkBox
                        Case Else 'Special Types
                            If Properties(i).PropertyType.IsEnum Then
                                Dim EnumBox As New ComboBox
                                EnumBox.DataSource = Properties(i).PropertyType.GetEnumValues
                                InputControls(1, i) = EnumBox
                            ElseIf Properties(i).PropertyType.IsArray Then
                                InputControls(1, i) = New TextBox
                                paramLabel.Text += " Array Example: x;y;z"
                            Else
                                InputControls(1, i) = New TextBox
                            End If
                    End Select

                    _Panel.Controls.Add(InputControls(0, i)) 'add Label with Parametername to ParameterPanel
                    InputControls(0, i).Dock = DockStyle.Top
                    _Panel.Controls.SetChildIndex(InputControls(0, i), 0)

                    _Panel.Controls.Add(InputControls(1, i)) 'add Input Control to ParameterPanel
                    InputControls(1, i).Dock = DockStyle.Top
                    _Panel.Controls.SetChildIndex(InputControls(1, i), 0)
                    InputControls(1, i).Name = Properties(i).Name

                End If
            Next

            Return _Panel
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
        Me.Hide()
    End Sub

#End Region

#Region "Event Handlers"

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

        If Not _SessionProperties Is Nothing Then
            Dim i As Integer = 0

            For Each prop As PropertyInfo In _SessionProperties.Properties

                Dim InputCtrl As Control = _SessionProperties.PropertiesPanel.Controls(prop.Name)
                If Not InputCtrl Is Nothing Then SetPropertyValue(InputCtrl, prop)

            Next

            ' DeviceMethods.RunMethod(Params.MethodName, ParameterValues, Params.ParameterTypes)
        End If

    End Sub

#End Region

#Region "Private Methodes"

    Private Sub SetPropertyValue(ByVal UserInput As Control, ByVal Prop As PropertyInfo)

        Dim val As Object = Nothing

        Try

            If Prop.PropertyType.IsEnum Then
                val = [Enum].Parse(Prop.PropertyType, UserInput.Text)

            ElseIf Prop.PropertyType = GetType(Boolean) Then
                val = Convert.ChangeType(CType(UserInput, CheckBox).Checked, Prop.PropertyType)

            ElseIf Prop.PropertyType.IsArray Then
                Dim strArray() As String = UserInput.Text.Split(";")
                Dim arrObj = Array.CreateInstance(Prop.PropertyType.GetElementType, strArray.Length)

                For i As Integer = 0 To strArray.Length - 1
                    arrObj(i) = Convert.ChangeType(strArray(i), Prop.PropertyType.GetElementType)
                Next i

                val = arrObj

            Else
                val = Convert.ChangeType(UserInput.Text, Prop.PropertyType)
            End If

            Prop.SetValue(_Session, val)

            If Prop.GetValue(_Session) <> val Then
                MsgBox(String.Format("User input of {0} not valid. Value will be automatically changed to next possible value: {1}", Prop.Name, Prop.GetValue(_Session)))

                If Not Prop.PropertyType = GetType(Boolean) Then
                    UserInput.Text = Prop.GetValue(_Session)
                Else
                    CType(UserInput, CheckBox).Checked = Prop.GetValue(_Session)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

End Class