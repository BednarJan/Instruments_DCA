<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSessionConfig
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.pnlTopBar = New System.Windows.Forms.Panel()
        Me.lbTitle = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.pnlMethods = New System.Windows.Forms.TableLayoutPanel()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.lbMethods = New System.Windows.Forms.Label()
        Me.pnlProperties = New System.Windows.Forms.Panel()
        Me.pnlTopBar.SuspendLayout()
        Me.pnlMethods.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTopBar
        '
        Me.pnlTopBar.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(131, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.pnlTopBar.Controls.Add(Me.lbTitle)
        Me.pnlTopBar.Controls.Add(Me.btnClose)
        Me.pnlTopBar.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTopBar.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.pnlTopBar.Location = New System.Drawing.Point(0, 0)
        Me.pnlTopBar.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlTopBar.Name = "pnlTopBar"
        Me.pnlTopBar.Size = New System.Drawing.Size(408, 30)
        Me.pnlTopBar.TabIndex = 2
        '
        'lbTitle
        '
        Me.lbTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbTitle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lbTitle.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbTitle.ForeColor = System.Drawing.Color.White
        Me.lbTitle.Location = New System.Drawing.Point(0, 0)
        Me.lbTitle.Margin = New System.Windows.Forms.Padding(0)
        Me.lbTitle.Name = "lbTitle"
        Me.lbTitle.Padding = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.lbTitle.Size = New System.Drawing.Size(368, 30)
        Me.lbTitle.TabIndex = 5
        Me.lbTitle.Text = "VisaSession"
        Me.lbTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnClose
        '
        Me.btnClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnClose.FlatAppearance.BorderSize = 0
        Me.btnClose.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Maroon
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Wingdings 2", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.btnClose.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.btnClose.Location = New System.Drawing.Point(368, 0)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(0)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(40, 30)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Í"
        Me.btnClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'pnlMethods
        '
        Me.pnlMethods.ColumnCount = 2
        Me.pnlMethods.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 81.95876!))
        Me.pnlMethods.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.04124!))
        Me.pnlMethods.Controls.Add(Me.btnOK, 1, 0)
        Me.pnlMethods.Controls.Add(Me.lbMethods, 0, 0)
        Me.pnlMethods.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMethods.Location = New System.Drawing.Point(0, 30)
        Me.pnlMethods.Name = "pnlMethods"
        Me.pnlMethods.Padding = New System.Windows.Forms.Padding(10, 0, 0, 0)
        Me.pnlMethods.RowCount = 1
        Me.pnlMethods.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 47.14286!))
        Me.pnlMethods.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.pnlMethods.Size = New System.Drawing.Size(408, 34)
        Me.pnlMethods.TabIndex = 5
        '
        'btnOK
        '
        Me.btnOK.Dock = System.Windows.Forms.DockStyle.Top
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOK.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(339, 3)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(66, 28)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'lbMethods
        '
        Me.lbMethods.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbMethods.Location = New System.Drawing.Point(10, 0)
        Me.lbMethods.Margin = New System.Windows.Forms.Padding(0)
        Me.lbMethods.Name = "lbMethods"
        Me.lbMethods.Padding = New System.Windows.Forms.Padding(0, 5, 0, 5)
        Me.lbMethods.Size = New System.Drawing.Size(326, 34)
        Me.lbMethods.TabIndex = 1
        Me.lbMethods.Text = "Visa Session Attributes"
        '
        'pnlProperties
        '
        Me.pnlProperties.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.pnlProperties.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlProperties.Location = New System.Drawing.Point(0, 64)
        Me.pnlProperties.Name = "pnlProperties"
        Me.pnlProperties.Size = New System.Drawing.Size(408, 101)
        Me.pnlProperties.TabIndex = 6
        '
        'frmSessionConfig
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(30, Byte), Integer), CType(CType(30, Byte), Integer), CType(CType(30, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(408, 165)
        Me.Controls.Add(Me.pnlProperties)
        Me.Controls.Add(Me.pnlMethods)
        Me.Controls.Add(Me.pnlTopBar)
        Me.Font = New System.Drawing.Font("Century Gothic", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmSessionConfig"
        Me.Text = "frmDeviceControl"
        Me.pnlTopBar.ResumeLayout(False)
        Me.pnlMethods.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlTopBar As Windows.Forms.Panel
    Friend WithEvents lbTitle As Windows.Forms.Label
    Friend WithEvents btnClose As Windows.Forms.Button
    Friend WithEvents pnlMethods As Windows.Forms.TableLayoutPanel
    Friend WithEvents lbMethods As Windows.Forms.Label
    Friend WithEvents pnlProperties As Windows.Forms.Panel
    Friend WithEvents btnOK As Windows.Forms.Button
End Class
