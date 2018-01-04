<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
    Me.components = New System.ComponentModel.Container()
    Me.ComboBox1 = New System.Windows.Forms.ComboBox()
    Me.InventoryTimer = New System.Windows.Forms.Timer(Me.components)
    Me.Beacon = New System.Windows.Forms.CheckBox()
    Me.RefreshTimer = New System.Windows.Forms.Timer(Me.components)
    Me.Label1 = New System.Windows.Forms.Label()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'ComboBox1
    '
    Me.ComboBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.ComboBox1.Enabled = False
    Me.ComboBox1.FormattingEnabled = True
    Me.ComboBox1.Location = New System.Drawing.Point(2, 12)
    Me.ComboBox1.Name = "ComboBox1"
    Me.ComboBox1.Size = New System.Drawing.Size(298, 21)
    Me.ComboBox1.TabIndex = 0
    '
    'InventoryTimer
    '
    '
    'Beacon
    '
    Me.Beacon.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Beacon.AutoSize = True
    Me.Beacon.Location = New System.Drawing.Point(237, 79)
    Me.Beacon.Name = "Beacon"
    Me.Beacon.Size = New System.Drawing.Size(63, 17)
    Me.Beacon.TabIndex = 4
    Me.Beacon.Text = "Beacon"
    Me.Beacon.UseVisualStyleBackColor = True
    '
    'RefreshTimer
    '
    Me.RefreshTimer.Interval = 200
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label1.Location = New System.Drawing.Point(109, 59)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(72, 37)
    Me.Label1.TabIndex = 5
    Me.Label1.Text = "N/A"
    '
    'ToolStripStatusLabel1
    '
    Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
    Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(111, 17)
    Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
    '
    'ToolStripStatusLabel2
    '
    Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
    Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(111, 17)
    Me.ToolStripStatusLabel2.Text = "ToolStripStatusLabel2"
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel3})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 117)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(301, 22)
    Me.StatusStrip1.SizingGrip = False
    Me.StatusStrip1.TabIndex = 6
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'ToolStripStatusLabel3
    '
    Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
    Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(111, 17)
    Me.ToolStripStatusLabel3.Text = "ToolStripStatusLabel3"
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(301, 139)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.Beacon)
    Me.Controls.Add(Me.ComboBox1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.Name = "Form1"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Yocto-Demo"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
  Friend WithEvents InventoryTimer As System.Windows.Forms.Timer
  Friend WithEvents Beacon As System.Windows.Forms.CheckBox
  Friend WithEvents RefreshTimer As System.Windows.Forms.Timer
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel

End Class
