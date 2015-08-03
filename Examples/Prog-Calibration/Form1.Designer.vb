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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
    Me.label1 = New System.Windows.Forms.Label()
    Me.RawValueDisplay = New System.Windows.Forms.Label()
    Me.unsupported_warning = New System.Windows.Forms.Label()
    Me.saveBtn = New System.Windows.Forms.Button()
    Me.C4 = New System.Windows.Forms.TextBox()
    Me.C3 = New System.Windows.Forms.TextBox()
    Me.C2 = New System.Windows.Forms.TextBox()
    Me.C1 = New System.Windows.Forms.TextBox()
    Me.C0 = New System.Windows.Forms.TextBox()
    Me.R4 = New System.Windows.Forms.TextBox()
    Me.R3 = New System.Windows.Forms.TextBox()
    Me.R2 = New System.Windows.Forms.TextBox()
    Me.R1 = New System.Windows.Forms.TextBox()
    Me.R0 = New System.Windows.Forms.TextBox()
    Me.CalibratedLabel = New System.Windows.Forms.Label()
    Me.RawLabel = New System.Windows.Forms.Label()
    Me.label4 = New System.Windows.Forms.Label()
    Me.functionsList = New System.Windows.Forms.ComboBox()
    Me.devicesList = New System.Windows.Forms.ComboBox()
    Me.ValueDisplayUnits = New System.Windows.Forms.Label()
    Me.ValueDisplay = New System.Windows.Forms.Label()
    Me.label3 = New System.Windows.Forms.Label()
    Me.label2 = New System.Windows.Forms.Label()
    Me.nosensorfunction = New System.Windows.Forms.Label()
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.toolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.cancelBtn = New System.Windows.Forms.Button()
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'label1
    '
    Me.label1.AutoSize = True
    Me.label1.Location = New System.Drawing.Point(12, 9)
    Me.label1.Name = "label1"
    Me.label1.Size = New System.Drawing.Size(267, 13)
    Me.label1.TabIndex = 1
    Me.label1.Text = "Connect one or more Yocto-devices featuring a sensor:"
    '
    'RawValueDisplay
    '
    Me.RawValueDisplay.AutoSize = True
    Me.RawValueDisplay.Location = New System.Drawing.Point(333, 180)
    Me.RawValueDisplay.Name = "RawValueDisplay"
    Me.RawValueDisplay.Size = New System.Drawing.Size(10, 13)
    Me.RawValueDisplay.TabIndex = 48
    Me.RawValueDisplay.Text = "-"
    '
    'unsupported_warning
    '
    Me.unsupported_warning.AutoSize = True
    Me.unsupported_warning.ForeColor = System.Drawing.Color.Red
    Me.unsupported_warning.Location = New System.Drawing.Point(25, 377)
    Me.unsupported_warning.Name = "unsupported_warning"
    Me.unsupported_warning.Size = New System.Drawing.Size(330, 13)
    Me.unsupported_warning.TabIndex = 47
    Me.unsupported_warning.Text = "This device does not support calibration,  firmware  upgrade needed."
    '
    'saveBtn
    '
    Me.saveBtn.Location = New System.Drawing.Point(358, 324)
    Me.saveBtn.Name = "saveBtn"
    Me.saveBtn.Size = New System.Drawing.Size(75, 23)
    Me.saveBtn.TabIndex = 46
    Me.saveBtn.Text = "Save"
    Me.saveBtn.UseVisualStyleBackColor = True
    '
    'C4
    '
    Me.C4.Location = New System.Drawing.Point(302, 324)
    Me.C4.Name = "C4"
    Me.C4.Size = New System.Drawing.Size(50, 20)
    Me.C4.TabIndex = 44
    '
    'C3
    '
    Me.C3.Location = New System.Drawing.Point(246, 324)
    Me.C3.Name = "C3"
    Me.C3.Size = New System.Drawing.Size(50, 20)
    Me.C3.TabIndex = 41
    '
    'C2
    '
    Me.C2.Location = New System.Drawing.Point(190, 324)
    Me.C2.Name = "C2"
    Me.C2.Size = New System.Drawing.Size(50, 20)
    Me.C2.TabIndex = 37
    '
    'C1
    '
    Me.C1.Location = New System.Drawing.Point(135, 324)
    Me.C1.Name = "C1"
    Me.C1.Size = New System.Drawing.Size(50, 20)
    Me.C1.TabIndex = 35
    '
    'C0
    '
    Me.C0.Location = New System.Drawing.Point(79, 324)
    Me.C0.Name = "C0"
    Me.C0.Size = New System.Drawing.Size(50, 20)
    Me.C0.TabIndex = 31
    '
    'R4
    '
    Me.R4.Location = New System.Drawing.Point(302, 298)
    Me.R4.Name = "R4"
    Me.R4.Size = New System.Drawing.Size(50, 20)
    Me.R4.TabIndex = 43
    '
    'R3
    '
    Me.R3.Location = New System.Drawing.Point(246, 298)
    Me.R3.Name = "R3"
    Me.R3.Size = New System.Drawing.Size(50, 20)
    Me.R3.TabIndex = 40
    '
    'R2
    '
    Me.R2.Location = New System.Drawing.Point(190, 298)
    Me.R2.Name = "R2"
    Me.R2.Size = New System.Drawing.Size(50, 20)
    Me.R2.TabIndex = 36
    '
    'R1
    '
    Me.R1.Location = New System.Drawing.Point(135, 298)
    Me.R1.Name = "R1"
    Me.R1.Size = New System.Drawing.Size(50, 20)
    Me.R1.TabIndex = 33
    '
    'R0
    '
    Me.R0.Location = New System.Drawing.Point(79, 298)
    Me.R0.Name = "R0"
    Me.R0.Size = New System.Drawing.Size(50, 20)
    Me.R0.TabIndex = 30
    '
    'CalibratedLabel
    '
    Me.CalibratedLabel.AutoSize = True
    Me.CalibratedLabel.Location = New System.Drawing.Point(16, 327)
    Me.CalibratedLabel.Name = "CalibratedLabel"
    Me.CalibratedLabel.Size = New System.Drawing.Size(54, 13)
    Me.CalibratedLabel.TabIndex = 42
    Me.CalibratedLabel.Text = "Calibrated"
    '
    'RawLabel
    '
    Me.RawLabel.AccessibleRole = System.Windows.Forms.AccessibleRole.Grip
    Me.RawLabel.AutoSize = True
    Me.RawLabel.Location = New System.Drawing.Point(16, 301)
    Me.RawLabel.Name = "RawLabel"
    Me.RawLabel.Size = New System.Drawing.Size(29, 13)
    Me.RawLabel.TabIndex = 39
    Me.RawLabel.Text = "Raw"
    '
    'label4
    '
    Me.label4.Location = New System.Drawing.Point(45, 224)
    Me.label4.Name = "label4"
    Me.label4.Size = New System.Drawing.Size(372, 55)
    Me.label4.TabIndex = 38
    Me.label4.Text = resources.GetString("label4.Text")
    '
    'functionsList
    '
    Me.functionsList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.functionsList.Enabled = False
    Me.functionsList.FormattingEnabled = True
    Me.functionsList.Location = New System.Drawing.Point(96, 75)
    Me.functionsList.Name = "functionsList"
    Me.functionsList.Size = New System.Drawing.Size(292, 21)
    Me.functionsList.TabIndex = 27
    '
    'devicesList
    '
    Me.devicesList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.devicesList.Enabled = False
    Me.devicesList.FormattingEnabled = True
    Me.devicesList.Location = New System.Drawing.Point(96, 48)
    Me.devicesList.Name = "devicesList"
    Me.devicesList.Size = New System.Drawing.Size(292, 21)
    Me.devicesList.TabIndex = 26
    '
    'ValueDisplayUnits
    '
    Me.ValueDisplayUnits.AutoSize = True
    Me.ValueDisplayUnits.Font = New System.Drawing.Font("Microsoft Sans Serif", 32.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.ValueDisplayUnits.Location = New System.Drawing.Point(258, 133)
    Me.ValueDisplayUnits.Name = "ValueDisplayUnits"
    Me.ValueDisplayUnits.Size = New System.Drawing.Size(37, 51)
    Me.ValueDisplayUnits.TabIndex = 34
    Me.ValueDisplayUnits.Text = "-"
    '
    'ValueDisplay
    '
    Me.ValueDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 32.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.ValueDisplay.Location = New System.Drawing.Point(48, 133)
    Me.ValueDisplay.Name = "ValueDisplay"
    Me.ValueDisplay.Size = New System.Drawing.Size(215, 51)
    Me.ValueDisplay.TabIndex = 32
    Me.ValueDisplay.Text = "N/A"
    Me.ValueDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'label3
    '
    Me.label3.AutoSize = True
    Me.label3.Location = New System.Drawing.Point(35, 78)
    Me.label3.Name = "label3"
    Me.label3.Size = New System.Drawing.Size(48, 13)
    Me.label3.TabIndex = 29
    Me.label3.Text = "Function"
    '
    'label2
    '
    Me.label2.AutoSize = True
    Me.label2.Location = New System.Drawing.Point(35, 51)
    Me.label2.Name = "label2"
    Me.label2.Size = New System.Drawing.Size(44, 13)
    Me.label2.TabIndex = 28
    Me.label2.Text = "Device:"
    '
    'nosensorfunction
    '
    Me.nosensorfunction.AutoSize = True
    Me.nosensorfunction.ForeColor = System.Drawing.Color.Red
    Me.nosensorfunction.Location = New System.Drawing.Point(27, 377)
    Me.nosensorfunction.Name = "nosensorfunction"
    Me.nosensorfunction.Size = New System.Drawing.Size(215, 13)
    Me.nosensorfunction.TabIndex = 49
    Me.nosensorfunction.Text = "No supported sensor function on this device"
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.toolStripStatusLabel1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 417)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(449, 22)
    Me.StatusStrip1.TabIndex = 50
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'toolStripStatusLabel1
    '
    Me.toolStripStatusLabel1.Name = "toolStripStatusLabel1"
    Me.toolStripStatusLabel1.Size = New System.Drawing.Size(234, 17)
    Me.toolStripStatusLabel1.Text = "Plug a Yoctopuce device featuring a sensor"
    '
    'Timer1
    '
    '
    'cancelBtn
    '
    Me.cancelBtn.Location = New System.Drawing.Point(358, 296)
    Me.cancelBtn.Name = "cancelBtn"
    Me.cancelBtn.Size = New System.Drawing.Size(75, 23)
    Me.cancelBtn.TabIndex = 51
    Me.cancelBtn.Text = "Cancel"
    Me.cancelBtn.UseVisualStyleBackColor = True
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(449, 439)
    Me.Controls.Add(Me.cancelBtn)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.nosensorfunction)
    Me.Controls.Add(Me.RawValueDisplay)
    Me.Controls.Add(Me.unsupported_warning)
    Me.Controls.Add(Me.saveBtn)
    Me.Controls.Add(Me.C4)
    Me.Controls.Add(Me.C3)
    Me.Controls.Add(Me.C2)
    Me.Controls.Add(Me.C1)
    Me.Controls.Add(Me.C0)
    Me.Controls.Add(Me.R4)
    Me.Controls.Add(Me.R3)
    Me.Controls.Add(Me.R2)
    Me.Controls.Add(Me.R1)
    Me.Controls.Add(Me.R0)
    Me.Controls.Add(Me.CalibratedLabel)
    Me.Controls.Add(Me.RawLabel)
    Me.Controls.Add(Me.label4)
    Me.Controls.Add(Me.functionsList)
    Me.Controls.Add(Me.devicesList)
    Me.Controls.Add(Me.ValueDisplayUnits)
    Me.Controls.Add(Me.ValueDisplay)
    Me.Controls.Add(Me.label3)
    Me.Controls.Add(Me.label2)
    Me.Controls.Add(Me.label1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.Name = "Form1"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Calibration settings"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Private WithEvents label1 As System.Windows.Forms.Label
  Private WithEvents RawValueDisplay As System.Windows.Forms.Label
  Private WithEvents unsupported_warning As System.Windows.Forms.Label
  Private WithEvents saveBtn As System.Windows.Forms.Button
  Private WithEvents C4 As System.Windows.Forms.TextBox
  Private WithEvents C3 As System.Windows.Forms.TextBox
  Private WithEvents C2 As System.Windows.Forms.TextBox
  Private WithEvents C1 As System.Windows.Forms.TextBox
  Private WithEvents C0 As System.Windows.Forms.TextBox
  Private WithEvents R4 As System.Windows.Forms.TextBox
  Private WithEvents R3 As System.Windows.Forms.TextBox
  Private WithEvents R2 As System.Windows.Forms.TextBox
  Private WithEvents R1 As System.Windows.Forms.TextBox
  Private WithEvents R0 As System.Windows.Forms.TextBox
  Private WithEvents CalibratedLabel As System.Windows.Forms.Label
  Private WithEvents RawLabel As System.Windows.Forms.Label
  Private WithEvents label4 As System.Windows.Forms.Label
  Private WithEvents functionsList As System.Windows.Forms.ComboBox
  Private WithEvents devicesList As System.Windows.Forms.ComboBox
  Private WithEvents ValueDisplayUnits As System.Windows.Forms.Label
  Private WithEvents ValueDisplay As System.Windows.Forms.Label
  Private WithEvents label3 As System.Windows.Forms.Label
  Private WithEvents label2 As System.Windows.Forms.Label
  Private WithEvents nosensorfunction As System.Windows.Forms.Label
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents toolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Private WithEvents cancelBtn As System.Windows.Forms.Button

End Class
