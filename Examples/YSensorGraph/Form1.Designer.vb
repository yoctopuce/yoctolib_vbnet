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
    Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
    Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
    Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
    Me.label1 = New System.Windows.Forms.Label()
    Me.comboBox1 = New System.Windows.Forms.ComboBox()
    Me.RecordButton = New System.Windows.Forms.Button()
    Me.PauseButton = New System.Windows.Forms.Button()
    Me.DeleteButton = New System.Windows.Forms.Button()
    Me.loading = New System.Windows.Forms.Label()
    Me.ConnectPlz = New System.Windows.Forms.Label()
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.Status = New System.Windows.Forms.ToolStripStatusLabel()
    Me.progressBar = New System.Windows.Forms.ToolStripProgressBar()
    Me.InventoryTimer = New System.Windows.Forms.Timer(Me.components)
    Me.RefreshTimer = New System.Windows.Forms.Timer(Me.components)
    Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
    Me.StatusStrip1.SuspendLayout()
    CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'label1
    '
    Me.label1.AutoSize = True
    Me.label1.Location = New System.Drawing.Point(12, 17)
    Me.label1.Name = "label1"
    Me.label1.Size = New System.Drawing.Size(92, 13)
    Me.label1.TabIndex = 2
    Me.label1.Text = "Available sensors:"
    '
    'comboBox1
    '
    Me.comboBox1.AccessibleDescription = ""
    Me.comboBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.comboBox1.DisplayMember = "friendlyName"
    Me.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.comboBox1.Enabled = False
    Me.comboBox1.FormattingEnabled = True
    Me.comboBox1.Location = New System.Drawing.Point(110, 14)
    Me.comboBox1.Name = "comboBox1"
    Me.comboBox1.Size = New System.Drawing.Size(433, 21)
    Me.comboBox1.TabIndex = 3
    '
    'RecordButton
    '
    Me.RecordButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.RecordButton.Image = CType(resources.GetObject("RecordButton.Image"), System.Drawing.Image)
    Me.RecordButton.Location = New System.Drawing.Point(554, 11)
    Me.RecordButton.Name = "RecordButton"
    Me.RecordButton.Size = New System.Drawing.Size(25, 25)
    Me.RecordButton.TabIndex = 4
    Me.RecordButton.UseVisualStyleBackColor = True
    '
    'PauseButton
    '
    Me.PauseButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.PauseButton.Image = CType(resources.GetObject("PauseButton.Image"), System.Drawing.Image)
    Me.PauseButton.Location = New System.Drawing.Point(585, 10)
    Me.PauseButton.Name = "PauseButton"
    Me.PauseButton.Size = New System.Drawing.Size(25, 25)
    Me.PauseButton.TabIndex = 5
    Me.PauseButton.UseVisualStyleBackColor = True
    '
    'DeleteButton
    '
    Me.DeleteButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.DeleteButton.Image = CType(resources.GetObject("DeleteButton.Image"), System.Drawing.Image)
    Me.DeleteButton.Location = New System.Drawing.Point(616, 10)
    Me.DeleteButton.Name = "DeleteButton"
    Me.DeleteButton.Size = New System.Drawing.Size(25, 25)
    Me.DeleteButton.TabIndex = 6
    Me.DeleteButton.UseVisualStyleBackColor = True
    '
    'loading
    '
    Me.loading.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.loading.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.loading.Location = New System.Drawing.Point(130, 171)
    Me.loading.Name = "loading"
    Me.loading.Size = New System.Drawing.Size(392, 37)
    Me.loading.TabIndex = 9
    Me.loading.Text = "Loading data, please wait"
    Me.loading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
    Me.loading.Visible = False
    '
    'ConnectPlz
    '
    Me.ConnectPlz.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.ConnectPlz.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.ConnectPlz.Location = New System.Drawing.Point(130, 134)
    Me.ConnectPlz.Name = "ConnectPlz"
    Me.ConnectPlz.Size = New System.Drawing.Size(392, 37)
    Me.ConnectPlz.TabIndex = 10
    Me.ConnectPlz.Text = "Please connect a Yoctopuce device featuring a sensor."
    Me.ConnectPlz.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
    Me.ConnectPlz.Visible = False
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Status, Me.progressBar})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 357)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(653, 22)
    Me.StatusStrip1.TabIndex = 11
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'Status
    '
    Me.Status.AutoSize = False
    Me.Status.Name = "Status"
    Me.Status.Size = New System.Drawing.Size(200, 17)
    Me.Status.Text = "ToolStripStatusLabel1"
    '
    'progressBar
    '
    Me.progressBar.Name = "progressBar"
    Me.progressBar.Size = New System.Drawing.Size(100, 16)
    '
    'InventoryTimer
    '
    Me.InventoryTimer.Interval = 500
    '
    'RefreshTimer
    '
    '
    'Chart1
    '
    Me.Chart1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Chart1.BackColor = System.Drawing.Color.Transparent
    ChartArea1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
    ChartArea1.BackGradientStyle = System.Windows.Forms.DataVisualization.Charting.GradientStyle.TopBottom
    ChartArea1.BackSecondaryColor = System.Drawing.Color.White
    ChartArea1.Name = "ChartArea1"
    ChartArea1.ShadowOffset = 3
    Me.Chart1.ChartAreas.Add(ChartArea1)
    Legend1.Enabled = False
    Legend1.Name = "Legend1"
    Me.Chart1.Legends.Add(Legend1)
    Me.Chart1.Location = New System.Drawing.Point(15, 51)
    Me.Chart1.Name = "Chart1"
    Me.Chart1.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.EarthTones
    Series1.BorderWidth = 2
    Series1.ChartArea = "ChartArea1"
    Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine
    Series1.Color = System.Drawing.Color.Red
    Series1.Legend = "Legend1"
    Series1.Name = "Series1"
    Series1.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.DateTime
    Me.Chart1.Series.Add(Series1)
    Me.Chart1.Size = New System.Drawing.Size(626, 281)
    Me.Chart1.TabIndex = 13
    Me.Chart1.Text = "Chart2"
    Me.Chart1.Visible = False
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(653, 379)
    Me.Controls.Add(Me.Chart1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.ConnectPlz)
    Me.Controls.Add(Me.loading)
    Me.Controls.Add(Me.DeleteButton)
    Me.Controls.Add(Me.PauseButton)
    Me.Controls.Add(Me.RecordButton)
    Me.Controls.Add(Me.comboBox1)
    Me.Controls.Add(Me.label1)
    Me.Name = "Form1"
    Me.Text = "YoctoGraph"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Private WithEvents label1 As System.Windows.Forms.Label
  Private WithEvents comboBox1 As System.Windows.Forms.ComboBox
  Public WithEvents RecordButton As System.Windows.Forms.Button
  Private WithEvents PauseButton As System.Windows.Forms.Button
  Private WithEvents DeleteButton As System.Windows.Forms.Button
  Private WithEvents loading As System.Windows.Forms.Label
  Private WithEvents ConnectPlz As System.Windows.Forms.Label
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Private WithEvents InventoryTimer As System.Windows.Forms.Timer
  Private WithEvents RefreshTimer As System.Windows.Forms.Timer
  Private WithEvents toolTip1 As System.Windows.Forms.ToolTip
  Private WithEvents Chart1 As System.Windows.Forms.DataVisualization.Charting.Chart
  Friend WithEvents Status As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents progressBar As System.Windows.Forms.ToolStripProgressBar

End Class
