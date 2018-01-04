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
    Me.comboBox1 = New System.Windows.Forms.ComboBox()
    Me.myMap = New GMap.NET.WindowsForms.GMapControl()
    Me.GPS_Status = New System.Windows.Forms.Label()
    Me.Orient_value = New System.Windows.Forms.Label()
    Me.label4 = New System.Windows.Forms.Label()
    Me.Speed_value = New System.Windows.Forms.Label()
    Me.Lon_value = New System.Windows.Forms.Label()
    Me.Lat_value = New System.Windows.Forms.Label()
    Me.label2 = New System.Windows.Forms.Label()
    Me.label1 = New System.Windows.Forms.Label()
    Me.timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.timer2 = New System.Windows.Forms.Timer(Me.components)
    Me.SuspendLayout()
    '
    'comboBox1
    '
    Me.comboBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.comboBox1.DisplayMember = "FriendlyName"
    Me.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.comboBox1.FormattingEnabled = True
    Me.comboBox1.Location = New System.Drawing.Point(0, 2)
    Me.comboBox1.Name = "comboBox1"
    Me.comboBox1.Size = New System.Drawing.Size(418, 21)
    Me.comboBox1.TabIndex = 2
    Me.comboBox1.ValueMember = "get_friendlyName"
    '
    'myMap
    '
    Me.myMap.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.myMap.Bearing = 0.0!
    Me.myMap.CanDragMap = True
    Me.myMap.EmptyTileColor = System.Drawing.Color.Navy
    Me.myMap.GrayScaleMode = False
    Me.myMap.HelperLineOption = GMap.NET.WindowsForms.HelperLineOptions.DontShow
    Me.myMap.LevelsKeepInMemmory = 5
    Me.myMap.Location = New System.Drawing.Point(3, 29)
    Me.myMap.MarkersEnabled = True
    Me.myMap.MaxZoom = 2
    Me.myMap.MinZoom = 2
    Me.myMap.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.MousePositionAndCenter
    Me.myMap.Name = "myMap"
    Me.myMap.NegativeMode = False
    Me.myMap.PolygonsEnabled = True
    Me.myMap.RetryLoadTile = 0
    Me.myMap.RoutesEnabled = True
    Me.myMap.ScaleMode = GMap.NET.WindowsForms.ScaleModes.[Integer]
    Me.myMap.SelectedAreaFillColor = System.Drawing.Color.FromArgb(CType(CType(33, Byte), Integer), CType(CType(65, Byte), Integer), CType(CType(105, Byte), Integer), CType(CType(225, Byte), Integer))
    Me.myMap.ShowTileGridLines = False
    Me.myMap.Size = New System.Drawing.Size(414, 205)
    Me.myMap.TabIndex = 3
    Me.myMap.Zoom = 0.0R
    '
    'GPS_Status
    '
    Me.GPS_Status.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GPS_Status.Location = New System.Drawing.Point(295, 254)
    Me.GPS_Status.Name = "GPS_Status"
    Me.GPS_Status.Size = New System.Drawing.Size(121, 14)
    Me.GPS_Status.TabIndex = 17
    Me.GPS_Status.Text = "N/A"
    Me.GPS_Status.TextAlign = System.Drawing.ContentAlignment.BottomRight
    '
    'Orient_value
    '
    Me.Orient_value.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Orient_value.Location = New System.Drawing.Point(359, 241)
    Me.Orient_value.Name = "Orient_value"
    Me.Orient_value.Size = New System.Drawing.Size(57, 13)
    Me.Orient_value.TabIndex = 16
    Me.Orient_value.Text = "N/A"
    Me.Orient_value.TextAlign = System.Drawing.ContentAlignment.BottomRight
    '
    'label4
    '
    Me.label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.label4.AutoSize = True
    Me.label4.Location = New System.Drawing.Point(241, 255)
    Me.label4.Name = "label4"
    Me.label4.Size = New System.Drawing.Size(33, 13)
    Me.label4.TabIndex = 15
    Me.label4.Text = "Km/h"
    '
    'Speed_value
    '
    Me.Speed_value.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Speed_value.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Speed_value.Location = New System.Drawing.Point(158, 240)
    Me.Speed_value.Name = "Speed_value"
    Me.Speed_value.Size = New System.Drawing.Size(89, 28)
    Me.Speed_value.TabIndex = 14
    Me.Speed_value.Text = "N/A"
    Me.Speed_value.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'Lon_value
    '
    Me.Lon_value.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Lon_value.Location = New System.Drawing.Point(58, 256)
    Me.Lon_value.Name = "Lon_value"
    Me.Lon_value.Size = New System.Drawing.Size(96, 13)
    Me.Lon_value.TabIndex = 13
    Me.Lon_value.Text = "N/A"
    '
    'Lat_value
    '
    Me.Lat_value.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Lat_value.Location = New System.Drawing.Point(58, 241)
    Me.Lat_value.Name = "Lat_value"
    Me.Lat_value.Size = New System.Drawing.Size(96, 13)
    Me.Lat_value.TabIndex = 12
    Me.Lat_value.Text = "N/A"
    '
    'label2
    '
    Me.label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.label2.AutoSize = True
    Me.label2.Location = New System.Drawing.Point(4, 256)
    Me.label2.Name = "label2"
    Me.label2.Size = New System.Drawing.Size(54, 13)
    Me.label2.TabIndex = 11
    Me.label2.Text = "Longitude"
    '
    'label1
    '
    Me.label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.label1.AutoSize = True
    Me.label1.Location = New System.Drawing.Point(4, 241)
    Me.label1.Name = "label1"
    Me.label1.Size = New System.Drawing.Size(48, 13)
    Me.label1.TabIndex = 10
    Me.label1.Text = "Latitude:"
    '
    'timer1
    '
    Me.timer1.Enabled = True
    Me.timer1.Interval = 500
    '
    'timer2
    '
    Me.timer2.Enabled = True
    Me.timer2.Interval = 25
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(420, 273)
    Me.Controls.Add(Me.GPS_Status)
    Me.Controls.Add(Me.Orient_value)
    Me.Controls.Add(Me.label4)
    Me.Controls.Add(Me.Speed_value)
    Me.Controls.Add(Me.Lon_value)
    Me.Controls.Add(Me.Lat_value)
    Me.Controls.Add(Me.label2)
    Me.Controls.Add(Me.label1)
    Me.Controls.Add(Me.myMap)
    Me.Controls.Add(Me.comboBox1)
    Me.MinimumSize = New System.Drawing.Size(400, 300)
    Me.Name = "Form1"
    Me.Text = "Yocto-GPS-demo"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Private WithEvents comboBox1 As System.Windows.Forms.ComboBox
  Private WithEvents myMap As GMap.NET.WindowsForms.GMapControl
  Private WithEvents GPS_Status As System.Windows.Forms.Label
  Private WithEvents Orient_value As System.Windows.Forms.Label
  Private WithEvents label4 As System.Windows.Forms.Label
  Private WithEvents Speed_value As System.Windows.Forms.Label
  Private WithEvents Lon_value As System.Windows.Forms.Label
  Private WithEvents Lat_value As System.Windows.Forms.Label
  Private WithEvents label2 As System.Windows.Forms.Label
  Private WithEvents label1 As System.Windows.Forms.Label
  Private WithEvents timer1 As System.Windows.Forms.Timer
  Private WithEvents timer2 As System.Windows.Forms.Timer

End Class
