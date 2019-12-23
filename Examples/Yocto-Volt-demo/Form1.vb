Public Class Form1
  Dim maxvalue As Double = 10
  Dim needleposition As Double = -5

  Private Sub modulesInventory()
    Dim m, currentmodule As YModule
    Dim name As String
    Dim index As Integer

    ComboBox1.Items.Clear()
    currentmodule = Nothing
    m = YModule.FirstModule()
    While m IsNot Nothing
      name = m.get_serialNumber()
      If Mid(name, 1, 8) = "VOLTAGE1" Then
        ComboBox1.Items.Add(m)
      End If
      m = m.nextModule()
    End While

    If (ComboBox1.Items.Count = 0) Then
      ComboBox1.Enabled = False
      bt_10V.Enabled = False
      bt_50V.Enabled = False
      bt_300V.Enabled = False
      ACDCcheckBox.Enabled = False
      ToolStripStatusLabel1.Text = "Connect a Yocto-Volt device"
    Else
      index = 0
      ComboBox1.Enabled = True
      bt_10V.Enabled = True
      bt_50V.Enabled = True
      bt_300V.Enabled = True
      ACDCcheckBox.Enabled = True

      For i = 0 To ComboBox1.Items.Count - 1
        If ComboBox1.Items(i).Equals(currentmodule) Then index = i
      Next

      If (ComboBox1.Items.Count = 1) Then
        ToolStripStatusLabel1.Text = "One Yocto-Volt device connected"
      Else
        ToolStripStatusLabel1.Text = ComboBox1.Items.Count.ToString() + " Yocto-Volt devices connected"
      End If
      ComboBox1.SelectedIndex = index
    End If
  End Sub

  Private Sub refreshUI()
    REM draw  the dial.
    Dim value As Double = -5
    Dim DialIsOn As Boolean = False

    If (ComboBox1.Enabled) Then
      REM if a yocto-volt device is connected, lets check it value
      Dim m As YModule = CType(ComboBox1.Items(ComboBox1.SelectedIndex), YModule)
      Dim DC As YVoltage = YVoltage.FindVoltage(m.get_serialNumber() + ".voltage1")
      Dim AC As YVoltage = YVoltage.FindVoltage(m.get_serialNumber() + ".voltage2")
      If (DC.isOnline()) Then
        REM read DC or AC value, according to ACDCcheckBox
        If (ACDCcheckBox.Checked) Then
          value = 100 * (AC.get_currentValue() / maxvalue)
        Else
          value = 100 * (DC.get_currentValue() / maxvalue)
        End If

        DialIsOn = True
      End If
    End If

    REM lets use a double buffering technique to avoid flickering
    Dim BackBuffer As Bitmap
    If DialIsOn Then
      BackBuffer = New Bitmap(My.Resources.bg)
    Else
      BackBuffer = New Bitmap(My.Resources.bgoff)
    End If

    Dim buffer As Graphics = Graphics.FromImage(BackBuffer)
    buffer.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
    Dim DialWidth As Integer = BackBuffer.Width
    Dim DialHeight As Integer = BackBuffer.Height

    REM add inertia to the needle
    needleposition = needleposition + (value - needleposition) / 10

    REM make sure une needle won't go off chart
    If (needleposition < -5) Then needleposition = -5
    If (needleposition > 105) Then needleposition = 105

    Dim angle As Double = 3.1416 * (180 - 30 - 120 * (needleposition / 100)) / 180
    Dim x As Integer = Convert.ToInt32(DialWidth / 2 + Math.Cos(angle) * (DialHeight - 15))
    Dim y As Integer = Convert.ToInt32(DialHeight * 1.066 - Math.Sin(angle) * (DialHeight - 15))

    REM draw the needle shadow
    Dim shadow As Pen = New Pen(Color.FromArgb(16, 0, 0, 0), 3)
    Dim point1 As Point = New Point(Convert.ToInt32(DialWidth / 2 - 3), DialHeight + 3)
    Dim point2 As Point = New Point(Convert.ToInt32(x) - 3, Convert.ToInt32(y) + 3)
    buffer.DrawLine(shadow, point1, point2)

    REM draw the needle
    Dim red As Pen
    If DialIsOn Then
      red = New Pen(Color.FromArgb(255, 255, 0, 0), 3)
    Else
      red = New Pen(Color.FromArgb(255, 64, 0, 0), 3)
    End If

    point1 = New Point(Convert.ToInt32(DialWidth / 2), DialHeight)
    point2 = New Point(Convert.ToInt32(x), Convert.ToInt32(y))
    buffer.DrawLine(red, point1, point2)

    REM draw the scale
    Dim fontFamily As FontFamily = New FontFamily("Arial Narrow")
    Dim font As Font = New Font(fontFamily, Convert.ToInt32(DialHeight / 10), FontStyle.Regular, GraphicsUnit.Pixel)
    buffer.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
    Dim solidBrush As SolidBrush = New SolidBrush(Color.FromArgb(255, 20, 20, 20))

    Dim i As Integer
    For i = 0 To 10

      Dim dvalue As Double = (maxvalue * i / 10)
      angle = 3.1416 * (180 - 30 - 120 * (i / 10.0)) / 180
      Dim text As String = dvalue.ToString()
      Dim size As SizeF = buffer.MeasureString(text, font)
      Dim tx As Integer = Convert.ToInt32(DialWidth / 2 + Math.Cos(angle) * DialHeight * 1.01 - size.Width / 2)
      Dim ty As Integer = Convert.ToInt32(DialHeight * 1.066 - Math.Sin(angle) * DialHeight * 0.98 - size.Height / 2)
      buffer.DrawString(dvalue.ToString(), font, solidBrush, New PointF(tx, ty))
    Next i

    Dim frame As Bitmap = New Bitmap(My.Resources.frame)
    buffer.DrawImage(frame, 0, 0)

    Dim Viewable As Graphics = PictureBox1.CreateGraphics()

    REM fast rendering
    REM Viewable.DrawImageUnscaled(BackBuffer, 0, 0);

    REM slower, but pictureBox can be resized, rendering will still be ok,
    REM try to respect a 2:1 ratio anyway
    Viewable.DrawImage(BackBuffer, New Rectangle(0, 0, PictureBox1.Width, PictureBox1.Height))
    Viewable.Dispose()

  End Sub

  Private Sub devicelistchanged(ByVal m As YModule)
    modulesInventory()
    refreshUI()
  End Sub

  Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    modulesInventory()
    REM we wanna know when device list changes
    YAPI.RegisterDeviceArrivalCallback(AddressOf devicelistchanged)
    YAPI.RegisterDeviceRemovalCallback(AddressOf devicelistchanged)
    InventoryTimer.Interval = 1000
    InventoryTimer.Start()
    RefreshTimer.Interval = 20
    RefreshTimer.Start()
  End Sub

  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InventoryTimer.Tick
    Dim errmsg As String = ""
    YAPI.UpdateDeviceList(errmsg) REM scan for changes
  End Sub

  Private Sub RefreshTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshTimer.Tick
    refreshUI()
  End Sub

  Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

  End Sub

  Private Sub bt_10V_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_10V.Click
    maxvalue = 10
  End Sub

  Private Sub bt_50V_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_50V.Click
    maxvalue = 50
  End Sub

  Private Sub bt_300V_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_300V.Click
    maxvalue = 300
  End Sub
End Class
