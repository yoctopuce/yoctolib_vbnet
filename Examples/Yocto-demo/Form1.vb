Public Class Form1


  Private Sub modulesInventory()
    Dim m, currentmodule As YModule
    Dim name As String
    Dim index As Integer

    ComboBox1.Items.Clear()
    currentmodule = Nothing
    m = YModule.FirstModule()
    While m IsNot Nothing
      name = m.get_serialNumber()
      If Mid(name, 1, 8) = "YCTOPOC1" Then
        ComboBox1.Items.Add(m)
      End If
      m = m.nextModule()
    End While

    If (ComboBox1.Items.Count = 0) Then
      ComboBox1.Enabled = False
      Beacon.Enabled = False
      TestLed.Enabled = False


      ToolStripStatusLabel1.Text = "Connect a Yocto-Demo device"

    Else
      index = 0
      ComboBox1.Enabled = True
      Beacon.Enabled = True
      TestLed.Enabled = True

      For i = 0 To ComboBox1.Items.Count - 1
        If ComboBox1.Items(i).Equals(currentmodule) Then index = i
      Next

      If (ComboBox1.Items.Count = 1) Then
        ToolStripStatusLabel1.Text = "One Yocto-Demo device connected"
      Else
        ToolStripStatusLabel1.Text = ComboBox1.Items.Count.ToString() + " Yocto-Demo devices connected"
      End If
      ComboBox1.SelectedIndex = index

    End If

  End Sub

  Private Sub refreshUI()
    Dim index As Integer

    Dim m As YModule
    Dim led As YLed

    If Not (ComboBox1.Enabled) Then
      index = 4
    Else
      m = CType(ComboBox1.Items(ComboBox1.SelectedIndex), YModule)
      led = YLed.FindLed(m.get_serialNumber() + ".led")
      If led.isOnline() Then
        If led.get_power() = Y_POWER_ON Then
          index = index Or 1
          TestLed.Checked = True
        Else
          TestLed.Checked = False
        End If

        If m.get_beacon() = Y_BEACON_ON Then
          index = index Or 2
          Beacon.Checked = True
        Else
          Beacon.Checked = False
        End If
      End If
    End If

    If index = 0 Then PictureBox1.Image = My.Resources.poc
    If index = 1 Then PictureBox1.Image = My.Resources.pocg
    If index = 2 Then PictureBox1.Image = My.Resources.pocb
    If index = 3 Then PictureBox1.Image = My.Resources.pocbg
    If index = 4 Then PictureBox1.Image = My.Resources.nopoc
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
    RefreshTimer.Interval = 200
    RefreshTimer.Start()

  End Sub

  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InventoryTimer.Tick
    Dim errmsg As String = ""
    YAPI.UpdateDeviceList(errmsg) REM scan for changes
  End Sub



  Private Sub RefreshTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshTimer.Tick
    refreshUI()
  End Sub



  Private Sub TestLed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TestLed.Click
    Dim m As YModule
    Dim led As YLed
    If Not (ComboBox1.Enabled) Then Return
    m = CType(ComboBox1.Items(ComboBox1.SelectedIndex), YModule)
    If Not (m.isOnline()) Then Return
    led = YLed.FindLed(m.get_serialNumber + ".led")
    If led.get_power() = Y_POWER_OFF Then
      led.set_power(Y_POWER_ON)
    Else
      led.set_power(Y_POWER_OFF)
    End If
    refreshUI()
  End Sub


  Private Sub Beacon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Beacon.Click
    Dim m As YModule
    If Not (ComboBox1.Enabled) Then Return
    m = CType(ComboBox1.Items(ComboBox1.SelectedIndex), YModule)
    If Not (m.isOnline()) Then Return

    If m.get_beacon() = Y_BEACON_OFF Then
      m.set_beacon(Y_BEACON_ON)
    Else
      m.set_beacon(Y_BEACON_OFF)
    End If
    refreshUI()
  End Sub
End Class
