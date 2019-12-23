Public Class Form1


  Private Sub modulesInventory()
    Dim m, currentmodule As YModule
    Dim sensor As YTemperature
    Dim index As Integer

    ComboBox1.Items.Clear()
    currentmodule = Nothing
    sensor = YTemperature.FirstTemperature()
    While sensor IsNot Nothing
      m = sensor.get_module()
      ComboBox1.Items.Add(m)
      sensor = sensor.nextTemperature()
    End While

    If (ComboBox1.Items.Count = 0) Then
      ComboBox1.Enabled = False
      Beacon.Enabled = False
      Label1.Enabled = False
      Label1.Text = "N/A"
      ToolStripStatusLabel3.Text = "Connect a Yocto-device featuring temperature sensor"

    Else
      index = 0
      ComboBox1.Enabled = True
      Beacon.Enabled = True
      Label1.Enabled = True

      For i = 0 To ComboBox1.Items.Count - 1
        If ComboBox1.Items(i).Equals(currentmodule) Then index = i
      Next

      If (ComboBox1.Items.Count = 1) Then
        ToolStripStatusLabel3.Text = "One device connected"
      Else
        ToolStripStatusLabel3.Text = ComboBox1.Items.Count.ToString() + " devices connected"
      End If
      ComboBox1.SelectedIndex = index

    End If

  End Sub

  Private Sub refreshUI()
    Dim index As Integer

    Dim m As YModule
    Dim sensor As YTemperature

    If Not (ComboBox1.Enabled) Then
      index = 4
    Else
      m = ComboBox1.Items(ComboBox1.SelectedIndex)
      sensor = YTemperature.FindTemperature(m.get_serialNumber() + ".temperature")
      If sensor.isOnline() Then

        Label1.Text = sensor.get_currentValue().ToString("0.#") + " °C"
        If m.get_beacon() = Y_BEACON_ON Then
          index = index Or 2
          Beacon.Checked = True
        Else
          Beacon.Checked = False
        End If


      End If
    End If


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


  Private Sub Beacon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Beacon.Click
    Dim m As YModule
    If Not (ComboBox1.Enabled) Then Return
    m = ComboBox1.Items(ComboBox1.SelectedIndex)
    If Not (m.isOnline()) Then Return

    If m.get_beacon() = Y_BEACON_OFF Then
      m.set_beacon(Y_BEACON_ON)
    Else
      m.set_beacon(Y_BEACON_OFF)
    End If
    refreshUI()
  End Sub
End Class
