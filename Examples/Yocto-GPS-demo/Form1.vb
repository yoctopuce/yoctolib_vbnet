Imports GMap.NET.WindowsForms
Imports GMap.NET
Imports GMap.NET.MapProviders

Public Class Form1

  Dim overlayOne As GMapOverlay
  Dim currentPos As GMapMarker
  Dim currentGps As YGps
  Dim currentLat As YLatitude
  Dim currentLon As YLongitude
  Dim GPSOk As Boolean = False
  Dim centeringNeeded As Boolean = False


  Private Sub arrivalCallback(m As YModule)
    Dim serial As String = m.get_serialNumber()
    If (serial.Substring(0, 8) = "YGNSSMK1") Then
      comboBox1.Items.Add(m)
      If (comboBox1.Items.Count = 1) Then
        comboBox1.SelectedIndex = 0
      End If
    End If
  End Sub

  Private Sub removalCallback(m As YModule)
    If (comboBox1.Items.Contains(m)) Then
      comboBox1.Items.Remove(m)
    End If
  End Sub

  Private Sub timer1_Tick(sender As System.Object, e As System.EventArgs) Handles timer1.Tick
    Dim errmsg As String = ""
    YAPI.UpdateDeviceList(errmsg)
    refreshUI()
  End Sub


  Private Sub timer2_Tick(sender As System.Object, e As System.EventArgs) Handles timer2.Tick
    Dim errmsg As String = ""
    YAPI.HandleEvents(errmsg)
  End Sub

  Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    YAPI.RegisterDeviceArrivalCallback(AddressOf arrivalCallback)
    YAPI.RegisterDeviceRemovalCallback(AddressOf removalCallback)
  End Sub

  Private Sub myMap_Load(sender As System.Object, e As System.EventArgs) Handles myMap.Load
    myMap.Position = New PointLatLng(46.207388, 6.155904)
    myMap.Manager.Mode = GMap.NET.AccessMode.CacheOnly
    myMap.MapProvider = GMapProviders.BingMap
    myMap.MinZoom = 3
    myMap.MaxZoom = 17
    myMap.Zoom = 10
    myMap.ShowCenter = False
    myMap.DragButton = MouseButtons.Left
    myMap.Manager.Mode = AccessMode.ServerAndCache
    overlayOne = New GMapOverlay("gps position")
    currentPos = New GMap.NET.WindowsForms.Markers.GMarkerGoogle(New PointLatLng(46.207388, 6.155904), New Bitmap(My.Resources.marker))
    overlayOne.Markers.Add(currentPos)
    overlayOne.IsVisibile = False
    myMap.Overlays.Add(overlayOne)
  End Sub

  Private Sub refreshUI()
    GPSOk = False
    If Not (IsNothing(currentGps)) Then
      If (currentGps.isOnline()) Then
        If (currentGps.get_isFixed() = YGps.ISFIXED_TRUE) Then
          GPSOk = True
          Dim lat As Double = currentLat.get_currentValue() / 1000
          Dim lon As Double = currentLon.get_currentValue() / 1000
          currentPos.Position = New PointLatLng(lat, lon)
          Lat_value.Text = currentGps.get_latitude()
          Lon_value.Text = currentGps.get_longitude()
          Speed_value.Text = Math.Round(currentGps.get_groundSpeed()).ToString()
          Orient_value.Text = Math.Round(currentGps.get_direction()).ToString() + "°"
          GPS_Status.Text = currentGps.get_satCount().ToString() + " sat"
          overlayOne.IsVisibile = True
          If (centeringNeeded) Then myMap.Position = currentPos.Position
          centeringNeeded = False
        Else : GPS_Status.Text = "fixing"
        End If
      Else : GPS_Status.Text = "Yocto-GPS disconnected"
      End If
    Else : GPS_Status.Text = "No Yocto-GPS connected"
    End If

    If Not (GPSOk) Then
      Lat_value.Text = "N/A"
      Lon_value.Text = "N/A"
      Speed_value.Text = "N/A"
      Orient_value.Text = "N/A"
      overlayOne.IsVisibile = False
      centeringNeeded = True
    End If
  End Sub


  Private Sub comboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles comboBox1.SelectedIndexChanged
    If (comboBox1.SelectedIndex >= 0) Then
      Dim m As YModule = CType(comboBox1.Items(comboBox1.SelectedIndex), YModule)
      Dim serial As String = m.get_serialNumber()
      currentGps = YGps.FindGps(serial + ".gps")
      currentLat = YLatitude.FindLatitude(serial + ".latitude")
      currentLon = YLongitude.FindLongitude(serial + ".longitude")
      overlayOne.IsVisibile = True
      refreshUI()
      myMap.Position = currentPos.Position
    Else
      overlayOne.IsVisibile = False
      currentGps = Nothing
      currentLat = Nothing
      currentLon = Nothing
    End If
  End Sub



  Private Sub myMap_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles myMap.Paint
    If Not (GPSOk) Then
      e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
      Dim fl As Pen = New Pen(Color.Red, 10)
      e.Graphics.DrawLine(fl, 0, 0, myMap.Size.Width, myMap.Size.Height)
      e.Graphics.DrawLine(fl, 0, myMap.Size.Height, myMap.Size.Width, 0)
    End If
  End Sub
End Class
