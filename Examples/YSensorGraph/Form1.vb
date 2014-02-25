Imports System.Windows.Forms.DataVisualization.Charting

Public Class Form1

  Private Sub Chart1_Click(sender As System.Object, e As System.EventArgs)

  End Sub

  REM used to compute the graph length and time label format
  Private FirstPointDate As Double
  Private LastPointDate As Double

  Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    REM we wanna know when device list changes
    YAPI.RegisterDeviceArrivalCallback(AddressOf deviceArrival)
    YAPI.RegisterDeviceRemovalCallback(AddressOf deviceRemoval)
    InventoryTimer.Interval = 500
    InventoryTimer.Start()
    RefreshTimer.Interval = 500
    RefreshTimer.Start()
  End Sub

  REM MS doesn't seem to like UNIX timestamps, we have to do the convertion ourself :-)
  Private Function UnixTimeStampToDateTime(unixTimeStamp As Double) As DateTime
    REM, Unix timestamp is seconds past epoch
    Dim d As System.DateTime = New DateTime(1970, 1, 1, 0, 0, 0, 0)
    d = d.AddSeconds(unixTimeStamp).ToLocalTime()
    Return d
  End Function

  REM update the UI according to the sensors count
  Private Sub setSensorCount()
    If (comboBox1.Items.Count <= 0) Then
      Status.Text = "No sensor found, check USB cable"
    ElseIf (comboBox1.Items.Count = 1) Then
      Status.Text = "One sensor found"
    Else
      Status.Text = Convert.ToString(comboBox1.Items.Count) + " sensors found"
    End If

    If (comboBox1.Items.Count <= 0) Then Chart1.Visible = False
    ConnectPlz.Visible = comboBox1.Items.Count <= 0
    Application.DoEvents()
  End Sub

  REM automatically called each time a new yoctopuce device is plugged
  Private Sub deviceArrival(m As YModule)
    REM new device just arrived, lets enumerate all sensors and
    REM add the one missing to the combobox           
    Dim s As YSensor = YSensor.FirstSensor()
    While Not (IsNothing(s))
      If Not (comboBox1.Items.Contains(s)) Then
        comboBox1.Items.Add(s)
      End If
      s = s.nextSensor()
      comboBox1.Enabled = comboBox1.Items.Count > 0
      If ((comboBox1.SelectedIndex < 0) And (comboBox1.Items.Count > 0)) Then
        comboBox1.SelectedIndex = 0
      End If
      setSensorCount()
    End While
  End Sub

  REM automatically called each time a new yoctopuce device is unplugged
  Private Sub deviceRemoval(m As YModule)
    REM a device vas just removed, lets remove the offline sensors
    REM from the combo box
    Dim i As Integer
    For i = comboBox1.Items.Count - 1 To 0 Step -1
      If Not (CType(comboBox1.Items(i), YSensor).isOnline()) Then
        comboBox1.Items.RemoveAt(i)
      End If

    Next i
    setSensorCount()
  End Sub

  REM automatically called on a regular basis with sensor value
  Private Sub newSensorValue(f As YFunction, m As YMeasure)
    Dim t As Double = m.get_endTimeUTC()
    Chart1.Series(0).Points.AddXY(UnixTimeStampToDateTime(t), m.get_averageValue())
    If (FirstPointDate < 0) Then FirstPointDate = t
    LastPointDate = t
    setGraphScale()
  End Sub

  Private Sub InventoryTimer_Tick(sender As System.Object, e As System.EventArgs) Handles InventoryTimer.Tick
    Dim errmsg As String = ""
    YAPI.UpdateDeviceList(errmsg)
  End Sub

  REM returns the sensor selected in the combobox
  Private Function getSelectedSensor() As YSensor
    Dim index As Integer = comboBox1.SelectedIndex
    If (index < 0) Then Return Nothing
    REM configure timed report callback for this function
    Return CType(comboBox1.Items(index), YSensor)
  End Function

  REM update the datalogger control buttons
  Private Sub refreshDatloggerButton(s As YSensor)
    If Not (IsNothing(s)) Then
      Dim m As YModule = s.get_module() REM get the module harboring the sensor
      Dim dtl As YDataLogger = YDataLogger.FindDataLogger(m.get_serialNumber() + ".dataLogger")
      If (dtl.isOnline()) Then
        If (dtl.get_recording() = YDataLogger.RECORDING_ON) Then
          RecordButton.Enabled = False
          PauseButton.Enabled = True
          DeleteButton.Enabled = False
          Return
        Else
          RecordButton.Enabled = True
          PauseButton.Enabled = False
          DeleteButton.Enabled = True
          Return
        End If
      End If
    End If
    RecordButton.Enabled = False
    PauseButton.Enabled = False
    DeleteButton.Enabled = False
  End Sub

  REM update the date labels format according to graph length
  Private Sub setGraphScale()
    Dim count As Integer = Chart1.Series(0).Points.Count
    If (count > 0) Then
      Dim total As Double = LastPointDate - FirstPointDate
      If (total < 180) Then
        Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "H:mm:ss"
      ElseIf (total < 3600) Then
        Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "H:mm"
      ElseIf (total < 3600 * 24) Then
        Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "h:mm"
      ElseIf (total < 3600 * 24 * 7) Then
        Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "ddd H"
      ElseIf (total < 3600 * 24 * 30) Then
        Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "dd-MMM"
      Else
        Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "MMM"
      End If
    Else
      Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "mm:ss"
    End If
  End Sub

  REM clear the graph
  Private Sub clearGraph()
    Chart1.Series(0).XValueType = ChartValueType.DateTime
    Chart1.Series(0).Points.SuspendUpdates()
    REM chart1.Series[0].Points.Clear();  indecently slow
    While (Chart1.Series(0).Points.Count > 0)
      Chart1.Series(0).Points.RemoveAt(Chart1.Series(0).Points.Count - 1)
    End While
    Chart1.Series(0).Points.ResumeUpdates()
  End Sub

  REM the core function :  load data from datalogger to send it to the graph
  Private Sub comboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles comboBox1.SelectedIndexChanged
    REM lets hide the graph wgile updating
    Chart1.Visible = False
    comboBox1.Enabled = False

    REM remove any previous timed report call back 
    Dim i As Integer
    For i = 0 To comboBox1.Items.Count - 1
      CType(comboBox1.Items(i), YSensor).registerTimedReportCallback(Nothing)
    Next i

    REM allow zooming
    Chart1.ChartAreas(0).CursorX.Interval = 0.001
    Chart1.ChartAreas(0).CursorX.IsUserEnabled = True
    Chart1.ChartAreas(0).CursorX.IsUserSelectionEnabled = True
    Chart1.ChartAreas(0).CursorX.AutoScroll = True
    Chart1.ChartAreas(0).AxisX.ScaleView.Zoomable = True
    Chart1.ChartAreas(0).AxisX.ScrollBar.IsPositionedInside = True

    Dim index As Integer = comboBox1.SelectedIndex
    If (index >= 0) Then clearGraph()


    Dim s As YSensor = getSelectedSensor()
    If Not (IsNothing(s)) Then
      FirstPointDate = -1
      LastPointDate = -1
      REM some ui control
      loading.Visible = True
      refreshDatloggerButton(Nothing)
      progressBar.Visible = True
      Status.Text = "Loading data from datalogger..."
      Dim j As Integer
      For j = 1 To 100
        REM makes sure the UI changes are repainted
        Application.DoEvents()
      Next j


      REM load data from datalogger
      Dim data As YDataSet = s.get_recordedData(0, 0)
      Dim progress As Integer = data.loadMore()
      While (progress < 100)
        progressBar.Value = progress
        Application.DoEvents()
        progress = data.loadMore()
      End While

      REM sets the unit (because ° is not a ASCII-128  character, Yoctopuce temperature
      REM sensors report unit as 'C , so we fix it).
      Chart1.ChartAreas(0).AxisY.Title = s.get_unit().Replace("'C", "°C")
      Chart1.ChartAreas(0).AxisY.TitleFont = New Font("Arial", 12, FontStyle.Regular)

      REM send the data to the graph
      Dim alldata As List(Of YMeasure) = data.get_measures()
      For i = 0 To alldata.Count - 1
        Chart1.Series(0).Points.AddXY(UnixTimeStampToDateTime(alldata(i).get_endTimeUTC()), alldata(i).get_averageValue())
      Next i

      REM used to compute graph length
      If (alldata.Count > 0) Then
        FirstPointDate = alldata(0).get_endTimeUTC()
        LastPointDate = alldata(alldata.Count - 1).get_endTimeUTC()
      End If

      setGraphScale()

      REM restore UI
      comboBox1.Enabled = True
      progressBar.Visible = False
      setSensorCount()
      s.set_reportFrequency("3/s")
      s.registerTimedReportCallback(AddressOf newSensorValue)
      loading.Visible = False
      Chart1.Visible = True
      refreshDatloggerButton(s)
    End If
  End Sub


  Private Sub RefreshTimer_Tick(sender As System.Object, e As System.EventArgs) Handles RefreshTimer.Tick
    Dim errmsg As String = ""
    YAPI.HandleEvents(errmsg)
  End Sub





  Private Sub DataLoggerButton_Click(sender As System.Object, e As System.EventArgs) Handles RecordButton.Click
    Dim s As YSensor = getSelectedSensor()

    If Not (IsNothing(s)) Then
      Dim m As YModule = s.get_module()  REM get the module harboring the sensor
      Dim dtl As YDataLogger = YDataLogger.FindDataLogger(m.get_serialNumber() + ".dataLogger")
      If (dtl.isOnline()) Then
        If (sender Is RecordButton) Then dtl.set_recording(YDataLogger.RECORDING_ON)
        If (sender Is PauseButton) Then dtl.set_recording(YDataLogger.RECORDING_OFF)
        If (sender Is DeleteButton) Then
          dtl.set_recording(YDataLogger.RECORDING_OFF)
          MessageBox.Show("clear")
          dtl.forgetAllDataStreams()
          clearGraph()
        End If
      End If
    End If
    refreshDatloggerButton(s)
  End Sub
End Class
