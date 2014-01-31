Public Class Form1

  Dim caledit(5) As TextBox
  Dim rawedit(5) As TextBox


  Public Sub New()

    ' This call is required by the designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    caledit(0) = C0
    caledit(1) = C1
    caledit(2) = C2
    caledit(3) = C3
    caledit(4) = C4

    rawedit(0) = R0
    rawedit(1) = R1
    rawedit(2) = R2
    rawedit(3) = R3
    rawedit(4) = R4

    'register arrival/removal callbacks
    YAPI.RegisterDeviceArrivalCallback(AddressOf arrivalCallback)
    YAPI.RegisterDeviceRemovalCallback(AddressOf removalCallback)

    ' lets rock!
    Timer1.Enabled = True
  End Sub

  Private Sub arrivalCallback(m As YModule)
    ' add the device in the 1srt combo list
    devicesList.Items.Add(m)
    If devicesList.Items.Count = 1 Then
      devicesList.SelectedIndex = 0
      choosenDeviceChanged()
    End If
  End Sub

  Private Sub removalCallback(m As YModule)
    Dim index As Integer = -1
    Dim i As Integer
    Dim mustrefresh As Boolean

    For i = 0 To devicesList.Items.Count - 1
      If Object.ReferenceEquals(devicesList.Items(i), m) Then index = i
    Next i

    ' if we removed the current module, we must fully refresh the ui
    If (index = devicesList.SelectedIndex) Then
      mustrefresh = True
    Else
      mustrefresh = False
    End If

    If (index >= 0) Then
      ' remove it from the combo box
      devicesList.Items.RemoveAt(index)
      If (devicesList.SelectedIndex >= devicesList.Items.Count) Then devicesList.SelectedIndex = 0
      ' if  we  deleted the active device, we need a refresh
      If (mustrefresh) Then

        functionsList.Enabled = False
        choosenDeviceChanged()
      End If
    End If
  End Sub

  Private Sub choosenDeviceChanged()

    Dim currentDevice As YModule
    Dim fctcount As Integer
    Dim fctName, fctFullName As String
    Dim fct As YFunction

    If devicesList.Items.Count > 0 Then
      devicesList.Enabled = True
    Else
      devicesList.Enabled = False
    End If

    'clear the functions drop down
    functionsList.Items.Clear()
    functionsList.Enabled = devicesList.Enabled

    If Not (devicesList.Enabled) Then

      unsupported_warning.Visible = False
      nosensorfunction.Visible = False
      Exit Sub  ' no device at all connected
    End If

    If (devicesList.SelectedIndex < 0) Then devicesList.SelectedIndex = 0
    currentDevice = CType(devicesList.Items(devicesList.SelectedIndex), YModule)

    ' populate the second drop down
    If (currentDevice.isOnline()) Then
      ' device capabilities inventory
      fctcount = currentDevice.functionCount()
      For i = 0 To fctcount - 1

        fctName = currentDevice.functionId(i)
        fctFullName = currentDevice.get_serialNumber() + "." + fctName
        fct = Nothing
        ' We have to have handle each sensor type independtly, (sorry about that)
        If (fctName.IndexOf("temperature") = 0) Then fct = CType(YTemperature.FindTemperature(fctFullName), YFunction)
        If (fctName.IndexOf("humidity") = 0) Then fct = CType(YHumidity.FindHumidity(fctFullName), YFunction)
        If (fctName.IndexOf("pressure") = 0) Then fct = CType(YPressure.FindPressure(fctFullName), YFunction)
        If (fctName.IndexOf("lightSensor") = 0) Then fct = CType(YLightSensor.FindLightSensor(fctFullName), YFunction)
        If (fctName.IndexOf("carbonDioxide") = 0) Then fct = CType(YCarbonDioxide.FindCarbonDioxide(fctFullName), YFunction)
        If (fctName.IndexOf("voltage") = 0) Then fct = CType(YVoltage.FindVoltage(fctFullName), YFunction)
        If (fctName.IndexOf("current") = 0) Then fct = CType(YCurrent.FindCurrent(fctFullName), YFunction)
        ' add the function in the second drop down
        If Not (fct Is Nothing) Then functionsList.Items.Add(fct)
      Next i

    End If

    If functionsList.Items.Count > 0 Then
      functionsList.Enabled = True
    Else
      functionsList.Enabled = False
    End If

    If (functionsList.Enabled) Then functionsList.SelectedIndex = 0
    refreshFctUI(True)

  End Sub

  Private Sub refreshFctUI(newone As Boolean)
    Dim fct As YFunction
    Dim i As Integer

    nosensorfunction.Visible = False
    toolStripStatusLabel1.Text = devicesList.Items.Count.ToString() + " device(s) found"

    If Not (functionsList.Enabled) Then
      ' disable the UI
      ValueDisplay.Text = "N/A"
      ValueDisplayUnits.Text = "-"
      RawValueDisplay.Text = "-"
      EnableCalibrationUI(False)
      If (devicesList.Enabled) Then
        nosensorfunction.Visible = True
      Else
        toolStripStatusLabel1.Text = "Plug a Yocto-device"
      End If
      Exit Sub
    End If

    fct = CType(functionsList.Items(functionsList.SelectedIndex), YFunction)
    If (newone) Then
      ' enable the ui
      EnableCalibrationUI(True)
      For i = 0 To 4

        caledit(i).Text = ""
        caledit(i).BackColor = System.Drawing.SystemColors.Window
        rawedit(i).Text = ""
        rawedit(i).BackColor = System.Drawing.SystemColors.Window


        If (TypeOf fct Is YTemperature) Then DisplayTemperatureCalPoints(CType(fct, YTemperature))
        If (TypeOf fct Is YPressure) Then DisplayPressureCalPoints(CType(fct, YPressure))
        If (TypeOf fct Is YHumidity) Then DisplayHumidityCalPoints(CType(fct, YHumidity))
        If (TypeOf fct Is YLightSensor) Then DisplayLightSensorCalPoints(CType(fct, YLightSensor))
        If (TypeOf fct Is YCarbonDioxide) Then DisplayCarbonDioxideCalPoints(CType(fct, YCarbonDioxide))
        If (TypeOf fct Is YVoltage) Then DisplayVoltageCalPoints(CType(fct, YVoltage))
        If (TypeOf fct Is YCurrent) Then DisplayCurrentCalPoints(CType(fct, YCurrent))
      Next i
    End If
    If (fct.isOnline()) Then

      If (TypeOf fct Is YTemperature) Then DisplayTemperature(CType(fct, YTemperature))
      If (TypeOf fct Is YPressure) Then DisplayPressure(CType(fct, YPressure))
      If (TypeOf fct Is YHumidity) Then DisplayHumidity(CType(fct, YHumidity))
      If (TypeOf fct Is YLightSensor) Then DisplayLightSensor(CType(fct, YLightSensor))
      If (TypeOf fct Is YCarbonDioxide) Then DisplayCarbonDioxide(CType(fct, YCarbonDioxide))
      If (TypeOf fct Is YVoltage) Then DisplayVoltage(CType(fct, YVoltage))
      If (TypeOf fct Is YCurrent) Then DisplayCurrent(CType(fct, YCurrent))
    End If
  End Sub

  Private Sub displayValue(value As Double, rawvalue As Double, resolution As Double, valunit As String)

    Dim Format As String

    ' displays the sensor value on the ui
    ValueDisplayUnits.Text = valunit

    If (resolution <> YTemperature.RESOLUTION_INVALID) Then
      'if resolution is available on the device the use it to  round the value

      Format = "F" + (-CInt(Math.Round(Math.Log10(resolution)))).ToString()
      RawValueDisplay.Text = "(raw value: " + (resolution * Math.Round(rawvalue / resolution)).ToString(Format) + ")"
      ValueDisplay.Text = (resolution * Math.Round(value / resolution)).ToString(Format)
    Else
      ValueDisplay.Text = value.ToString()
      RawValueDisplay.Text = ""
    End If
  End Sub

  ' enable /disable the calibration data edition
  Private Sub EnableCalibrationUI(state As Boolean)
    Dim i As Integer

    For i = 0 To 4

      caledit(i).Enabled = state
      rawedit(i).Enabled = state
      If Not (state) Then
        caledit(i).Text = ""
        rawedit(i).Text = ""
        caledit(i).BackColor = System.Drawing.SystemColors.Window
        rawedit(i).BackColor = System.Drawing.SystemColors.Window
      End If
    Next i
    RawLabel.Enabled = state
    CalibratedLabel.Enabled = state
    saveBtn.Enabled = state
    cancelBtn.Enabled = state


  End Sub

  Private Sub DisplayCalPoints(ValuesRaw As List(Of Double), ValuesCal As List(Of Double), resolution As Double)
    Dim i As Integer
    ' little trick: if resolution is not available on the device, the
    ' calibration in not available either
    If resolution = YTemperature.RESOLUTION_INVALID Then
      EnableCalibrationUI(False)
      unsupported_warning.Visible = True
      Exit Sub
    End If

    ' display the calibration points
    unsupported_warning.Visible = False
    For i = 0 To ValuesRaw.Count - 1
      caledit(i).Text = ValuesCal(i).ToString()
      rawedit(i).Text = ValuesRaw(i).ToString()
      rawedit(i).BackColor = System.Drawing.Color.FromArgb(&HA0, &HFF, &HA0)
      caledit(i).BackColor = System.Drawing.Color.FromArgb(&HA0, &HFF, &HA0)
    Next i
  End Sub


  Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
    Dim errmsg As String = ""

    ' force an inventory, arrivalCallback and removalCallback
    ' will be called if something changed
    YAPI.UpdateDeviceList(errmsg)

    ' refresh the UI values
    refreshFctUI(False)
  End Sub

  ' this the weak point of the API, methods get_currentValue,
  ' get_resolution,  get_unit etc... are present in all sensor classes, but 
  ' are  not inherited from the parent class (to keep the object model
  ' simple) we have to handle them independtly for each sensor type.

  Private Sub DisplayTemperature(fct As YTemperature)
    displayValue(fct.get_currentValue(), fct.get_currentRawValue(), fct.get_resolution(), fct.get_unit())
  End Sub

  Private Sub DisplayPressure(fct As YPressure)
    displayValue(fct.get_currentValue(), fct.get_currentRawValue(), fct.get_resolution(), fct.get_unit())
  End Sub

  Private Sub DisplayHumidity(fct As YHumidity)
    displayValue(fct.get_currentValue(), fct.get_currentRawValue(), fct.get_resolution(), fct.get_unit())
  End Sub

  Private Sub DisplayLightSensor(fct As YLightSensor)
    displayValue(fct.get_currentValue(), fct.get_currentRawValue(), fct.get_resolution(), fct.get_unit())
  End Sub

  Private Sub DisplayCarbonDioxide(fct As YCarbonDioxide)
    displayValue(fct.get_currentValue(), fct.get_currentRawValue(), fct.get_resolution(), fct.get_unit())
  End Sub

  Private Sub DisplayVoltage(fct As YVoltage)
    displayValue(fct.get_currentValue(), fct.get_currentRawValue(), fct.get_resolution(), fct.get_unit())
  End Sub

  Private Sub DisplayCurrent(fct As YCurrent)
    displayValue(fct.get_currentValue(), fct.get_currentRawValue(), fct.get_resolution(), fct.get_unit())
  End Sub

  Private Sub DisplayTemperatureCalPoints(fct As YTemperature)
    Dim ValuesRaw As List(Of Double) = New List(Of Double)()
    Dim ValuesCal As List(Of Double) = New List(Of Double)()
    fct.loadCalibrationPoints(ValuesRaw, ValuesCal)
    DisplayCalPoints(ValuesRaw, ValuesCal, fct.get_resolution())
  End Sub

  Private Sub DisplayPressureCalPoints(fct As YPressure)
    Dim ValuesRaw As List(Of Double) = New List(Of Double)()
    Dim ValuesCal As List(Of Double) = New List(Of Double)()
    fct.loadCalibrationPoints(ValuesRaw, ValuesCal)
    DisplayCalPoints(ValuesRaw, ValuesCal, fct.get_resolution())
  End Sub

  Private Sub DisplayHumidityCalPoints(fct As YHumidity)
    Dim ValuesRaw As List(Of Double) = New List(Of Double)()
    Dim ValuesCal As List(Of Double) = New List(Of Double)()
    fct.loadCalibrationPoints(ValuesRaw, ValuesCal)
    DisplayCalPoints(ValuesRaw, ValuesCal, fct.get_resolution())
  End Sub

  Private Sub DisplayLightSensorCalPoints(fct As YLightSensor)
    Dim ValuesRaw As List(Of Double) = New List(Of Double)()
    Dim ValuesCal As List(Of Double) = New List(Of Double)()
    fct.loadCalibrationPoints(ValuesRaw, ValuesCal)
    DisplayCalPoints(ValuesRaw, ValuesCal, fct.get_resolution())
  End Sub

  Private Sub DisplayCarbonDioxideCalPoints(fct As YCarbonDioxide)
    Dim ValuesRaw As List(Of Double) = New List(Of Double)()
    Dim ValuesCal As List(Of Double) = New List(Of Double)()
    fct.loadCalibrationPoints(ValuesRaw, ValuesCal)
    DisplayCalPoints(ValuesRaw, ValuesCal, fct.get_resolution())
  End Sub

  Private Sub DisplayVoltageCalPoints(fct As YVoltage)
    Dim ValuesRaw As List(Of Double) = New List(Of Double)()
    Dim ValuesCal As List(Of Double) = New List(Of Double)()
    fct.loadCalibrationPoints(ValuesRaw, ValuesCal)
    DisplayCalPoints(ValuesRaw, ValuesCal, fct.get_resolution())
  End Sub

  Private Sub DisplayCurrentCalPoints(fct As YCurrent)
    Dim ValuesRaw As List(Of Double) = New List(Of Double)()
    Dim ValuesCal As List(Of Double) = New List(Of Double)()
    fct.loadCalibrationPoints(ValuesRaw, ValuesCal)
    DisplayCalPoints(ValuesRaw, ValuesCal, fct.get_resolution())
  End Sub


  Private Sub cancelBtn_Click(sender As System.Object, e As System.EventArgs) Handles cancelBtn.Click
    ' reload the device configuration from the flash
    Dim m As YModule = CType(devicesList.Items(devicesList.SelectedIndex), YModule)
    If (m.isOnline()) Then m.revertFromFlash()
    refreshFctUI(True)
  End Sub


  Private Sub saveBtn_Click(sender As System.Object, e As System.EventArgs) Handles saveBtn.Click
    ' saves the device current configuration into flash
    Dim m As YModule = CType(devicesList.Items(devicesList.SelectedIndex), YModule)
    If m.isOnline() Then m.saveToFlash()
  End Sub

  '  This is the key function: it sets the calibration
  '  data in the device. Note: the parameters are written
  '  in the device RAM, if you want the calibration
  '  to be persistent, you have to call saveToflash();
  Private Sub CalibrationChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles R4.Leave, R3.Leave, R2.Leave, R1.Leave, R0.Leave, C4.Leave, C3.Leave, C2.Leave, C1.Leave, C0.Leave
    Dim ValuesRaw As List(Of Double) = New List(Of Double)()
    Dim ValuesCal As List(Of Double) = New List(Of Double)()
    Dim fct As YFunction
    Dim stopplz As Boolean = False
    Dim i As Integer = 0
    Dim j As Integer = 0

    If (functionsList.SelectedIndex < 0) Then Exit Sub

    Try
      While ((caledit(i).Text <> "") And (rawedit(i).Text <> "") And (i < 5) And Not (stopplz))
        ValuesCal.Add(Convert.ToDouble(caledit(i).Text))
        ValuesRaw.Add(Convert.ToDouble(rawedit(i).Text))
        If (i > 0) Then
          If ValuesRaw(i) <= ValuesRaw(i - 1) Then
            stopplz = True
            i = i - 1
          End If
        End If
        i = i + 1
      End While
    Catch ex As Exception
    End Try

    While ValuesCal.Count > ValuesRaw.Count
      ValuesCal.RemoveAt(ValuesCal.Count - 1)
    End While

    While ValuesRaw.Count > ValuesCal.Count
      ValuesRaw.RemoveAt(ValuesRaw.Count - 1)



    End While

    ' some ui cosmetics: correct values are turned to green
    For j = 0 To i - 1
      caledit(j).BackColor = System.Drawing.Color.FromArgb(&HA0, &HFF, &HA0)
      rawedit(j).BackColor = System.Drawing.Color.FromArgb(&HA0, &HFF, &HA0)
    Next j

    For j = i To 4
      caledit(j).BackColor = System.Drawing.SystemColors.Window
      rawedit(j).BackColor = System.Drawing.SystemColors.Window
    Next j

    ' send the calibration point to the device
    fct = CType(functionsList.Items(functionsList.SelectedIndex), YFunction)
    If fct.isOnline() Then
      If (TypeOf fct Is YTemperature) Then CType(fct, YTemperature).calibrateFromPoints(ValuesRaw, ValuesCal)
      If (TypeOf fct Is YPressure) Then CType(fct, YHumidity).calibrateFromPoints(ValuesRaw, ValuesCal)
      If (TypeOf fct Is YLightSensor) Then CType(fct, YLightSensor).calibrateFromPoints(ValuesRaw, ValuesCal)
      If (TypeOf fct Is YCarbonDioxide) Then CType(fct, YCarbonDioxide).calibrateFromPoints(ValuesRaw, ValuesCal)
      If (TypeOf fct Is YVoltage) Then CType(fct, YVoltage).calibrateFromPoints(ValuesRaw, ValuesCal)
      If (TypeOf fct Is YCurrent) Then CType(fct, YCurrent).calibrateFromPoints(ValuesRaw, ValuesCal)
      If (TypeOf fct Is YHumidity) Then CType(fct, YHumidity).calibrateFromPoints(ValuesRaw, ValuesCal)
    End If

  End Sub

  Private Sub devicesList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles devicesList.SelectedIndexChanged
    choosenDeviceChanged()
  End Sub

  Private Sub functionsList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles functionsList.SelectedIndexChanged
    refreshFctUI(True)
  End Sub
End Class
