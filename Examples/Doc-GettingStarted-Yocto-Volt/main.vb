Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + " <serial_number>")
    Console.WriteLine(execname + " <logical_name>")
    Console.WriteLine(execname + " any  ")
    System.Threading.Thread.Sleep(2500)

    End
  End Sub

  Sub Die(ByVal msg As String)
    Console.WriteLine(msg + "(check USB cable)")
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim sensor As YVoltage
    Dim sensorDC As YVoltage = Nothing
    Dim sensorAC As YVoltage = Nothing
    Dim m As YModule = Nothing

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      REM retreive any voltage sensor (can be AC or DC)
      sensor = yFirstVoltage()
      If sensor Is Nothing Then Die("No module connected")
    Else
      sensor = yFindVoltage(target + ".voltage1")
    End If

    REM  we need to retreive both DC and AC voltage from the device.
    If (sensor.isOnline()) Then
      m = sensor.get_module()
      sensorDC = yFindVoltage(m.get_serialNumber() + ".voltage1")
      sensorAC = yFindVoltage(m.get_serialNumber() + ".voltage2")
    Else
      Die("Module not connected")
    End If

    While (True)
      If Not (m.isOnline()) Then Die("Module not connected")
      Console.Write("DC: " + sensorDC.get_currentValue().ToString() + " v ")
      Console.Write("AC: " + sensorAC.get_currentValue().ToString() + " v ")
      Console.WriteLine("  (press Ctrl-C to exit)")
      ySleep(1000, errmsg)
    End While
    yFreeAPI()

  End Sub
End Module
