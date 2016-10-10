Module Module1

  Private Sub Usage()
    Dim execname As String = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + " demo <serial_number>")
    Console.WriteLine(execname + " demo <logical_name>")
    Console.WriteLine(execname + " demo any  ")
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
    Dim sensor As YCurrent
    Dim sensorDC As YCurrent = Nothing
    Dim sensorAC As YCurrent = Nothing
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
      sensor = yFirstCurrent()
      If sensor Is Nothing Then Die("No module connected")
    Else
      sensor = yFindCurrent(target + ".voltage1")
    End If

    REM  we need to retreive both DC and AC voltage from the device.
    If (sensor.isOnline()) Then
      m = sensor.get_module()
      sensorDC = yFindCurrent(m.get_serialNumber() + ".current1")
      sensorAC = yFindCurrent(m.get_serialNumber() + ".current2")
    Else
      Die("Module not connected")
    End If

    While (True)
      If Not (m.isOnline()) Then Die("Module not connected")
      Console.Write("DC: " + sensorDC.get_currentValue().ToString() + " mA ")
      Console.Write("AC: " + sensorAC.get_currentValue().ToString() + " mA ")
      Console.WriteLine("  (press Ctrl-C to exit)")
      ySleep(1000, errmsg)
    End While
    yFreeAPI()

  End Sub
End Module
