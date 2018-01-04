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

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target, serial As String
    Dim sensor, ch1 As YGenericSensor

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      sensor = yFirstGenericSensor()
      If sensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      Console.WriteLine("using:" + sensor.get_module().get_serialNumber())
    Else
      sensor = yFindGenericSensor(target + ".genericSensor1")
    End If

    REM retreive module serial number
    serial = sensor.get_module().get_serialNumber()

    REM retreive both channels
    ch1 = yFindGenericSensor(serial + ".genericSensor1")

    While (ch1.isOnline())
      Console.Write("Voltage: " + Str(ch1.get_currentValue()) + ch1.get_unit())
      Console.WriteLine("  (press Ctrl-C to exit)")
      ySleep(1000, errmsg)
    End While
    Console.WriteLine("Module not connected (check identification and USB cable)")
    yFreeAPI()
  End Sub

End Module
