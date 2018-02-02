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
    Dim target As String
    Dim sensor As YWeighScale

    If argv.Length < 2 Then Usage()

    target = argv(1)
    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      sensor = yFirstWeighScale()
      If sensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      Console.WriteLine("using:" + sensor.get_module().get_serialNumber())
    Else
      sensor = yFindWeighScale(target + ".weighScale1")
    End If

    If sensor.isOnline() Then
      REM On startup, enable excitation and tare weigh scale
      Console.WriteLine("Resetting tare weight...")
      sensor.set_excitation(YWeighScale.EXCITATION_AC)
      YAPI.Sleep(3000, errmsg)
      sensor.tare()
    End If

    While sensor.isOnline()
      Console.WriteLine("Channel 1: " + Str(sensor.get_currentValue()) + sensor.get_unit())
      Console.WriteLine("  (press Ctrl-C to exit)")
      ySleep(1000, errmsg)
    End While
    yFreeAPI()
    Console.WriteLine("Module not connected (check identification and USB cable)")
  End Sub

End Module
