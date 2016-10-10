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

    Dim tsensor, ch1, ch2 As YTemperature

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      tsensor = yFirstTemperature()
      If tsensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      Console.WriteLine("using:" + tsensor.get_module().get_serialNumber())
    Else
      tsensor = yFindTemperature(target + ".temperature1")
    End If

    REM retreive module serial number
    serial = tsensor.get_module().get_serialNumber()

    REM retreive both channels
    ch1 = yFindTemperature(serial + ".temperature1")
    ch2 = yFindTemperature(serial + ".temperature2")

    While (True)
      If Not (tsensor.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If
      Console.Write("channel 1: " + Str(ch1.get_currentValue()) + " °C  ")
      Console.Write("channel 2: " + Str(ch2.get_currentValue()) + " °C  ")
      Console.WriteLine("  (press Ctrl-C to exit)")
      ySleep(1000, errmsg)
    End While
    yFreeAPI()
  End Sub

End Module
