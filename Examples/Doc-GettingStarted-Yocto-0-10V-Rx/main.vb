' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-0-10V-Rx
'
'  You can find more information on our web site:
'   Yocto-0-10V-Rx documentation:
'      https://www.yoctopuce.com/EN/products/yocto-0-10v-rx/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

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
    Dim sensor, ch1, ch2 As YGenericSensor

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      sensor = YGenericSensor.FirstGenericSensor()
      If sensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      Console.WriteLine("using:" + sensor.get_module().get_serialNumber())
    Else
      sensor = YGenericSensor.FindGenericSensor(target + ".genericSensor1")
    End If

    REM retreive module serial number
    serial = sensor.get_module().get_serialNumber()

    REM retreive both channels
    ch1 = YGenericSensor.FindGenericSensor(serial + ".genericSensor1")
    ch2 = YGenericSensor.FindGenericSensor(serial + ".genericSensor2")
    While (ch1.isOnline() And ch2.isOnline())
      Console.Write("Channel 1: " + Str(ch1.get_currentValue()) + ch1.get_unit())
      Console.Write("  Channel 2: " + Str(ch2.get_currentValue()) + ch1.get_unit())
      Console.WriteLine("  (press Ctrl-C to exit)")
      YAPI.Sleep(1000, errmsg)
    End While
    YAPI.FreeAPI()
    Console.WriteLine("Module not connected (check identification and USB cable)")
  End Sub

End Module
