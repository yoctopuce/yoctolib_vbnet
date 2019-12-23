' ********************************************************************
'
'  $Id: main.vb 38840 2019-12-19 10:23:04Z seb $
'
'  An example that show how to use a  Yocto-milliVolt-Rx
'
'  You can find more information on our web site:
'   Yocto-milliVolt-Rx documentation:
'      https://www.yoctopuce.com/EN/products/yocto-millivolt-rx/doc.html
'   VB .NET API Reference:
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
    Dim sensor, ch1 As YGenericSensor

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

    While (ch1.isOnline())
      Console.Write("Voltage: " + Str(ch1.get_currentValue()) + ch1.get_unit())
      Console.WriteLine("  (press Ctrl-C to exit)")
      YAPI.Sleep(1000, errmsg)
    End While
    Console.WriteLine("Module not connected (check identification and USB cable)")
    YAPI.FreeAPI()
  End Sub

End Module
