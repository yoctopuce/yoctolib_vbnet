' ********************************************************************
'
'  $Id: main.vb 32622 2018-10-10 13:11:04Z seb $
'
'  An example that show how to use a  Yocto-RangeFinder
'
'  You can find more information on our web site:
'   Yocto-RangeFinder documentation:
'      https://www.yoctopuce.com/EN/products/yocto-rangefinder/doc.html
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
    Dim target As String
    Dim rf As YRangeFinder
    Dim ir As YLightSensor
    Dim tmp As YTemperature

    If argv.Length < 2 Then Usage()
    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      rf = yFirstRangeFinder()
      If rf Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = rf.get_module().get_serialNumber()
    Else
      rf = yFindRangeFinder(target + ".rangeFinder1")
    End If

    ir = yFindLightSensor(target + ".lightSensor1")
    tmp = yFindTemperature(target + ".temperature1")

    While (True)
      If Not (rf.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If
      Console.WriteLine("Distance    : " + Str(rf.get_currentValue()))
      Console.WriteLine("Ambiant IR  : " + Str(ir.get_currentValue()))
      Console.WriteLine("Temperature : " + Str(tmp.get_currentValue()))
      Console.WriteLine("  (press Ctrl-C to exit)")
      ySleep(1000, errmsg)
    End While
    yFreeAPI()

  End Sub

End Module
