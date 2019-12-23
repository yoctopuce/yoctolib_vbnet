' ********************************************************************
'
'  $Id: main.vb 38840 2019-12-19 10:23:04Z seb $
'
'  An example that show how to use a  Yocto-Demo
'
'  You can find more information on our web site:
'   Yocto-Demo documentation:
'      https://www.yoctopuce.com/EN/products/yocto-demo/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + " <serial_number>  [ on | off ]")
    Console.WriteLine(execname + " <logical_name> [ on | off ]")
    Console.WriteLine(execname + " any [ on | off ] ")
    System.Threading.Thread.Sleep(2500)

    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim led As YLed

    Dim on_off As String

    If argv.Length < 3 Then Usage()

    target = argv(1)
    on_off = argv(2).ToUpper()

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      led = YLed.FirstLed()
      If led Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If

    Else
      led = YLed.FindLed(target + ".led")

    End If

    If (led.isOnline()) Then
      If on_off = "ON" Then led.set_power(Y_POWER_ON) Else led.set_power(Y_POWER_OFF)
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
