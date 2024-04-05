' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-MaxiPowerRelay
'
'  You can find more information on our web site:
'   Yocto-MaxiPowerRelay documentation:
'      https://www.yoctopuce.com/EN/products/yocto-maxipowerrelay/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim execname As String = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + " <serial_number> <channel> [ ON | OFF ]")
    Console.WriteLine(execname + " <logical_name>  <channel> [ ON | OFF ]")
    Console.WriteLine(execname + " any <channel> [ ON | OFF ]")
    Console.WriteLine("Example:")
    Console.WriteLine(execname + " any 1 [ ON | OFF ]")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target, channel As String
    Dim relay As YRelay
    Dim state As String

    If argv.Length < 3 Then Usage()

    target = argv(1)
    channel = argv(2)
    state = argv(3).ToUpper

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      relay = YRelay.FirstRelay()
      If relay Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = relay.get_module().get_serialNumber()
    End If

    Console.WriteLine("using " + target)
    relay = YRelay.FindRelay(target + ".relay" + channel)

    If (relay.isOnline()) Then
      If state = "ON" Then
        relay.set_output(Y_OUTPUT_ON)
      Else
        relay.set_output(Y_OUTPUT_OFF)
      End If
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
