' ********************************************************************
'
'  $Id: main.vb 32622 2018-10-10 13:11:04Z seb $
'
'  An example that show how to use a  Yocto-MaxiCoupler
'
'  You can find more information on our web site:
'   Yocto-MaxiCoupler documentation:
'      https://www.yoctopuce.com/EN/products/yocto-maxicoupler/doc.html
'   VB .NET API Reference:
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
    Dim state As Char

    If argv.Length < 3 Then Usage()

    target = argv(1)
    channel = argv(2)
    state = CChar(Mid(argv(3), 1, 1).ToUpper())

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      relay = yFirstRelay()
      If relay Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = relay.get_module().get_serialNumber()
    End If

    Console.WriteLine("using " + target)
    relay = yFindRelay(target + ".relay" + channel)

    If (relay.isOnline()) Then
      If state = "ON" Then
        relay.set_output(Y_OUTPUT_ON)
      Else
        relay.set_output(Y_OUTPUT_OFF)
      End If
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    yFreeAPI()
  End Sub

End Module
