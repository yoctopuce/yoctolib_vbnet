' ********************************************************************
'
'  $Id: main.vb 38840 2019-12-19 10:23:04Z seb $
'
'  An example that show how to use a  Yocto-4-20mA-Tx
'
'  You can find more information on our web site:
'   Yocto-4-20mA-Tx documentation:
'      https://www.yoctopuce.com/EN/products/yocto-4-20ma-tx/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(execname + " <serial_number> <value>")
    Console.WriteLine(execname + " <logical_name>  <value>")
    Console.WriteLine(execname + " any  <value>   (use any discovered device)")
    Console.WriteLine("<value>: floating point number between 4.00 and 20.00 mA")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim currentloop As YCurrentLoopOutput
    Dim pwr As Integer
    Dim value As Double

    If argv.Length < 2 Then Usage()

    target = argv(1)
    value = CDbl(argv(2))

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      currentloop = YCurrentLoopOutput.FirstCurrentLoopOutput()
      If currentloop Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = currentloop.get_module().get_serialNumber()
    End If
    currentloop = YCurrentLoopOutput.FindCurrentLoopOutput(target + ".currentLoopOutput")

    If (currentloop.isOnline()) Then
      currentloop.set_current(value)
      pwr = currentloop.get_loopPower()

      If (pwr = Y_LOOPPOWER_NOPWR) Then
        Console.WriteLine("Current loop is not powered")
        End
      End If
      If (pwr = Y_LOOPPOWER_NOPWR) Then
        Console.WriteLine("Insufficient voltage on current loop")
        End
      End If

      Console.WriteLine("Current in loop  set to " + value.ToString() + " mA")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
