' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-0-10V-Tx
'
'  You can find more information on our web site:
'   Yocto-0-10V-Tx documentation:
'      https://www.yoctopuce.com/EN/products/yocto-0-10v-tx/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim ex = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(ex + " <serial_number> <voltage>")
    Console.WriteLine(ex + " <logical_name>  <voltage>")
    Console.WriteLine(ex + " any  <voltage>  (use any discovered device)")
    Console.WriteLine("     <voltage>: floating point number between 0.0 and 10.000")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim vout1 As YVoltageOutput
    Dim vout2 As YVoltageOutput
    Dim voltage As Double

    If argv.Length < 2 Then Usage()

    target = argv(1)
    voltage = CDbl(argv(2))

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      vout1 = YVoltageOutput.FirstVoltageOutput()
      If vout1 Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = vout1.get_Module().get_serialNumber()
    End If
    vout1 = YVoltageOutput.FindVoltageOutput(target + ".voltageOutput1")
    vout2 = YVoltageOutput.FindVoltageOutput(target + ".voltageOutput2")

    If (vout1.isOnline()) Then
      REM output 1 : immediate change
      vout1.set_currentVoltage(voltage)
      REM  output 2 : smooth change
      vout2.voltageMove(voltage, 3000)
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
