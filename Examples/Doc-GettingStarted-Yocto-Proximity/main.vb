' ********************************************************************
'
'  $Id: main.vb 32622 2018-10-10 13:11:04Z seb $
'
'  An example that show how to use a  Yocto-Proximity
'
'  You can find more information on our web site:
'   Yocto-Proximity documentation:
'      https://www.yoctopuce.com/EN/products/yocto-proximity/doc.html
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
    Dim p As YProximity

    Dim al, ir As YLightSensor

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      p = yFirstProximity()
      If p Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = p.get_module().get_serialNumber()

    Else
      p = yFindProximity(target + ".proximity1")
    End If

    al = yFindLightSensor(target + ".lightSensor1")
    ir = yFindLightSensor(target + ".lightSensor2")

    While (True)
      If Not (p.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If
      Console.Write("proximity: " + Str(p.get_currentValue()))
      Console.Write(" ambiant: " + Str(al.get_currentValue()))
      Console.Write(" ir: " + Str(ir.get_currentValue()))

      Console.WriteLine("  (press Ctrl-C to exit)")
      ySleep(1000, errmsg)

    End While
    yFreeAPI()

  End Sub

End Module
