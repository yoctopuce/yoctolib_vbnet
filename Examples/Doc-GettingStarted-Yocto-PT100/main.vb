' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-PT100
'
'  You can find more information on our web site:
'   Yocto-PT100 documentation:
'      https://www.yoctopuce.com/EN/products/yocto-pt100/doc.html
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
    Dim target As String
    Dim tsensor As YTemperature

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      tsensor = YTemperature.FirstTemperature()

      If tsensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
    Else
      tsensor = YTemperature.FindTemperature(target + ".temperature")
    End If

    While (True)
      If Not (tsensor.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If
      Console.WriteLine("Current temperature: " + Str(tsensor.get_currentValue()) _
                        + " °C")
      Console.WriteLine("  (press Ctrl-C to exit)")
      YAPI.Sleep(1000, errmsg)
    End While
    YAPI.FreeAPI()
  End Sub

End Module
