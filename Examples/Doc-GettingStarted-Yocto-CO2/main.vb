' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-CO2
'
'  You can find more information on our web site:
'   Yocto-CO2 documentation:
'      https://www.yoctopuce.com/EN/products/yocto-co2/doc.html
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
    Dim tsensor As YCarbonDioxide

    If argv.Length < 2 Then Usage()

    target = argv(1)
    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      tsensor = YCarbonDioxide.FirstCarbonDioxide()
      If tsensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      Console.WriteLine("using " + tsensor.get_module().get_serialNumber())
    Else
      tsensor = YCarbonDioxide.FindCarbonDioxide(target + ".carbonDioxide")
    End If

    While (True)
      If Not (tsensor.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If

      Console.WriteLine("CO2: " + Str(tsensor.get_currentValue()) + " ppm")
      Console.WriteLine("  (press Ctrl-C to exit)")
      YAPI.Sleep(1000, errmsg)

    End While
    YAPI.FreeAPI()

  End Sub

End Module
