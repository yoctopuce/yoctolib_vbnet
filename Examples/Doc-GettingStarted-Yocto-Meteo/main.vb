' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-Meteo
'
'  You can find more information on our web site:
'   Yocto-Meteo documentation:
'      https://www.yoctopuce.com/EN/products/yocto-meteo/doc.html
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
    Dim hsensor As YHumidity
    Dim tsensor As YTemperature
    Dim psensor As YPressure

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      hsensor = YHumidity.FirstHumidity()
      tsensor = YTemperature.FirstTemperature()
      psensor = YPressure.FirstPressure()

      If hsensor Is Nothing Or tsensor Is Nothing Or psensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
    Else
      hsensor = YHumidity.FindHumidity(target + ".humidity")
      tsensor = YTemperature.FindTemperature(target + ".temperature")
      psensor = YPressure.FindPressure(target + ".pressure")
    End If

    While (True)
      If Not (hsensor.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If
      Console.WriteLine("Current humidity:    " + Str(hsensor.get_currentValue()) _
                        + " %RH")
      Console.WriteLine("Current temperature: " + Str(tsensor.get_currentValue()) _
                        + " °C")
      Console.WriteLine("Current pressure:    " + Str(psensor.get_currentValue()) _
                        + " hPa")
      Console.WriteLine("  (press Ctrl-C to exit)")
      YAPI.Sleep(1000, errmsg)
    End While
    YAPI.FreeAPI()

  End Sub

End Module
