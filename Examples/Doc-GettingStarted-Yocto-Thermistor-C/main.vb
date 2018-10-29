' ********************************************************************
'
'  $Id: main.vb 32622 2018-10-10 13:11:04Z seb $
'
'  An example that show how to use a  Yocto-Thermistor-C
'
'  You can find more information on our web site:
'   Yocto-Thermistor-C documentation:
'      https://www.yoctopuce.com/EN/products/yocto-thermistor-c/doc.html
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
    Dim target, serial As String

    Dim tsensor, ch1, ch2, ch3, ch4, ch5, ch6 As YTemperature

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      tsensor = yFirstTemperature()
      If tsensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      Console.WriteLine("using:" + tsensor.get_module().get_serialNumber())
    Else
      tsensor = yFindTemperature(target + ".temperature1")
    End If

    REM retreive module serial number
    serial = tsensor.get_module().get_serialNumber()

    REM retreive all channels
    ch1 = yFindTemperature(serial + ".temperature1")
    ch2 = yFindTemperature(serial + ".temperature2")
    ch3 = yFindTemperature(serial + ".temperature3")
    ch4 = yFindTemperature(serial + ".temperature4")
    ch5 = yFindTemperature(serial + ".temperature5")
    ch6 = yFindTemperature(serial + ".temperature6")

    While (True)
      If Not (tsensor.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If
      Console.Write("| 1: " + ch1.get_currentValue().ToString(" 0.0"))
      Console.Write("| 2: " + ch2.get_currentValue().ToString(" 0.0"))
      Console.Write("| 3: " + ch3.get_currentValue().ToString(" 0.0"))
      Console.Write("| 4: " + ch4.get_currentValue().ToString(" 0.0"))
      Console.Write("| 5: " + ch5.get_currentValue().ToString(" 0.0"))
      Console.Write("| 6: " + ch6.get_currentValue().ToString(" 0.0"))
      Console.WriteLine("| deg C |")
      ySleep(1000, errmsg)
    End While
    yFreeAPI()
  End Sub

End Module
