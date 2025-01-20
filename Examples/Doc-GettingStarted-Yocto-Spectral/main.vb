' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-Spectral
'
'  You can find more information on our web site:
'   Yocto-Spectral documentation:
'      https://www.yoctopuce.com/EN/products/yocto-spectral/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim ex = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(ex + " <serial_number>")
    Console.WriteLine(ex + " <logical_name>")
    Console.WriteLine(ex + " any              (use any discovered device)")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim spectralSensor As YSpectralSensor

    If argv.Length < 1 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      spectralSensor = YSpectralSensor.FirstSpectralSensor()
      If spectralSensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = spectralSensor.get_module().get_serialNumber()
    End If
    spectralSensor = YSpectralSensor.FindSpectralSensor(target + ".spectralSensor")

    If (spectralSensor.isOnline()) Then
      spectralSensor.set_gain(6)
      spectralSensor.set_integrationTime(150)
      spectralSensor.set_ledCurrent(6)

      Console.WriteLine("Currente Color : " + spectralSensor.get_nearSimpleColor())
      Console.WriteLine("Color HEX : #" + spectralSensor.get_estimatedRGB().ToString("x"))
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    REM wait 5 sec to show the output
    System.Threading.Thread.Sleep(5000)

    YAPI.FreeAPI()
  End Sub

End Module