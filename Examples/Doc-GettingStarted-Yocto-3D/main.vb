' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-3D
'
'  You can find more information on our web site:
'   Yocto-3D documentation:
'      https://www.yoctopuce.com/EN/products/yocto-3d/doc.html
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
    Dim serial As String
    Dim count As Integer
    Dim anytilt, tilt1, tilt2 As YTilt
    Dim compass As YCompass
    Dim accelerometer As YAccelerometer
    Dim gyro As YGyro

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      anytilt = YTilt.FirstTilt()
      If anytilt Is Nothing Then
        Console.WriteLine("No module connected (check USB cable)")
        End
      End If
    Else
      anytilt = YTilt.FindTilt(target + ".tilt1")
      If Not (anytilt.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If
    End If

    serial = anytilt.get_module().get_serialNumber()
    tilt1 = YTilt.FindTilt(serial + ".tilt1")
    tilt2 = YTilt.FindTilt(serial + ".tilt2")
    compass = YCompass.FindCompass(serial + ".compass")
    accelerometer = YAccelerometer.FindAccelerometer(serial + ".accelerometer")
    gyro = YGyro.FindGyro(serial + ".gyro")
    count = 0

    While (True)
      If (Not tilt1.isOnline()) Then
        Console.WriteLine("Module disconnected")
        End
      End If

      If (count Mod 10 = 0) Then
        Console.WriteLine("tilt1" + Chr(9) + "tilt2" + Chr(9) + "compass" _
                          + Chr(9) + "acc" + Chr(9) + "gyro")
      End If
      Console.Write(tilt1.get_currentValue().ToString() + Chr(9))
      Console.Write(tilt2.get_currentValue().ToString() + Chr(9))
      Console.Write(compass.get_currentValue().ToString() + Chr(9))
      Console.Write(accelerometer.get_currentValue().ToString() + Chr(9))
      Console.WriteLine(gyro.get_currentValue().ToString())
      count = count + 1
      YAPI.Sleep(250, errmsg)
    End While
    YAPI.FreeAPI()
  End Sub

End Module
