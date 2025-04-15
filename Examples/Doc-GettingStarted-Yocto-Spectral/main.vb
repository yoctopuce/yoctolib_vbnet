' ********************************************************************
'
'  $Id: main.vb 65694 2025-04-09 08:11:25Z tiago $
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
Imports System
Imports System.Threading
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
    Dim colorSensor As YColorSensor

    If argv.Length < 1 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      colorSensor = YColorSensor.FirstColorSensor()
      If colorSensor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = colorSensor.get_module().get_serialNumber()
    End If
    colorSensor = YColorSensor.FindColorSensor(target + ".colorSensor")

    If (colorSensor.isOnline()) Then
      colorSensor.set_workingMode(Y_WORKINGMODE_AUTO)
      colorSensor.set_estimationModel(Y_ESTIMATIONMODEL_REFLECTION)


      While colorSensor.isOnline()
        Console.WriteLine("Currente Color : " + colorSensor.get_nearSimpleColor())
        Console.WriteLine("RGB HEX : #" + colorSensor.get_estimatedRGB().ToString("x"))
        Console.WriteLine("---------------------------------")
        System.Threading.Thread.Sleep(5000)
      End While
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    REM wait 5 sec to show the output


    YAPI.FreeAPI()
  End Sub
End Module