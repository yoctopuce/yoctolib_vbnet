' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-Servo
'
'  You can find more information on our web site:
'   Yocto-Servo documentation:
'      https://www.yoctopuce.com/EN/products/yocto-servo/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + " <serial_number>  [ -1000 | ... | 1000 ]")
    Console.WriteLine(execname + " <logical_name> [ -1000 | ... | 1000 ]")
    Console.WriteLine(execname + " any [ -1000 | ... | 1000 ]")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim servo1 As YServo
    Dim servo5 As YServo
    Dim pos As Integer

    If argv.Length < 3 Then Usage()

    target = argv(1)
    pos = CInt(argv(2))

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      servo1 = YServo.FirstServo()
      If servo1 Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = servo1.get_Module().get_serialNumber()
    End If
    servo1 = YServo.FindServo(target + ".servo1")
    servo5 = YServo.FindServo(target + ".servo5")

    If (servo1.isOnline()) Then
      servo1.set_position(pos)
      servo5.move(pos, 3000)
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
