' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-PWM-Rx
'
'  You can find more information on our web site:
'   Yocto-PWM-Rx documentation:
'      https://www.yoctopuce.com/EN/products/yocto-pwm-rx/doc.html
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

  Sub Die(ByVal msg As String)
    Console.WriteLine(msg + "(check USB cable)")
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim pwm As YPwmInput
    Dim pwm1 As YPwmInput = Nothing
    Dim pwm2 As YPwmInput = Nothing
    Dim m As YModule = Nothing

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      REM  retreive any pwm input available
      pwm = YPwmInput.FirstPwmInput()
      If pwm Is Nothing Then Die("No module connected")
    Else
      REM retreive the first pwm input from the device given on command line
      pwm = YPwmInput.FindPwmInput(target + ".pwminput1")
    End If

    REM we need to retreive both channels from the device.
    If (pwm.isOnline()) Then
      m = pwm.get_module()
      pwm1 = YPwmInput.FindPwmInput(m.get_serialNumber() + ".pwmInput1")
      pwm2 = YPwmInput.FindPwmInput(m.get_serialNumber() + ".pwmInput2")
    Else
      Die("Module not connected")
    End If

    While (m.isOnline())
      Console.WriteLine("PWM1: " + pwm1.get_frequency().ToString() + "Hz " _
                                 + pwm1.get_dutyCycle().ToString() + "% " _
                                 + pwm1.get_pulseCounter().ToString() _
                                 + " pulse edges")
      Console.WriteLine("PWM2: " + pwm2.get_frequency().ToString() + "Hz " _
                                 + pwm2.get_dutyCycle().ToString() + "% " _
                                 + pwm2.get_pulseCounter().ToString() _
                                 + " pulse edges")
      Console.WriteLine("  (press Ctrl-C to exit)")
      YAPI.Sleep(1000, errmsg)
    End While
    YAPI.FreeAPI()
    Die("Module not connected")
  End Sub
End Module
