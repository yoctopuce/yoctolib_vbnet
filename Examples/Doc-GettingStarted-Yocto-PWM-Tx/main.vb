' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-PWM-Tx
'
'  You can find more information on our web site:
'   Yocto-PWM-Tx documentation:
'      https://www.yoctopuce.com/EN/products/yocto-pwm-tx/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim ex = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(ex + " <serial_number>  <frequency> <dutyCycle>")
    Console.WriteLine(ex + " <logical_name> <frequency> <dutyCycle>")
    Console.WriteLine(ex + " any  <frequency> <dutyCycle> (use any discovered device)")
    Console.WriteLine("     <frequency>: integer between 1Hz and 1000000Hz")
    Console.WriteLine("     <dutyCycle>: floating point number between 0.0 and 100.0")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim pwmoutput1 As YPwmOutput
    Dim pwmoutput2 As YPwmOutput
    Dim frequency As Integer
    Dim dutyCycle As Double

    If argv.Length < 3 Then Usage()

    target = argv(1)
    frequency = CInt(argv(2))
    dutyCycle = CDbl(argv(3))

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      pwmoutput1 = YPwmOutput.FirstPwmOutput()
      If pwmoutput1 Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = pwmoutput1.get_Module().get_serialNumber()
    End If
    pwmoutput1 = YPwmOutput.FindPwmOutput(target + ".pwmOutput1")
    pwmoutput2 = YPwmOutput.FindPwmOutput(target + ".pwmOutput2")

    If (pwmoutput1.isOnline()) Then
      REM output 1 : immediate change
      pwmoutput1.set_frequency(frequency)
      pwmoutput1.set_enabled(YPwmOutput.ENABLED_TRUE)
      pwmoutput1.set_dutyCycle(dutyCycle)
      REM  output 2 : smooth change
      pwmoutput2.set_frequency(frequency)
      pwmoutput2.set_enabled(YPwmOutput.ENABLED_TRUE)
      pwmoutput2.dutyCycleMove(dutyCycle, 3000)
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
