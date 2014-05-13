Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(execname + " <serial_number>  <frequency> <dutyCycle>")
    Console.WriteLine(execname + " <logical_name> <frequency> <dutyCycle>")
    Console.WriteLine(execname + " any  <frequency> <dutyCycle>   (use any discovered device)")
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
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      pwmoutput1 = yFirstPwmOutput()
      If pwmoutput1 Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = pwmoutput1.get_Module().get_serialNumber()
    End If
    pwmoutput1 = yFindPwmOutput(target + ".pwmOutput1")
    pwmoutput2 = yFindPwmOutput(target + ".pwmOutput2")

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
  End Sub

End Module
