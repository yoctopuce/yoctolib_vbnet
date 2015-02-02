Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + "  <serial_number>  frequency")
    Console.WriteLine(execname + "  <logical_name> frequency")
    Console.WriteLine(execname + "  any frequency")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""

    Dim buz As YBuzzer
    Dim led, led1, led2 As YLed
    Dim button1, button2 As YAnButton
    Dim target, serial As String
    Dim b1, b2 As Boolean
    Dim freq As Integer

    If argv.Length < 2 Then Usage()
    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      buz = yFirstBuzzer()
      If buz Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
    Else
      buz = yFindBuzzer(target + ".buzzer")
    End If

    If (buz.isOnline()) Then
      Console.WriteLine("press any test button or hit Ctrl-C")
      serial = buz.get_module().get_serialNumber()
      led1 = yFindLed(serial + ".led1")
      led2 = yFindLed(serial + ".led2")
      button1 = yFindAnButton(serial + ".anButton1")
      button2 = yFindAnButton(serial + ".anButton2")

      While (True)
        b1 = button1.get_isPressed()
        b2 = button2.get_isPressed()
        If b1 Or b2 Then
          If b1 Then
            led = led1
            freq = 1500
          Else
            led = led2
            freq = 750
          End If

          led.set_power(Y_POWER_ON)
          led.set_luminosity(100)
          led.set_blinking(Y_BLINKING_PANIC)
          For i = 0 To 4    REM this can be done using sequence as well
            buz.set_frequency(freq)
            buz.freqMove(2 * freq, 250)
            ySleep(250, errmsg)
          Next i
          buz.set_frequency(0)
          led.set_power(Y_POWER_OFF)
        End If
      End While

    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
  End Sub

End Module
