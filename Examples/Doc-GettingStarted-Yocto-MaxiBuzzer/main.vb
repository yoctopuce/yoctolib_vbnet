' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-MaxiBuzzer
'
'  You can find more information on our web site:
'   Yocto-MaxiBuzzer documentation:
'      https://www.yoctopuce.com/EN/products/yocto-maxibuzzer/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

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
    Dim led As YColorLed
    Dim button1, button2 As YAnButton
    Dim target, serial As String
    Dim b1, b2 As Boolean
    Dim freq As Integer
    Dim vol As Integer
    Dim color As Integer

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
      led = yFindColorLed(serial + ".colorLed")
      button1 = yFindAnButton(serial + ".anButton1")
      button2 = yFindAnButton(serial + ".anButton2")

      While (True)
        b1 = button1.get_isPressed() = YAnButton.ISPRESSED_TRUE
        b2 = button2.get_isPressed() = YAnButton.ISPRESSED_TRUE
        If b1 Or b2 Then
          If b1 Then
            freq = 1500
            vol = 60
            color = &hff0000
          Else
            freq = 750
            vol = 30
            color = &h00ff00
          End If

          led.resetBlinkSeq()
          led.addRgbMoveToBlinkSeq(color, 100)
          led.addRgbMoveToBlinkSeq(0, 100)
          led.startBlinkSeq()
          buz.set_volume(vol)
          For i = 0 To 4 REM this can be done using sequence as well
            buz.set_frequency(freq)
            buz.freqMove(2*freq, 250)
            ySleep(250, errmsg)
          Next i
          buz.set_frequency(0)
          led.stopBlinkSeq()
          led.set_rgbColor(0)
        End If
      End While
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    yFreeAPI()
  End Sub
End Module
