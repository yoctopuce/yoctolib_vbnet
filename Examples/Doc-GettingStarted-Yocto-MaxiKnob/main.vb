' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-MaxiKnob
'
'  You can find more information on our web site:
'   Yocto-MaxiKnob documentation:
'      https://www.yoctopuce.com/EN/products/yocto-maxiknob/doc.html
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
    Dim leds As YColorLedCluster
    Dim button1 As YAnButton
    Dim qd As YQuadratureDecoder
    Dim target, serial As String
    Dim lastpos As Integer

    If argv.Length < 2 Then Usage()
    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      buz = YBuzzer.FirstBuzzer()
      If buz Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
    Else
      buz = YBuzzer.FindBuzzer(target + ".buzzer")
    End If

    If (Not buz.isOnline()) Then
      Console.WriteLine("No module connected (check identification and USB cable)")
      End
    End If

    serial = buz.get_module().get_serialNumber()
    leds = YColorLedCluster.FindColorLedCluster(serial + ".colorLed")
    button1 = YAnButton.FindAnButton(serial + ".anButton1")
    qd = YQuadratureDecoder.FindQuadratureDecoder(serial + ".quadratureDecoder1")

    If (Not button1.isOnline() Or Not qd.isOnline()) Then
      Console.WriteLine("Make sure the Yocto-MaxiBuzzer is configured with at least one anButton and one quadrature Decoder")
      End
    End If

    Console.WriteLine("press a test button, or turn the encoder or hit Ctrl-C")
    lastpos = CType(qd.get_currentValue(), Integer)
    buz.set_volume(75)
    While (button1.isOnline())

      If (button1.get_isPressed() = YAnButton.ISPRESSED_TRUE And lastpos <> 0) Then
        lastpos = 0
        qd.set_currentValue(0)
        buz.playNotes("'E32 C8")
        leds.set_rgbColor(0, 1, 0)
      Else
        Dim p As Integer = CType(qd.get_currentValue(), Integer)
        If (lastpos <> p) Then
          lastpos = p
          buz.pulse(notefreq(p), 500)
          leds.set_hslColor(0, 1, &HFF7F Or (p Mod 255) << 16)
        End If
      End If

    End While
    YAPI.FreeAPI()
  End Sub

  Private Function notefreq(note As Integer) As Integer
    Return CType((220.0 * Math.Exp(note * Math.Log(2) / 12)), Integer)
  End Function
End Module
