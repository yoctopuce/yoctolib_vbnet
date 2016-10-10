Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + " <serial_number>  [ color | rgb ]")
    Console.WriteLine(execname + " <logical_name> [ color | rgb ]")
    Console.WriteLine(execname + "  any  [ color | rgb ] ")
    Console.WriteLine("Eg.")
    Console.WriteLine(execname + " any FF1493 ")
    Console.WriteLine(execname + " YRGBLED1-123456 red")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim led1 As YColorLed
    Dim led2 As YColorLed
    Dim color_str As String
    Dim color As Integer

    If argv.Length < 3 Then Usage()

    target = argv(1)
    color_str = argv(2).ToUpper()

    If (color_str = "RED") Then
      color = &HFF0000
    ElseIf (color_str = "GREEN") Then
      color = &HFF00
    ElseIf (color_str = "BLUE") Then
      color = &HFF
    Else
      color = CInt(Val("&H" + color_str))
    End If



    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      led1 = yFirstColorLed()
      If led1 Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      led2 = led1.nextColorLed()
    Else
      led1 = yFindColorLed(target + ".colorLed1")
      led2 = yFindColorLed(target + ".colorLed2")
    End If

    If (led1.isOnline()) Then
      led1.set_rgbColor(color) REM immediate switch
      led2.rgbMove(color, 1000) REM smooth transition

    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    yFreeAPI()
  End Sub

End Module
