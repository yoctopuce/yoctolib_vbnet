Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + " <serial_number>")
    Console.WriteLine(execname + " <logical_name> ")
    Console.WriteLine(execname + "  any")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

    Sub Main()
        Dim argv() As String = System.Environment.GetCommandLineArgs()
        Dim errmsg As String = ""
        Dim target As String
    Dim disp As YDisplay
    Dim l0 As YDisplayLayer
    Dim h, w As Integer


    If argv.Length <= 1 Then Usage()

    target = argv(1)




    REM Setup the API to use local USB devices
        If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
            Console.WriteLine("RegisterHub error: " + errmsg)
            End
        End If

        If target = "any" Then
      disp = YDisplay.FirstDisplay()
      If disp Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If

        Else
      disp = YDisplay.FindDisplay(target + ".display")

    End If

    If Not (disp.isOnline()) Then
      Console.WriteLine("Module not connected (check identification and USB cable)")
      End
    End If

    disp.resetAll()
    REM retreive the display size

    w = disp.get_displayWidth()
    h = disp.get_displayHeight()

    REM retreive the first layer
    l0 = disp.get_displayLayer(0)

    Dim count As Integer = 8
    Dim coord(2 * count + 1) As Integer
    Dim i As Integer

    REM precompute the "leds" position
    Dim ledwidth As Integer = CInt(w / count)
    For i = 0 To count - 1

      coord(i) = i * ledwidth
      coord(2 * count - i - 2) = coord(i)
    Next i

    Dim framesCount As Integer = 2 * count - 2

    REM start recording
    disp.newSequence()

    REM build one loop for recording
    For i = 0 To framesCount - 1

      l0.selectColorPen(0)
      l0.drawBar(coord((i + framesCount - 1) Mod framesCount),
                h - 1,
                coord((i + framesCount - 1) Mod framesCount) + ledwidth,
                h - 4)
      l0.selectColorPen(&HFFFFFF)
      l0.drawBar(coord(i), h - 1, coord(i) + ledwidth, h - 4)
      disp.pauseSequence(100) REM  records a 50ms pause.
    Next i
    REM self-call : causes an endless looop
    disp.playSequence("K2000.seq")
    REM stop recording and save to device filesystem
    disp.saveSequence("K2000.seq")

    REM play the sequence
    disp.playSequence("K2000.seq")

    Console.WriteLine("This animation is running in background.")

  End Sub

End Module
