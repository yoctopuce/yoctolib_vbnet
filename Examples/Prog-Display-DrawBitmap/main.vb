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
    Dim w, h As Integer


    If argv.Length <= 1 Then Usage()

    target = argv(1)
    



    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

        If target = "any" Then
      disp = yFirstDisplay()
      If disp Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If

        Else
      disp = yFindDisplay(target + ".display")

    End If

    If Not (disp.isOnline()) Then
      Console.WriteLine("Module not connected (check identification and USB cable)")
      End
    End If

    REM Display clean up
    disp.resetAll()
    REM retreive the display size
    w = disp.get_displayWidth()
    h = disp.get_displayHeight()

   

    REM reteive the first layer
    l0 = disp.get_displayLayer(0)
    Dim bytesPerLines As Integer = CInt(w / 8)
    Dim i, j As Integer

    Dim Data As Byte()
    ReDim Data(h * bytesPerLines)
    For i = 0 To h * bytesPerLines - 1
      Data(i) = 0
    Next i

    Dim max_iteration As Integer = 50
    Dim iteration, index As Integer
    Dim xtemp As Double
    Dim centerX As Double = 0
    Dim centerY As Double = 0
    Dim targetX As Double = 0.834555980181972
    Dim targetY As Double = 0.204552998862566
    Dim x, y, x0, y0 As Double
    Dim zoom As Double = 1
    Dim distance As Double = 1


    While (True)

      For i = 0 To Data.Count - 1
        Data(i) = 0
      Next i
      distance = distance * 0.95
      centerX = targetX * (1 - distance)
      centerY = targetY * (1 - distance)
      max_iteration = CInt(Math.Round(max_iteration + Math.Sqrt(zoom)))
      If (max_iteration > 1500) Then max_iteration = 1500
      For j = 0 To h - 1
        For i = 0 To w - 1
          x0 = (((i - w / 2.0) / (w / 8)) / zoom) - centerX
          y0 = (((j - h / 2.0) / (w / 8)) / zoom) - centerY
          x = 0
          y = 0
          iteration = 0

          While ((x * x + y * y < 4) And (iteration < max_iteration))

            xtemp = x * x - y * y + x0
            y = 2 * x * y + y0
            x = xtemp
            iteration += 1
          End While

          If (iteration >= max_iteration) Then
            index = j * bytesPerLines + (i >> 3)
            Data(index) = Data(index) Or CByte((128 >> (i Mod 8)))
          End If
        Next i
      Next j


      l0.drawBitmap(0, 0, w, Data, 0)
      zoom = zoom / 0.95
    End While

  End Sub

End Module
