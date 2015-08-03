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

  REM this is the recusive function to draw 1/3nd of the Von Koch flake
  Private Sub recursiveLine(layer As YDisplayLayer, x0 As Double, y0 As Double, x1 As Double, y1 As Double, deep As Integer)

    Dim dx, dy, mx, my As Double
    If (deep <= 0) Then
      layer.moveTo(CInt(x0 + 0.5), CInt(y0 + 0.5))
      layer.lineTo(CInt(x1 + 0.5), CInt(y1 + 0.5))
    Else
      dx = (x1 - x0) / 3
      dy = (y1 - y0) / 3
      mx = ((x0 + x1) / 2) + (0.87 * (y1 - y0) / 3)
      my = ((y0 + y1) / 2) - (0.87 * (x1 - x0) / 3)
      recursiveLine(layer, x0, y0, x0 + dx, y0 + dy, deep - 1)
      recursiveLine(layer, x0 + dx, y0 + dy, mx, my, deep - 1)
      recursiveLine(layer, mx, my, x1 - dx, y1 - dy, deep - 1)
      recursiveLine(layer, x1 - dx, y1 - dy, x1, y1, deep - 1)
    End If
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim disp As YDisplay
    Dim l1, l2 As YDisplayLayer



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

    REM display clean up
    disp.resetAll()

    l1 = disp.get_displayLayer(1)
    l2 = disp.get_displayLayer(2)
    l1.hide()    REM L1 is hidden, l2 stay visible
    Dim centerX As Double = disp.get_displayWidth() / 2
    Dim centerY As Double = disp.get_displayHeight() / 2
    Dim radius As Double = disp.get_displayHeight() / 2
    Dim a As Double = 0
    Dim i As Integer
    While (True)

      REM we draw in the hidden layer
      l1.clear()
      For i = 0 To 2
        recursiveLine(l1, centerX + radius * Math.Cos(a + i * 2.094),
                         centerY + radius * Math.Sin(a + i * 2.094),
                         centerX + radius * Math.Cos(a + (i + 1) * 2.094),
                         centerY + radius * Math.Sin(a + (i + 1) * 2.094), 2)
      Next i
      REM then we swap contents with the visible layer

      disp.swapLayerContent(1, 2)
      REM change the flake angle
      a += 0.1257
    End While

  End Sub

End Module
