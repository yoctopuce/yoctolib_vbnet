' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-MaxiDisplay
'
'  You can find more information on our web site:
'   Yocto-MaxiDisplay documentation:
'      https://www.yoctopuce.com/EN/products/yocto-maxidisplay/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

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
    Dim l0, l1 As YDisplayLayer
    Dim h, w, y, x, vx, vy As Integer

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

    REM Display clean up
    disp.resetAll()
    REM retreive the display size
    w = disp.get_displayWidth()
    h = disp.get_displayHeight()

    REM reteive the first layer
    l0 = disp.get_displayLayer(0)

    REM display a text in the middle of the screen
    l0.drawText(CInt(w / 2), CInt(h / 2), Y_ALIGN.CENTER, "Hello world!")

    REM visualize each corner
    l0.moveTo(0, 5)
    l0.lineTo(0, 0)
    l0.lineTo(5, 0)
    l0.moveTo(0, h - 6)
    l0.lineTo(0, h - 1)
    l0.lineTo(5, h - 1)
    l0.moveTo(w - 1, h - 6)
    l0.lineTo(w - 1, h - 1)
    l0.lineTo(w - 6, h - 1)
    l0.moveTo(w - 1, 5)
    l0.lineTo(w - 1, 0)
    l0.lineTo(w - 6, 0)

    REM draw a circle in the top left corner of layer 1
    l1 = disp.get_displayLayer(1)
    l1.clear()
    l1.drawCircle(CInt(h / 8), CInt(h / 8), CInt(h / 8))

    REM and animate the layer
    Console.WriteLine("Use Ctrl-C to stop")
    x = 0
    y = 0
    vx = 1
    vy = 1

    While (disp.isOnline())
      x += vx
      y += vy
      If ((x < 0) Or (x > w - (h / 4))) Then vx = -vx
      If ((y < 0) Or (y > h - (h / 4))) Then vy = -vy
      l1.setLayerPosition(x, y, 0)
      YAPI.Sleep(5, errmsg)
    End While
    YAPI.FreeAPI()
  End Sub

End Module
