' ********************************************************************
'
'  $Id: main.vb 75107 2026-07-03 17:05:23Z mvuilleu $
'
'  An example that shows how to use a  Yocto-Display-ePaper
'
'  You can find more information on our web site:
'   Yocto-Display-ePaper documentation:
'      https://www.yoctopuce.com/EN/products/yocto-display-epaper/doc.html
'   Visual Basic .Net API Reference:
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

    If argv.Length <= 1 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      System.Threading.Thread.Sleep(2500)
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

    Dim colors() As Integer = {&HFFFFFF, &H0, &HFF0000, &HFFFF00}

    REM Makes sure the Panel type Is set
    Dim paneltype As String = disp.get_displayPanel()
    If paneltype = "NOT_SET" Then
      Console.WriteLine("Use the virtual to Configure the panel first")
      YAPI.Sleep(3000, errmsg)
      Environment.Exit(0)
    End If

    REM retrieve the display size
    Dim w As Integer = disp.get_displayWidth()
    Dim h As Integer = disp.get_displayHeight()
    Dim middleX As Integer = CInt(w / 2)
    Dim middleY As Integer = CInt(h / 2)

    Console.WriteLine("Using device " + disp.get_serialNumber() + " (panel: " + paneltype + " " + w.ToString() + "x" + h.ToString() + "pixels)")
    disp.resetAll()

    REM forces a full the next refresh only the 1rst time
    disp.regenerateDisplay()

    REM retrieve the first layer
    Dim l0 As YDisplayLayer = disp.get_displayLayer(0)
    l0.selectFont("medium.yfm")
    Dim animation As Boolean = True

    While animation
      REM prevent refreshing for 2 sec
      disp.postponeRefresh(2000)
      l0.clear()
      REM draw a few circle
      For i As Integer = 1 To 15
        Dim cx As Integer = CInt(w * Rnd())
        Dim cy As Integer = CInt(h * Rnd())
        Dim r As Integer = CInt(h / 20) + CInt(Rnd() * h / 10)
        Dim c As Integer = CInt(3 * Rnd())
        l0.selectFillColor(colors(c))
        l0.drawDisc(cx, cy, r)
        l0.drawCircle(cx, cy, r)
      Next i
      REM draw a rectangle with panel type in it
      l0.selectFillColor(&HFFFFFF)
      l0.drawBar(middleX - 75, middleY - 10, middleX + 75, middleY + 12)
      l0.drawRect(middleX - 75, middleY - 10, middleX + 75, middleY + 12)
      l0.drawText(middleX, middleY, Y_ALIGN.CENTER, paneltype)
      disp.triggerRefresh() REM  display is allowed To refresh  again
      YAPI.Sleep(1000, errmsg)
      REM  if no fast refresh available, don't even try to run animations
      If paneltype.IndexOf("KS") < 0 Then animation = False
    End While
    YAPI.FreeAPI()


  End Sub

End Module
