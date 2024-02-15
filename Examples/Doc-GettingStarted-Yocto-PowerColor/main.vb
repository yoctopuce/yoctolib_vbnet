' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-PowerColor
'
'  You can find more information on our web site:
'   Yocto-PowerColor documentation:
'      https://www.yoctopuce.com/EN/products/yocto-powercolor/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Imports System.Reflection
Imports System.IO

Module Module1

  Private Sub Usage()
    Dim errmsg As String = ""
    Dim exe As String = Path.GetFileName(Assembly.GetExecutingAssembly().Location)
    Console.WriteLine("Bad command line arguments")
    Console.WriteLine(exe + " <serial_number>  [ color | rgb ]")
    Console.WriteLine(exe + " <logical_name> [ color | rgb ]")
    Console.WriteLine(exe + " any  [ color | rgb ] ")
    Console.WriteLine("Eg.")
    Console.WriteLine(exe + " any FF1493 ")
    Console.WriteLine(exe + " YRGBHI01-123456 red")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim led1 As YColorLed

    Dim color_str As String
    Dim color As Integer

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

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

    If target = "any" Then
      led1 = YColorLed.FirstColorLed()
      If led1 Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
    Else
      led1 = YColorLed.FindColorLed(target + ".colorLed1")
    End If

    If (led1.isOnline()) Then
      led1.rgbMove(color, 1000) REM smooth transition
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub
End Module
