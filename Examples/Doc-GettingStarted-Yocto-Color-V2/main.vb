' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-Color-V2
'
'  You can find more information on our web site:
'   Yocto-Color-V2 documentation:
'      https://www.yoctopuce.com/EN/products/yocto-color-v2/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

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
        Dim ledCluster As YColorLedCluster
        Dim color_str As String
        Dim color As Integer
        DIM nb_leds As Integer

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
        If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
            Console.WriteLine("RegisterHub error: " + errmsg)
            End
        End If

        If target = "any" Then
            ledCluster = YColorLedCluster.FirstColorLedCluster()
            If ledCluster Is Nothing Then
                Console.WriteLine("No module connected (check USB cable) ")
                End
            End If
        Else
            ledCluster = YColorLedCluster.FindColorLedCluster(target + ".colorLedCluster")
        End If

        If (ledCluster.isOnline()) Then
            REM configure led cluster
            nb_leds = 2
            ledCluster.set_activeLedCount(nb_leds)
            ledCluster.set_ledType(Y_LEDTYPE_RGB)

            REM immediate transition for fist half of leds
            ledCluster.set_rgbColor(0, nb_leds\2, color)
            REM immediate transition for second half of leds
            ledCluster.rgb_move(nb_leds\2, nb_leds\2, color, 2000)

        Else
            Console.WriteLine("Module not connected (check identification and USB cable)")
        End If
        YAPI.FreeAPI()
    End Sub
End Module
