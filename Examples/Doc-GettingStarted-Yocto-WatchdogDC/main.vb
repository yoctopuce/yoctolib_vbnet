' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-WatchdogDC
'
'  You can find more information on our web site:
'   Yocto-WatchdogDC documentation:
'      https://www.yoctopuce.com/EN/products/yocto-watchdogdc/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Console.WriteLine("demo  <serial_number>  [ on | off | reset]")
    Console.WriteLine("demo  <logical_name>  [ on | off | reset]")
    Console.WriteLine("demo  any [ on | off | reset]")
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim watchdog As YWatchdog
    Dim state As String

    If argv.Length < 3 Then Usage()

    target = argv(1)
    state = argv(2).ToUpper()

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      watchdog = YWatchdog.FirstWatchdog()
      If watchdog Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
    Else
      watchdog = YWatchdog.FindWatchdog(target + ".watchdog1")
    End If

    If (watchdog.isOnline()) Then
      If state = "ON" Then watchdog.set_running(Y_RUNNING_ON)
      If state = "OFF" Then watchdog.set_running(Y_RUNNING_OFF)
      If state = "RESET" Then watchdog.resetWatchdog()
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
