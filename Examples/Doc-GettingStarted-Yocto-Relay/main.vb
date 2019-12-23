' ********************************************************************
'
'  $Id: main.vb 38840 2019-12-19 10:23:04Z seb $
'
'  An example that show how to use a  Yocto-Relay
'
'  You can find more information on our web site:
'   Yocto-Relay documentation:
'      https://www.yoctopuce.com/EN/products/yocto-relay/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + "  <serial_number>  < A | B >")
    Console.WriteLine(execname + "  <logical_name>  < A | B >")
    Console.WriteLine(execname + "  any < A | B >")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim relay As YRelay
    Dim state As Char

    If argv.Length < 3 Then Usage()

    target = argv(1)
    state = CChar(Mid(argv(2), 1, 1).ToUpper())

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      relay = YRelay.FirstRelay()
      If relay Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
    Else
      relay = YRelay.FindRelay(target + ".relay1")
    End If

    If (relay.isOnline()) Then
      If state = "A" Then relay.set_state(Y_STATE_A) Else relay.set_state(Y_STATE_B)
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
