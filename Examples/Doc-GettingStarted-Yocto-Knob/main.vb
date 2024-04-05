' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-Knob
'
'  You can find more information on our web site:
'   Yocto-Knob documentation:
'      https://www.yoctopuce.com/EN/products/yocto-knob/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + " <serial_number>")
    Console.WriteLine(execname + " <logical_name>")
    Console.WriteLine(execname + " any  ")
    System.Threading.Thread.Sleep(2500)

    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim input1 As YAnButton = Nothing
    Dim input5 As YAnButton = Nothing

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      input1 = YAnButton.FirstAnButton()
      If input1 Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = input1.get_Module().get_serialNumber()
    End If

    input1 = YAnButton.FindAnButton(target + ".anButton1")
    input5 = YAnButton.FindAnButton(target + ".anButton5")

    While (True)
      If Not (input1.isOnline()) Then
        Console.WriteLine("Module not connected (check identification and USB cable)")
        End
      End If
      If (input1.get_isPressed() = Y_ISPRESSED_TRUE) Then
        Console.Write("Button1: pressed      ")
      Else
        Console.Write("Button1: not pressed  ")
      End If
      Console.WriteLine("- analog value:  " + Str(input1.get_calibratedValue()))

      If (input5.get_isPressed() = Y_ISPRESSED_TRUE) Then
        Console.Write("Button5: pressed      ")
      Else
        Console.Write("Button5: not pressed  ")
      End If
      Console.WriteLine("- analog value:  " + Str(input5.get_calibratedValue()))

      YAPI.Sleep(1000, errmsg)

    End While
    YAPI.FreeAPI()

  End Sub

End Module
