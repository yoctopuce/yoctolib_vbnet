Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(execname + " <serial_number> <value>")
    Console.WriteLine(execname + " <logical_name>  <value>")
    Console.WriteLine(execname + " any  <value>   (use any discovered device)")
    Console.WriteLine("<value>: floating point number between 4.00 and 20.00 mA")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub


  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim currentloop As YCurrentLoopOutput
    Dim pwr As Integer
    Dim value As Double

    If argv.Length < 2 Then Usage()

    target = argv(1)
    value = CDbl(argv(2))

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      currentloop = yFirstCurrentLoopOutput()
      If currentloop Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = currentloop.get_module().get_serialNumber()
    End If
    currentloop = yFindCurrentLoopOutput(target + ".currentLoopOutput")

    If (currentloop.isOnline()) Then
      currentloop.set_current(value)
      pwr = currentloop.get_loopPower()

      If (pwr = Y_LOOPPOWER_NOPWR) Then
        Console.WriteLine("Current loop is not powered")
        End
      End If
      If (pwr = Y_LOOPPOWER_NOPWR) Then
        Console.WriteLine("Insufficient voltage on current loop")
        End
      End If

      Console.WriteLine("Current in loop  set to " + value.ToString() + " mA")
    End If
    yFreeAPI()
  End Sub

End Module
