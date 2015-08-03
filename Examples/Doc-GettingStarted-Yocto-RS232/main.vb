Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(execname + " <serial_number>  <frequency> <dutyCycle>")
    Console.WriteLine(execname + " <logical_name> <frequency> <dutyCycle>")
    Console.WriteLine(execname + " any  <frequency> <dutyCycle>   (use any discovered device)")
    Console.WriteLine("     <frequency>: integer between 1Hz and 1000000Hz")
    Console.WriteLine("     <dutyCycle>: floating point number between 0.0 and 100.0")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub


  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim serialPort As YSerialPort
    Dim line As String


    If (YAPI.RegisterHub("usb", errmsg) <> YAPI.SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      Environment.Exit(0)
    End If
    If (argv.Length > 1) Then
      target = argv(1)
      serialPort = YSerialPort.FindSerialPort(target + ".serialPort")
      If (Not serialPort.isOnline()) Then
        Console.WriteLine("No module connected (check cable)")
        Environment.Exit(0)
      End If
    Else
      serialPort = YSerialPort.FirstSerialPort()
      If (serialPort Is Nothing) Then
        Console.WriteLine("No module connected (check USB cable)")
        Environment.Exit(0)
      End If
    End If
    serialport.set_serialMode("9600,8N1")
    serialport.set_protocol("Line")
    serialport.reset()
    Do
      YAPI.Sleep(500, errmsg)
      Do
        line = serialPort.readLine()
        If (line <> "") Then
          Console.WriteLine("Received: " + line)
        End If
      Loop While (line <> "")
      Console.WriteLine("Type line to send, or Ctrl-C to exit: ")
      line = Console.ReadLine()
      serialPort.writeLine(line)
    Loop While (line <> "")
    YAPI.FreeAPI()
  End Sub

End Module
