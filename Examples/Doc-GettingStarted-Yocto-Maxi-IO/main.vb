' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-Maxi-IO
'
'  You can find more information on our web site:
'   Yocto-Maxi-IO documentation:
'      https://www.yoctopuce.com/EN/products/yocto-maxi-io/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + "  <serial_number>")
    Console.WriteLine(execname + "  <logical_name>")
    Console.WriteLine(execname + "  any")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()

    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim io As YDigitalIO
    Dim outputdata As Integer
    Dim inputdata As Integer
    Dim line As String

    If argv.Length < 2 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      io = YDigitalIO.FirstDigitalIO()
      If io Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
    Else
      io = YDigitalIO.FindDigitalIO(target + ".digitalIO")
    End If

    If (Not io.isOnline()) Then
      Console.WriteLine("Module not connected (check identification and USB cable)")
      End
    End If

    REM lets configure the channels direction
    REM bits 0..3 as output
    REM bits 4..7 as input
    io.set_portDirection(&HF)
    io.set_portPolarity(0)  REM polarity set to regular
    io.set_portOpenDrain(0) REM No open drain

    Console.WriteLine("Channels 0..3 are configured as outputs and channels 4..7")
    Console.WriteLine("are configred as inputs, you can connect some inputs to")
    Console.WriteLine("ouputs and see what happens")

    While (io.isOnline())
      inputdata = io.get_portState() REM read port values
      line = ""  REM display part state value as binary
      For i As Integer = 0 To 7 Step 1
        If CBool((inputdata And (128 >> i))) Then
          line = line + "1"
        Else
          line = line + "0"
        End If
      Next
      Console.WriteLine("port value = " + line)
      outputdata = (outputdata + 1) Mod 16 REM cycle ouput 0..15
      io.set_portState(outputdata) REM We could have used set_bitState as well
      YAPI.Sleep(1000, errmsg)
    End While
    Console.WriteLine("Module disconnected")
    YAPI.FreeAPI()
  End Sub

End Module
