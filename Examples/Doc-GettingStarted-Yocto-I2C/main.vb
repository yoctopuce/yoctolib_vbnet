' ********************************************************************
'
'  $Id: svn_id $
'
'  An example that show how to use a  Yocto-I2C
'
'  You can find more information on our web site:
'   Yocto-I2C documentation:
'      https://www.yoctopuce.com/EN/products/yocto-i2c/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim ex = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(ex + " <serial_number>")
    Console.WriteLine(ex + " <logical_name>")
    Console.WriteLine(ex + " any              (use any discovered device)")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim i2cPort As YI2cPort

    If argv.Length < 1 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      i2cPort = yFirstI2cPort()
      If i2cPort Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = i2cPort.get_module().get_serialNumber()
    End If
    i2cPort = yFindI2cPort(target + ".i2cPort")

    If (i2cPort.isOnline()) Then
      i2cPort.set_i2cMode("400kbps")
      i2cPort.set_i2cVoltageLevel(YI2cPort.I2CVOLTAGELEVEL_3V3)
      i2cPort.reset()
      REM do not forget to configure the powerOutput and 
      REM of the Yocto-I2C as well if used
      Console.WriteLine("****************************")
      Console.WriteLine("* make sure voltage levels *")
      Console.WriteLine("* are properly configured  *")
      Console.WriteLine("****************************")

      Dim toSend As List(Of Integer) = New List(Of Integer) From {&H5}
      Dim received As List(Of Integer) = i2cPort.i2cSendAndReceiveArray(&H1F, toSend, 2)
      Dim tempReg As Integer = (received(0) << 8) + received(1)
      If ((tempReg And &H1000) <> 0) Then
        tempReg = tempReg - &H2000 REM perform sign extension
      Else

        tempReg = tempReg And &HFFF REM clear status bits
      End If

      Console.WriteLine("Ambiant temperature: " + String.Format("{0:0.000}", (tempReg / 16.0)))
      Else
        Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    REM wait 5 sec to show the output
    System.Threading.Thread.Sleep(5000)

    YAPI.FreeAPI()
  End Sub

End Module