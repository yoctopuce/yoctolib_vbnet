' ********************************************************************
'
'  $Id: main.vb 38840 2019-12-19 10:23:04Z seb $
'
'  An example that show how to use a  Yocto-SPI
'
'  You can find more information on our web site:
'   Yocto-SPI documentation:
'      https://www.yoctopuce.com/EN/products/yocto-spi/doc.html
'   VB .NET API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim ex = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage")
    Console.WriteLine(ex + " <serial_number>  <frequency> <dutyCycle>")
    Console.WriteLine(ex + " <logical_name> <frequency> <dutyCycle>")
    Console.WriteLine(ex + " any  <frequency> <dutyCycle>   (use any discovered device)")
    Console.WriteLine("     <frequency>: integer between 1Hz and 1000000Hz")
    Console.WriteLine("     <dutyCycle>: floating point number between 0.0 and 100.0")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim value As Integer
    Dim spiPort As YSpiPort
    Dim i, digit As Integer

    If argv.Length < 2 Then Usage()

    target = argv(1)
    value = CInt(argv(2))

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      spiPort = YSpiPort.FirstSpiPort()
      If spiPort Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = spiPort.get_module().get_serialNumber()
    End If
    spiPort = YSpiPort.FindSpiPort(target + ".spiPort")

    If (spiPort.isOnline()) Then
      spiPort.set_spiMode("250000,3,msb")
      spiPort.set_ssPolarity(YSpiPort.SSPOLARITY_ACTIVE_LOW)
      spiPort.set_protocol("Frame:5ms")
      spiPort.reset()
      REM do not forget to configure the powerOutput of the Yocto-SPI
      REM ( for SPI7SEGDISP8.56 powerOutput need to be set at 5v )
      Console.WriteLine("****************************")
      Console.WriteLine("* make sure voltage levels *")
      Console.WriteLine("* are properly configured  *")
      Console.WriteLine("****************************")

      spiPort.writeHex("0c01") REM -- Exit from shutdown state
      spiPort.writeHex("09ff") REM -- Enable BCD for all digits
      spiPort.writeHex("0b07") REM -- Enable digits 0-7 (=8 in total)
      spiPort.writeHex("0a0a") REM -- Set medium brightness
      For i = 1 To 8
        digit = value Mod 10
        spiPort.writeArray(New List(Of Integer)(New Integer() {i, digit}))
        value = value \ 10
      Next
    Else
      Console.WriteLine("Module not connected (check identification and USB cable)")
    End If
    YAPI.FreeAPI()
  End Sub

End Module