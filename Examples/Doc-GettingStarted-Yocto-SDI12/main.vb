' ********************************************************************
'
'  $Id: main.vb 58233 2023-12-04 10:57:58Z seb $
'
'  An example that shows how to use a  Yocto-SDI12
'
'  You can find more information on our web site:
'   Yocto-SDI12 documentation:
'      https://www.yoctopuce.com/EN/products/yocto-sdi12/doc.html
'   Visual Basic .Net V2 API Reference:
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
        Dim sdi12Port As YSdi12Port

        If argv.Length < 1 Then Usage()

    target = argv(1)

        REM Setup the API to use local USB devices
        If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
            Console.WriteLine("RegisterHub error: " + errmsg)
            End
        End If

        If target = "any" Then
            sdi12Port = YSdi12Port.FirstSdi12Port()
            If sdi12Port Is Nothing Then
                Console.WriteLine("No module connected (check USB cable) ")
                End
            End If
            target = sdi12Port.get_module().get_serialNumber()
        End If
        sdi12Port = YSdi12Port.FindSdi12Port(target + ".sdi12Port")
        While True
            If (sdi12Port.isOnline()) Then
                Console.SetCursorPosition(0, 0)
                Dim singleSensor As YSdi12Sensor = sdi12Port.discoverSingleSensor()
                Console.WriteLine("Sensor address : " + singleSensor.get_sensorAddress())
                Console.WriteLine("Sensor SDI-12 compatibility : " + singleSensor.get_sensorProtocol())
                Console.WriteLine("Sensor company name : " + singleSensor.get_sensorVendor())
                Console.WriteLine("Sensor model number : " + singleSensor.get_sensorModel())
                Console.WriteLine("Sensor version : " + singleSensor.get_sensorVersion())
                Console.WriteLine("Sensor serial number : " + singleSensor.get_sensorSerial())
                Dim valSensor As List(Of Double) = sdi12Port.readSensor(singleSensor.get_sensorAddress(), "M", 5000)
                Dim i = 0
                While (i < valSensor.Count)
                    If singleSensor.get_measureCount() > 1 Then
                        Console.WriteLine(String.Format("{0} : {1:0.00} {2} {3}", singleSensor.get_measureSymbol(i), valSensor(i),
                                                singleSensor.get_measureUnit(i), singleSensor.get_measureDescription(i)))
                    Else
                        Console.WriteLine(valSensor(i))
                    End If
                    i = i + 1
                End While
            Else
                Console.WriteLine("Module not connected (check identification and USB cable)")
            End If

            REM wait 5 sec to show the output
            System.Threading.Thread.Sleep(5000)
        End While
        YAPI.FreeAPI()
  End Sub

End Module