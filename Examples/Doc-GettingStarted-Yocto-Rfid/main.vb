' ********************************************************************
'
'  $Id: main.vb 60135 2024-03-25 08:14:08Z seb $
'
'  An example that shows how to use a  Yocto-RFID-15693
'
'  You can find more information on our web site:
'   Yocto-RFID-15693 documentation:
'      https://www.yoctopuce.com/EN/products/yocto-rfid-15693/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1
    Private Sub Usage()
        Dim execname = System.AppDomain.CurrentDomain.FriendlyName
        Console.WriteLine("Usage:")
        Console.WriteLine(execname + "  <serial_number>  frequency")
        Console.WriteLine(execname + "  <logical_name> frequency")
        Console.WriteLine(execname + "  any frequency")
        System.Threading.Thread.Sleep(2500)
        End
    End Sub

    Sub Main()
        Dim argv() As String = System.Environment.GetCommandLineArgs()
        Dim errmsg As String = ""
        Dim buz As YBuzzer
        Dim leds As YColorLedCluster
        Dim button1 As YAnButton
        Dim reader As YRfidReader
        Dim target, serial As String
        Dim tagList As List(Of String)

        If argv.Length < 2 Then Usage()
        target = argv(1)

        REM Setup the API to use local USB devices
        If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
            Console.WriteLine("RegisterHub error: " + errmsg)
            End
        End If

        If target = "any" Then
            reader = YRfidReader.FirstRfidReader()
            If reader Is Nothing Then
                Console.WriteLine("No module connected (check USB cable) ")
                End
            End If
        Else
            reader = YRfidReader.FindRfidReader(target + ".rfidReader")
        End If

        If (Not reader.isOnline()) Then
            Console.WriteLine("No module connected (check identification and USB cable)")
            End
        End If

        serial = reader.get_module().get_serialNumber()
        leds = YColorLedCluster.FindColorLedCluster(serial + ".colorLedCluster")
        button1 = YAnButton.FindAnButton(serial + ".anButton1")
        buz = YBuzzer.FindBuzzer(serial + ".buzzer")

        tagList = New List(Of String)(32)
        leds.set_rgbColor(0, 1, 0)
        buz.set_volume(75)

        Console.WriteLine("Place a RFID tag near the Antenna")
        While (tagList.Count <= 0)
            tagList = reader.get_tagIdList()
            YAPI.Sleep(250, errmsg)
        End While

        Dim tagId As String = tagList(0)
        Dim opStatus As YRfidStatus = New YRfidStatus()
        Dim options As YRfidOptions = New YRfidOptions()
        Dim taginfo As YRfidTagInfo = reader.get_tagInfo(tagId, opStatus)
        Dim blocksize As Integer = taginfo.get_tagBlockSize()
        Dim firstBlock As Integer = taginfo.get_tagFirstBlock()
        Console.WriteLine("Tag ID          = " + taginfo.get_tagId())
        Console.WriteLine("Tag Memory size = " + taginfo.get_tagMemorySize().ToString() + " bytes")
        Console.WriteLine("Tag Block  size = " + taginfo.get_tagBlockSize().ToString() + " bytes")

        Dim Data As String = reader.tagReadHex(tagId, firstBlock, 3 * blocksize, options, opStatus)
        If (opStatus.get_errorCode() = YRfidStatus.SUCCESS) Then
            Console.WriteLine("First 3 blocks  = " + Data)
            leds.set_rgbColor(0, 1, &HFF00)
            buz.pulse(1000, 100)
        Else
            Console.WriteLine("Cannot read tag contents (" + opStatus.get_errorMessage() + ")")
            leds.set_rgbColor(0, 1, &HFF0000)
        End If
        leds.rgb_move(0, 1, &H0, 200)
        YAPI.FreeAPI()
    End Sub
End Module
