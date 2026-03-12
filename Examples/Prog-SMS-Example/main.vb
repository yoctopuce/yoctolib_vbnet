Module Module1

    Private Sub Usage()
        Dim execname = System.AppDomain.CurrentDomain.FriendlyName
        Console.WriteLine("Usage:")
        Console.WriteLine(execname + " <serial_number>")
        Console.WriteLine(execname + " <logical_name> ")
        Console.WriteLine(execname + "  any")
        System.Threading.Thread.Sleep(2500)
        End
    End Sub


    Private Sub smsCallback(ByVal mbox As YMessageBox, ByVal sms As YSms)
        Console.WriteLine("- New message dated " + sms.get_timestamp())
        Console.WriteLine("  from " + sms.get_sender())
        Console.WriteLine("  '" + sms.get_textData() + "'")
        sms.deleteFromSIM()
    End Sub


    Sub Main()
        Dim argv() As String = System.Environment.GetCommandLineArgs()
        Dim errmsg As String = ""
        Dim target As String
        Dim mbox As YMessageBox

        If argv.Length <= 1 Then Usage()
        target = argv(1)

        REM Setup the API to use local USB devices
        If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
            Console.WriteLine("RegisterHub error: " + errmsg)
            System.Threading.Thread.Sleep(2500)
            End
        End If

        If target = "any" Then
            mbox = YMessageBox.FirstMessageBox()
            If mbox Is Nothing Then
                Console.WriteLine("No module with SMS feature connected (check USB cable) ")
                System.Threading.Thread.Sleep(2500)
                End
            End If
        Else
            mbox = YMessageBox.FindMessageBox(target + ".messageBox")
        End If

        If Not (mbox.isOnline()) Then
            Console.WriteLine("Module not connected (check identification and USB cable)")
            System.Threading.Thread.Sleep(2500)
            End
        End If

        Console.WriteLine()
        Console.WriteLine("Using " + mbox.get_friendlyName())
        Console.WriteLine()

        REM list messages found on the device
        Console.WriteLine("Messages found on the SIM card:")
        Dim messages As New List(Of YSms)
        messages = mbox.get_messages()
        For i = 0 To messages.Count() - 1
            Dim sms As YSms = messages(i)
            Console.WriteLine("- dated " + sms.get_timestamp())
            Console.WriteLine("  from " + sms.get_sender())
            Console.WriteLine("  '" + sms.get_textData() + "'")
        Next i

        REM register a callback to receive any new message
        mbox.registerSmsCallback(AddressOf smsCallback)

        REM offer to send a new message
        Console.WriteLine("To test sending SMS, provide message recipient (+xxxxxxx).")
        Console.WriteLine("To skip sending, leave empty and press Enter.")
        Dim number As String = Console.ReadLine()
        If number <> "" Then
            REM if that call fails, make sure that your SIM operator
            REM allows you to send SMS given your current contract
            mbox.sendTextMessage(number, "Hello from YoctoHub-GSM !")
        End If

        Console.WriteLine("Waiting to receive SMS, press Ctrl-C to quit")
        While True
            YAPI.Sleep(2000, errmsg)
        End While
    End Sub

End Module
