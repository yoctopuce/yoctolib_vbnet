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



  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim files As YFiles



    If argv.Length <= 1 Then Usage()

    target = argv(1)

    REM Setup the API to use local USB devices
    If (yRegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      files = yFirstFiles()
      If files Is Nothing Then
        Console.WriteLine("No module with files feature connected (check USB cable) ")
        End
      End If

    Else
      files = yFindFiles(target + ".files")

    End If

    If Not (files.isOnline()) Then
      Console.WriteLine("No module with files feature connected (check identification and USB cable)")
      End
    End If

    Console.WriteLine()
    Console.WriteLine("Using " + files.get_friendlyName())
    Console.WriteLine()

    Dim binaryData As Byte()
    Dim encoding As New System.Text.ASCIIEncoding()

    REM create text files and upload them to the device
    For i = 1 To 5

      Dim contents As String = "This is file " + CStr(i)

      REM convert the string to binary data
      binaryData = encoding.GetBytes(contents)
      REM upload the file to the device
      files.upload("file" + CStr(i) + ".txt", binaryData)

    Next i

    REM list files found on the device
    Console.WriteLine("Files on device:")
    Dim filelist As New List(Of YFileRecord)
    filelist = files.get_list("*")
    For i = 0 To filelist.Count() - 1

      Dim filename As String = filelist(i).get_name()
      Console.Write(filename)
      Console.Write(New String(Chr(32), 40 - filename.Length)) REM align
      Console.Write(filelist(i).get_crc().ToString("X"))
      Console.Write("    ")
      Console.WriteLine(CStr(filelist(i).get_size()) + " bytes")
    Next i

    REM download a file
    binaryData = files.download("file1.txt")

    REM convert to string
    Dim st As String = encoding.GetString(binaryData)

    REM and display
    Console.WriteLine("")
    Console.WriteLine("contents of file1.txt:")
    Console.WriteLine(st)

  End Sub

End Module
