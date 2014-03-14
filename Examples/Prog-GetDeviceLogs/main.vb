Module Module1

  Sub logfun(ByVal m As YModule, ByVal logline As String)
    Console.WriteLine(m.get_serialNumber() + ": " + logline)
  End Sub

  Sub deviceArrival(ByVal m As YModule)
    Dim serial As String

    serial = m.get_serialNumber()
    Console.WriteLine("Device Arrival : " + serial)
    m.registerLogCallback(AddressOf logfun)

  End Sub

 
  Sub Main()
    Dim errmsg As String = ""

    REM Init API before first call
    If (YAPI.InitAPI(0, errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("InitAPI error: " + errmsg)
      End
    End If

    YAPI.RegisterDeviceArrivalCallback(AddressOf deviceArrival)

    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error : " + errmsg)
      End
    End If

    Console.WriteLine("Hit Ctrl-C to Stop ")

    While (True)
      YAPI.UpdateDeviceList(errmsg) REM traps plug/unplug events
      YAPI.Sleep(500, errmsg) REM   rem traps others events
    End While
  End Sub

End Module
