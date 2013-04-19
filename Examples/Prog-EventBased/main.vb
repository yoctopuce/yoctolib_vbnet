Module Module1

  Sub anButtonChangeCallBack(ByVal fct As YAnButton, ByVal value As String)
    Console.WriteLine("Position change         :" + fct.ToString() + " = " + value)
  End Sub

  Sub temperatureChangeCallBack(ByVal fct As YTemperature, ByVal value As String)
    Console.WriteLine("Temperature change      :" + fct.ToString() + " = " + value + "°C")
  End Sub

  Sub lightSensorChangeCallBack(ByVal fct As YLightSensor, ByVal value As String)
    Console.WriteLine("Light change            :" + fct.ToString() + " = " + value + "lx")
  End Sub

  Sub deviceArrival(ByVal m As YModule)
    Dim fctName, fctFullName As String
    Dim i As Integer
    Dim fctcount = m.functionCount()
    Dim bt As YAnButton
    Dim t As YTemperature
    Dim l As YLightSensor
    Console.WriteLine("Device arrival          : " + m.ToString())

    For i = 0 To fctcount - 1
      fctName = m.functionId(i)
      fctFullName = m.get_serialNumber() + "." + fctName

      REM register call back for anbuttons
      If fctName.IndexOf("anButton") = 0 Then
        bt = yFindAnButton(fctFullName)
        If bt.isOnline() Then bt.registerValueCallback(AddressOf anButtonChangeCallBack)
        Console.WriteLine("Callback registered for : " + fctFullName)
      End If

      REM register call back for anbuttons
      If fctName.IndexOf("temperature") = 0 Then
        t = yFindTemperature(fctFullName)
        If t.isOnline() Then t.registerValueCallback(AddressOf temperatureChangeCallBack)
        Console.WriteLine("Callback registered for : " + fctFullName)
      End If

      REM register call back for light sensors
      If fctName.IndexOf("lightSensor") = 0 Then
        l = yFindLightSensor(fctFullName)
        If l.isOnline() Then l.registerValueCallback(AddressOf lightSensorChangeCallBack)
        Console.WriteLine("Callback registered for : " + fctFullName)
      End If

      REM and so on for other sensor type.....
    Next

  End Sub

  Sub deviceRemoval(ByVal m As YModule)
    Console.WriteLine("Device removal          : " + m.get_serialNumber())
  End Sub


  Sub Main()
    Dim errmsg As String = ""

    REM Init API before first call
    If (YAPI.InitAPI(0, errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("InitAPI error: " + errmsg)
      End
    End If

    YAPI.RegisterDeviceArrivalCallback(AddressOf deviceArrival)
    YAPI.RegisterDeviceRemovalCallback(AddressOf deviceRemoval)

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
