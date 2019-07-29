﻿Module Module1
    Sub anButtonValueChangeCallBack(ByVal fct As YAnButton, ByVal value As String)
        Console.WriteLine(fct.get_hardwareId() + ": " + value + " (new value)")
    End Sub

    Sub sensorValueChangeCallBack(ByVal fct As YSensor, ByVal value As String)
        Console.WriteLine(
            fct.get_hardwareId() + ": " + value + " " + DirectCast(fct.get_userData(), String) + " (new value)")
    End Sub

    Sub sensorTimedReportCallBack(ByVal fct As YSensor, ByVal measure As YMeasure)
        Console.WriteLine(
            fct.get_hardwareId() + ": " + measure.get_averageValue().ToString() + " " +
            DirectCast(fct.get_userData(), String) + " (timed report)")
    End Sub

    Sub configChangeCallBack(ByVal modul As YModule)
        Console.WriteLine(modul.get_serialNumber() + ": configuration changed")
    End Sub

    Sub beaconCallBack(ByVal modul As YModule, ByVal beacon As Integer)
        Console.WriteLine(modul.get_serialNumber() + " beacon changed to " + beacon.ToString())
    End Sub

    Sub deviceArrival(ByVal m As YModule)
        Dim serial, hardwareId As String
        Dim fctcount, i As Integer
        Dim anButton As YAnButton
        Dim sensor As YSensor

        serial = m.get_serialNumber()
        Console.WriteLine("Device Arrival : " + serial)
        m.registerConfigChangeCallback(AddressOf configChangeCallBack)
        m.registerBeaconCallback(AddressOf beaconCallBack)

        REM // First solution: look for a specific type of function (eg. anButton)
        fctcount = m.functionCount()
        For i = 0 To fctcount - 1
            hardwareId = serial + "." + m.functionId(i)
            If hardwareId.IndexOf("anButton") >= 0 Then
                Console.WriteLine("- " + hardwareId)
                anButton = yFindAnButton(hardwareId)
                anButton.registerValueCallback(AddressOf anButtonValueChangeCallBack)
            End If
        Next

        REM // Alternate solution: register any kind of sensor on the device
        sensor = YSensor.FirstSensor()
        While sensor IsNot Nothing
            If sensor.get_module().get_serialNumber() = serial Then
                hardwareId = sensor.get_hardwareId()
                Console.WriteLine("- " + hardwareId)
                Dim unit As String = sensor.get_unit()
                sensor.set_userData(unit)
                sensor.registerValueCallback(AddressOf sensorValueChangeCallBack)
                sensor.registerTimedReportCallback(AddressOf sensorTimedReportCallBack)
            End If
            sensor = sensor.nextSensor()
        End While
    End Sub

    Sub deviceRemoval(ByVal m As YModule)
        Console.WriteLine("Device removal : " + m.get_serialNumber())
    End Sub

    Private Sub logFun(line As String)
        Console.Write("Log : " + line)
    End Sub

    Sub Main()
        Dim errmsg As String = ""

        REM Init API before first call
        If (YAPI.InitAPI(0, errmsg) <> YAPI_SUCCESS) Then
            Console.WriteLine("InitAPI error: " + errmsg)
            End
        End If

        YAPI.RegisterLogFunction(AddressOf logFun)
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
