Module Module1


  Public Class YoctoShield
    Protected _serial As String
    Protected _subdevices As New List(Of String)

    Public Sub New(serial As String)
      _serial = serial
      _subdevices = New List(Of String)()
    End Sub

    Public Function getSerial() As String
      Return _serial
    End Function

    Public Function addSubdevice(serial As String) As Boolean
      For i = 1 To 4
        Dim p As YHubPort
        p = YHubPort.FindHubPort(_serial + ".hubPort" + i.ToString())
        If (p.get_logicalName() = serial) Then
          _subdevices.Add(serial)
          Return True
        End If
      Next
      Return False
    End Function

    Public Sub removeSubDevice(serial As String)
      _subdevices.Remove(serial)
    End Sub

    Public Sub describe()
      Console.WriteLine("  " + _serial)
      For Each subdev In _subdevices
        Console.WriteLine("    " + subdev)
      Next
    End Sub
  End Class


  Public Class RootDevice
    Protected _serial As String
    Protected _url As String
    Protected _shields As New List(Of YoctoShield)
    Protected _subdevices As New List(Of String)

    Public Sub New(serial As String, url As String)
      _serial = serial
      _url = url
      _shields = New List(Of YoctoShield)()
      _subdevices = New List(Of String)()
    End Sub

    Public Function getSerial() As String
      Return _serial
    End Function

    Public Sub addSubdevice(serial As String)
      If (serial.Substring(0, 7) = "YHUBSHL") Then
        _shields.Add(New YoctoShield(serial))
      Else
        REM Device to plug look if the device is plugged on a shield
        For Each shield In _shields
          If (shield.addSubdevice(serial)) Then
            Return
          End If
        Next
        _subdevices.Add(serial)
      End If
    End Sub

    Public Sub removeSubDevice(serial As String)
      _subdevices.Remove(serial)
    End Sub

    Public Sub describe()
      Console.WriteLine("  " + _serial + " (" + _url + ")")
      For Each subdev In _subdevices
        Console.WriteLine("  " + subdev)
      Next
      For Each shield In _shields
          shield.describe()
        Next
    End Sub
  End Class


  Dim __rootDevices As New List(Of RootDevice)()

  Function getYoctoHub(serial As String) As RootDevice
    Dim rootdev As RootDevice
    rootdev = __rootDevices.ElementAt(1)
    For Each rootdev In __rootDevices
      If rootdev.getSerial() = serial Then
        Return rootdev
      End If
    Next
    Return Nothing
  End Function

  Function addRootDevice(serial As String, url As String) As RootDevice
    For Each root As RootDevice In __rootDevices
      If (root.getSerial() = serial) Then
        Return root
      End If
    Next
    Dim RootDevice As RootDevice = New RootDevice(serial, url)
    __rootDevices.Add(RootDevice)
    Return RootDevice

  End Function

  Sub showNetwork()
    Console.WriteLine("**** device inventory *****")
    For Each root In __rootDevices
      root.describe()
    Next
  End Sub




  Sub deviceArrival(ByVal m As YModule)
    Dim serial, parentHub As String
    Dim url As String
    Dim hub As RootDevice

    serial = m.get_serialNumber()
    parentHub = m.get_parentHub()
    If parentHub = "" Then
      REM root device
      url = m.get_url()
      addRootDevice(serial, url)
    Else
      hub = getYoctoHub(parentHub)
      If hub IsNot Nothing Then
        hub.addSubdevice(serial)
      End If
    End If
  End Sub

  Sub deviceRemoval(ByVal m As YModule)
    Dim serial As String = m.get_serialNumber()
    For i = __rootDevices.Count() - 1 To 0 Step -1
      __rootDevices.ElementAt(i).removeSubDevice(serial)
      If __rootDevices.ElementAt(i).getSerial() = serial Then
        __rootDevices.RemoveAt(i)
      End If
    Next
  End Sub

  Sub Main()
    Dim errmsg As String = ""

    REM Init API before first call

    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error : " + errmsg)
      End
    End If

    If (YAPI.RegisterHub("net", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error : " + errmsg)
      End
    End If


    REM each time a new device is connected/discovered
    REM arrivalCallback will be called.
    YAPI.RegisterDeviceArrivalCallback(AddressOf deviceArrival)
    REM each time a device is disconnected/removed
    REM removalCallback will be called.
    YAPI.RegisterDeviceRemovalCallback(AddressOf deviceRemoval)


    Console.WriteLine("Waiting for hubs to signal themselves...")

    While (True)
      YAPI.UpdateDeviceList(errmsg) REM traps plug/unplug events
      YAPI.Sleep(500, errmsg) REM   rem traps others events
      showNetwork()
    End While
  End Sub

End Module
