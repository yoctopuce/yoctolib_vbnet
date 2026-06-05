Module Module1
  Private KnownHubs As New HashSet(Of String)()

  Private Sub HubDiscovered(serial As String, url As String)
    ' The callback can be called multiple times for the same hub
    ' (discovery is based on periodic broadcast), so we use a set to avoid duplicates
    If KnownHubs.Contains(serial) Then Return

    Console.WriteLine("hub found: " & serial & " (" & url & ")")

    ' Add the hub to the set to avoid reprocessing
    KnownHubs.Add(serial)

    ' Connect to the hub
    Dim msg As String = ""
    Dim res As Integer = YAPI.RegisterHub(url, msg)
    If res <> YAPI.SUCCESS Then
      ' Ignore hubs with authentication
      Console.WriteLine("  Ignore hub " & serial & " (" & msg & ")")
      Return
    End If

    ' Find the hub module
    Dim hub As YModule = YModule.FindModule(serial)

    ' Iterate over all functions on the module and find the ports
    Dim fctCount As Integer = hub.functionCount()
    For i As Integer = 0 To fctCount - 1
      ' Retrieve the hardware name of the ith function
      Dim fctHwdName As String = hub.functionId(i)
      If fctHwdName.Length > 7 AndAlso fctHwdName.Substring(0, 7) = "hubPort" Then
        ' The port logical name is always the serial# of the connected device
        Dim deviceid As String = hub.functionName(i)
        Console.WriteLine("  " & fctHwdName & " : " & deviceid)
      End If
    Next

    ' Disconnect from the hub
    YAPI.UnregisterHub(url)
  End Sub

  Sub Main()
    Dim errmsg As String = ""

    REM create a dictionnary
    KnownHubs.Clear()

    Console.WriteLine("Waiting for hubs to signal themselves...")

    REM register the callback HubDiscovered will be
    REM invoked each time a hub signals its presence
    YAPI.RegisterHubDiscoveryCallback(AddressOf HubDiscovered)

    REM wait for 30 seconds, doing nothing.
    For i = 1 To 30
      YAPI.UpdateDeviceList(errmsg)
      YAPI.Sleep(1000, errmsg)
    Next
  End Sub
End Module
