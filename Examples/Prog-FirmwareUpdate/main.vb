
Imports System.IO
Imports System.Environment

Module Module1

  Sub upgradeSerialList(ByRef allserials As List(Of String))
    Dim errmsg As String = ""
    For Each serial As String In allserials
      Dim m As YModule = YModule.FindModule(serial)
      Dim product As String = m.get_productName()
      Dim current As String = m.get_firmwareRelease()

      REM check if a new firmare is available on yoctopuce.com
      Dim newfirm As String = m.checkFirmware("www.yoctopuce.com", True)
      If newfirm = "" Then
        Console.WriteLine(product + " " + serial + "(rev=" + current + ") is up to date")
      Else
        Console.WriteLine(product + " " + serial + "(rev=" + current + ") need be updated with firmare : ")
        Console.WriteLine("    " + newfirm)
        REM execute the firmware upgrade
        Dim update As YFirmwareUpdate = m.updateFirmware(newfirm)
        Dim status As Integer = update.startUpdate()
        While (status < 100 And status >= 0)
          Dim newstatus As Integer = update.get_progress()
          If (newstatus <> status) Then
            Console.WriteLine(status.ToString() + "% " + update.get_progressMessage())
          End If

          YAPI.Sleep(500, errmsg)
          status = newstatus
        End While
        If (status < 0) Then
          Console.WriteLine("    " + status.ToString() + " Firmware Update failed: " + update.get_progressMessage())
          Environment.Exit(1)
        Else
          If (m.isOnline()) Then
            Console.WriteLine(status.ToString() + "% Firmware Updated Successfully!")
          Else
            Console.WriteLine(status.ToString() + " Firmware Update failed: module " + serial + "is not online")
            Environment.Exit(1)
          End If
        End If
      End If
    Next
  End Sub


  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg = ""
    Dim i As Integer
    Dim hubs As List(Of String) = New List(Of String)()
    Dim shield As List(Of String) = New List(Of String)()
    Dim devices As List(Of String) = New List(Of String)()


    If (YAPI.RegisterHub("usb", errmsg) <> YAPI.SUCCESS) Then
      Console.WriteLine("yInitAPI failed: " + errmsg)
      End
    End If

    i = 1
    While i < argv.Count
      If (YAPI.RegisterHub(argv(i), errmsg) <> YAPI.SUCCESS) Then
        Console.WriteLine("yInitAPI failed: " + errmsg)
        End
      End If
    End While

    REM fist step construct the list of all hub /shield and devices connected
    Dim m As YModule = YModule.FirstModule()
    While m IsNot Nothing
      Dim product As String = m.get_productName()
      Dim serial As String = m.get_serialNumber()
      If (product = "YoctoHub-Shield") Then
        shield.Add(serial)
      ElseIf (product.StartsWith("YoctoHub-")) Then
        hubs.Add(serial)
      ElseIf (product <> "VirtualHub") Then
        devices.Add(serial)
      End If
      m = m.nextModule()
    End While
    REM fist upgrades all Hubs...
    upgradeSerialList(hubs)
    REM ... then all shield..
    upgradeSerialList(shield)
    REM ... and finaly all devices
    upgradeSerialList(devices)
    Console.WriteLine("All devices are now up to date")
    YAPI.FreeAPI()


  End Sub


End Module
