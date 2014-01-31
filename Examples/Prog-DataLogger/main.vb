
Imports System.IO
Imports System.Environment

Module Module1

  Sub dumpSensor(sensor As YSensor)
    Dim fmt As String = "dd MMM yyyy hh:mm:ss,fff"
    Console.WriteLine("Using DataLogger of " + sensor.get_friendlyName())
    Dim dataset As YDataSet = sensor.get_recordedData(0, 0)
    Console.WriteLine("loading summary... ")
    dataset.loadMore()
    Dim summary As YMeasure = dataset.get_summary()
    Dim line As String = String.Format("from {0} to {1} : min={2:0.00}{5} avg={3:0.00}{5}  max={4:0.00}{5}", summary.get_startTimeUTC_asDateTime().ToString(fmt), summary.get_endTimeUTC_asDateTime().ToString(fmt), summary.get_minValue(), summary.get_averageValue(), summary.get_maxValue(), sensor.get_unit())
    Console.WriteLine(line)
    Console.WriteLine("loading details :   0%")
    Dim progress As Integer = 0
    While (progress < 100)
      progress = dataset.loadMore()
      Console.WriteLine(String.Format("{0,3:##0}%", progress))
    End While
    Dim details As List(Of YMeasure) = dataset.get_measures()
    For Each m In details
      Console.WriteLine(String.Format("from {0} to {1} : min={2:0.00}{5} avg={3:0.00}{5}  max={4:0.00}{5}",
              m.get_startTimeUTC_asDateTime().ToString(fmt), m.get_endTimeUTC_asDateTime().ToString(fmt), m.get_minValue(), m.get_averageValue(), m.get_maxValue(), sensor.get_unit()))
    Next
  End Sub



  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim sensor As YSensor

    If (YAPI.RegisterHub("usb", errmsg) <> YAPI.SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      Environment.Exit(1)
    End If
    If ((argv.Length = 1) OrElse (argv(1) = "any")) Then
      sensor = YSensor.FirstSensor()
      If (sensor Is Nothing) Then
        Console.WriteLine("No module connected (check USB cable)")
        Environment.Exit(1)
      End If
    Else
      sensor = YSensor.FindSensor(argv(1))
      If (Not sensor.isOnline()) Then
        Console.WriteLine("Sensor " + sensor.get_hardwareId + " is not connected (check USB cable)")
        Environment.Exit(1)
      End If
    End If
    dumpSensor(sensor)
    YAPI.FreeAPI()
    Console.WriteLine("Done. Have a nice day :)")
  End Sub

End Module
