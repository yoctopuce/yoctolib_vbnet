
Imports System.IO
Imports System.Environment

Module Module1

  Sub Main()
    Dim _epoch As DateTime = New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
    Dim errmsg As String = ""
    Dim sensor As YSensor

    If (YAPI.RegisterHub("usb", errmsg) <> YAPI.SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      Environment.Exit(1)
    End If

    REM Enumerate all connected sensors
    Dim sensorList As List(Of YSensor) = New List(Of YSensor)
    sensor = YSensor.FirstSensor()
    While (sensor IsNot Nothing)
      sensorList.Add(sensor)
      sensor = sensor.nextSensor()
    End While
    If sensorList.Count = 0 Then
      Console.WriteLine("No module connected (check USB cable)")
      Environment.Exit(1)
    End If

    REM Generate consolidated CSV output For all sensors
    Dim data As YConsolidatedDataSet = New YConsolidatedDataSet(0, 0, sensorList)
    Dim record As List(Of Double) = New List(Of Double)
    While data.nextRecord(record) < 100
      Dim line As String = _epoch.AddSeconds(record(0)).ToString("yyyy-MM-ddTHH:mm:ss.fff")
      For i = 1 To record.Count - 1
        line += String.Format(";{0:0.000}", record(i))
      Next
      Console.WriteLine(line)
    End While

    YAPI.FreeAPI()
    Console.WriteLine("Done. Have a nice day :)")
  End Sub

End Module
