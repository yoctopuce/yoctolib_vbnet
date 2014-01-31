
Imports System.IO
Imports System.Environment

Module Module1

  Function UnixTimestampToDateTime(ByVal _UnixTimeStamp As Long) As DateTime
    Return (New DateTime(1970, 1, 1, 0, 0, 0)).AddSeconds(_UnixTimeStamp)
  End Function

  Sub Main()

    Dim errmsg As String = ""
    Dim logger As YDataLogger
    Dim dataStreams = New List(Of YDataStream)
    Dim i As Integer

    REM No exception please
    YAPI.DisableExceptions()

    If (YAPI.RegisterHub("usb", errmsg) <> YAPI.SUCCESS) Then
      Console.WriteLine("yInitAPI failed: " + errmsg)
      End
    End If


    logger = YDataLogger.FirstDataLogger()
    If (logger Is Nothing) Then
      Console.WriteLine("No module with data logger found")
      Console.WriteLine("(Device not connected or firmware too old)")
      End
    End If

    Console.WriteLine("Using DataLogger of " + logger.get_Module().get_serialNumber())

    If (logger.get_dataStreams(dataStreams) <> YAPI.SUCCESS) Then
      Console.WriteLine("get_dataStreams failed: " + errmsg)
      End
    End If

    Console.WriteLine("found: " + Str(dataStreams.Count) + " stream(s) of data.")
    For i = 0 To dataStreams.Count - 1
      Dim s As YDataStream = dataStreams(i)
      Dim r, c, nrows, ncols As Long
      Console.WriteLine("Data stream  " + Str(i))
      Console.Write("- Run #" + Str(s.get_runIndex()))
      Console.Write(", time=" + Str(s.get_startTime()))
      If (s.get_startTimeUTC() > 0) Then
        Console.Write(", UTC =" + CStr(UnixTimestampToDateTime(s.get_startTimeUTC())))
      End If
      Console.WriteLine()

      nrows = CInt(s.get_rowCount())
      ncols = CInt(s.get_columnCount())

      If (nrows > 0) Then
        Console.Write(" " + Str(nrows) + " samples taken every ")
        Console.WriteLine(Str(s.get_dataSamplesInterval()) + " [s]")
      End If


      Dim names As List(Of String) = s.get_columnNames()
      Console.Write("   ")
      For c = 0 To names.Count - 1
        Console.Write(names(CInt(c)) + Chr(9))
      Next c
      Console.WriteLine()

      Dim table As List(Of List(Of Double)) = s.get_dataRows()

      For r = 0 To nrows - 1
        Console.Write("   ")
        For c = 0 To ncols - 1
          Console.Write(Format(table.ElementAt(CInt(r)).ElementAt(CInt(c)), "f") + Chr(9))
        Next c
        Console.WriteLine()
      Next r
    Next i
    Console.WriteLine("Done. Have a nice day :)")
  End Sub

End Module
