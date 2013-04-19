'/********************************************************************
'*
'* $Id: yocto_datalogger.vb 10552 2013-03-20 16:13:16Z mvuilleu $
'*
'* High-level programming interface, common to all modules
'*
'* - - - - - - - - - License information: - - - - - - - - - 
'*
'* Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'*
'* 1) If you have obtained this file from www.yoctopuce.com,
'*    Yoctopuce Sarl licenses to you (hereafter Licensee) the
'*    right to use, modify, copy, and integrate this source file
'*    into your own solution for the sole purpose of interfacing
'*    a Yoctopuce product with Licensee's solution.
'*
'*    The use of this file and all relationship between Yoctopuce 
'*    and Licensee are governed by Yoctopuce General Terms and 
'*    Conditions.
'*
'*    THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'*    WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING 
'*    WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS 
'*    FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'*    EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'*    INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA, 
'*    COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR 
'*    SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT 
'*    LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'*    CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'*    BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'*    WARRANTY, OR OTHERWISE.
'*
'* 2) If your intent is not to interface with Yoctopuce products,
'*    you are not entitled to use, read or create any derived
'*    material from this source file.
'*
'/********************************************************************/

Imports ys8 = System.SByte
Imports ys16 = System.Int16
Imports ys32 = System.Int32
Imports ys64 = System.Int64
Imports yu8 = System.Byte
Imports yu16 = System.UInt16
Imports yu32 = System.UInt32
Imports yu64 = System.UInt64
Imports YDEV_DESCR = System.Int32      REM yStrRef of serial number
Imports YFUN_DESCR = System.Int32      REM yStrRef of serial + (ystrRef of funcId << 16)
Imports System.Runtime.InteropServices

Module yocto_datalogger

  REM --- (generated code: YDataLogger definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YDataLogger, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_OLDESTRUNINDEX_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_CURRENTRUNINDEX_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_SAMPLINGINTERVAL_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_TIMEUTC_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_RECORDING_OFF = 0
  Public Const Y_RECORDING_ON = 1
  Public Const Y_RECORDING_INVALID = -1

  REM Y_AUTOSTART is defined in yocto_api.vb
  Public Const Y_CLEARHISTORY_FALSE = 0
  Public Const Y_CLEARHISTORY_TRUE = 1
  Public Const Y_CLEARHISTORY_INVALID = -1



  REM --- (end of generated code: YDataLogger definitions)

  Public Const Y_DATA_INVALID = YAPI.INVALID_DOUBLE


  Public Class YDataStream

    Protected dataLogger As YDataLogger
    Protected runNo As Integer
    Protected timeStamp As Long
    Protected interval As Long
    Protected utcStamp As Long
    Protected nRows As Integer
    Protected nCols As Integer
    Protected columnNames As List(Of String)
    Protected values(,) As Double

    Public Sub New(ByVal parent As YDataLogger, ByVal run As Integer, ByVal stamp As Integer, ByVal utc As Long, ByVal itv As Integer)
      dataLogger = parent
      runNo = run
      timeStamp = stamp
      utcStamp = utc
      interval = itv
      nRows = 0
      nCols = 0
      columnNames = New List(Of String)
      values = Nothing
    End Sub

    Protected Overridable Overloads Sub Dispose()
      columnNames = Nothing
      values = Nothing
    End Sub

    '''*
    ''' <summary>
    '''   Returns the run index of the data stream.
    ''' <para>
    '''   A run can be made of
    '''   multiple datastreams, for different time intervals.
    ''' </para>
    ''' <para>
    '''   This method does not cause any access to the device, as the value
    '''   is preloaded in the object at instantiation time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the run index.
    ''' </returns>
    '''/
    Public Function get_runIndex() As Integer
      get_runIndex = runNo
    End Function

    '''*
    ''' <summary>
    '''   Returns the start time of the data stream, relative to the beginning
    '''   of the run.
    ''' <para>
    '''   If you need an absolute time, use <c>get_startTimeUTC()</c>.
    ''' </para>
    ''' <para>
    '''   This method does not cause any access to the device, as the value
    '''   is preloaded in the object at instantiation time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of seconds
    '''   between the start of the run and the beginning of this data
    '''   stream.
    ''' </returns>
    '''/
    Public Function get_startTime() As Integer
      get_startTime = CInt(timeStamp)
    End Function

    '''*
    ''' <summary>
    '''   Returns the start time of the data stream, relative to the Jan 1, 1970.
    ''' <para>
    '''   If the UTC time was not set in the datalogger at the time of the recording
    '''   of this data stream, this method returns 0.
    ''' </para>
    ''' <para>
    '''   This method does not cause any access to the device, as the value
    '''   is preloaded in the object at instantiation time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of seconds
    '''   between the Jan 1, 1970 and the beginning of this data
    '''   stream (i.e. Unix time representation of the absolute time).
    ''' </returns>
    '''/
    Public Function get_startTimeUTC() As Long
      get_startTimeUTC = utcStamp
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of seconds elapsed between  two consecutive
    '''   rows of this data stream.
    ''' <para>
    '''   By default, the data logger records one row
    '''   per second, but there might be alternative streams at lower resolution
    '''   created by summarizing the original stream for archiving purposes.
    ''' </para>
    ''' <para>
    '''   This method does not cause any access to the device, as the value
    '''   is preloaded in the object at instantiation time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to a number of seconds.
    ''' </returns>
    '''/
    Public Function get_dataSamplesInterval() As Integer
      get_dataSamplesInterval = CInt(interval)
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of data rows present in this stream.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   This method fetches the whole data stream from the device,
    '''   if not yet done.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of rows.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns zero.
    ''' </para>
    '''/
    Public Function get_rowCount() As Integer
      If (nRows = 0) Then loadStream()
      get_rowCount = nRows
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of data columns present in this stream.
    ''' <para>
    '''   The meaning of the values present in each column can be obtained
    '''   using the method <c>get_columnNames()</c>.
    ''' </para>
    ''' <para>
    '''   This method fetches the whole data stream from the device,
    '''   if not yet done.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of rows.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns zero.
    ''' </para>
    '''/
    Public Function get_columnCount() As Integer
      If (nCols = 0) Then loadStream()
      get_columnCount = nCols
    End Function

    '''*
    ''' <summary>
    '''   Returns the title (or meaning) of each data column present in this stream.
    ''' <para>
    '''   In most case, the title of the data column is the hardware identifier
    '''   of the sensor that produced the data. For archived streams created by
    '''   summarizing a high-resolution data stream, there can be a suffix appended
    '''   to the sensor identifier, such as _min for the minimum value, _avg for the
    '''   average value and _max for the maximal value.
    ''' </para>
    ''' <para>
    '''   This method fetches the whole data stream from the device,
    '''   if not yet done.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list containing as many strings as there are columns in the
    '''   data stream.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Function get_columnNames() As List(Of String)
      If (columnNames.Count = 0) Then loadStream()
      get_columnNames = columnNames
    End Function

    '''*
    ''' <summary>
    '''   Returns the whole data set contained in the stream, as a bidimensional
    '''   table of numbers.
    ''' <para>
    '''   The meaning of the values present in each column can be obtained
    '''   using the method <c>get_columnNames()</c>.
    ''' </para>
    ''' <para>
    '''   This method fetches the whole data stream from the device,
    '''   if not yet done.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list containing as many elements as there are rows in the
    '''   data stream. Each row itself is a list of floating-point
    '''   numbers.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Function get_dataRows() As Array
      If (values Is Nothing) Then loadStream()
      get_dataRows = values
    End Function

    '''*
    ''' <summary>
    '''   Returns a single measure from the data stream, specified by its
    '''   row and column index.
    ''' <para>
    '''   The meaning of the values present in each column can be obtained
    '''   using the method get_columnNames().
    ''' </para>
    ''' <para>
    '''   This method fetches the whole data stream from the device,
    '''   if not yet done.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="row">
    '''   row index
    ''' </param>
    ''' <param name="col">
    '''   column index
    ''' </param>
    ''' <returns>
    '''   a floating-point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_DATA_INVALID.
    ''' </para>
    '''/
    Public Function get_data(ByVal row As Integer, ByVal col As Integer) As Double
      If (values Is Nothing) Then loadStream()

      If (row >= nRows Or row < 0) Then
        get_data = Y_DATA_INVALID
        Exit Function
      End If

      If (col >= nCols Or col < 0) Then
        get_data = Y_DATA_INVALID
        Exit Function
      End If

      get_data = values(row, col)
    End Function

        Private Function loadStream() As Integer


            Dim json As TJsonParser = Nothing
            Dim res, count As Integer
            Dim root, el As TJSONRECORD
            Dim name As String
            Dim coldiv As List(Of Integer) = New List(Of Integer)
            Dim coltype As List(Of Integer) = New List(Of Integer)
            Dim udat As List(Of UInteger) = New List(Of UInteger)
            Dim dat As List(Of Double) = New List(Of Double)
            Dim colscl As List(Of Double) = New List(Of Double)
            Dim colofs As List(Of Integer) = New List(Of Integer)
            Dim caltyp As List(Of Integer) = New List(Of Integer)
            Dim calhdl As List(Of yCalibrationHandler) = New List(Of yCalibrationHandler)
            Dim calpar As List(Of Integer()) = New List(Of Integer())
            Dim calraw As List(Of Double()) = New List(Of Double())
            Dim calref As List(Of Double()) = New List(Of Double())


            Dim x, y, i, j As Integer
            Dim value As Double
            Dim sdat As String

            res = dataLogger.getData(runNo, timeStamp, json)
            If (res <> YAPI_SUCCESS) Then
                loadStream = res
                Exit Function
            End If

            nRows = 0
            nCols = 0
            columnNames.Clear()
            ReDim values(0, 0)


            root = json.GetRootNode()
            For i = 0 To root.membercount - 1

                el = root.members(i)
                name = el.name
                If (name = "time") Then

                    timeStamp = el.ivalue
                ElseIf (name = "UTC") Then
                    utcStamp = el.ivalue
                ElseIf (name = "interval") Then
                    interval = el.ivalue
                ElseIf (name = "nRows") Then
                    nRows = el.ivalue
                ElseIf (name = "keys") Then
                    nCols = el.itemcount
                    For j = 0 To nCols - 1
                        columnNames.Add(el.items(j).svalue)
                    Next j
                ElseIf (name = "div") Then
                    nCols = el.itemcount
                    For j = 0 To nCols - 1
                        coldiv.Add(el.items(j).ivalue)
                    Next j
                ElseIf (name = "type") Then
                    nCols = el.itemcount
                    For j = 0 To nCols - 1
                        coltype.Add(el.items(j).ivalue)
                    Next j
                ElseIf (name = "scal") Then
                    nCols = el.itemcount
                    For j = 0 To nCols - 1
                        colscl.Add(el.items(j).ivalue / 65536.0)
                        If coltype(j) <> 0 Then
                            colofs.Add(-32767)
                        Else
                            colofs.Add(0)
                        End If
                    Next j
                ElseIf (name = "cal") Then
                    nCols = el.itemcount
                    For j = 0 To nCols - 1
                        Dim calibration_Str As String = el.items(j).svalue
                        Dim cur_calpar() As Integer = Nothing
                        Dim cur_calraw() As Double = Nothing
                        Dim cur_calref() As Double = Nothing
                        Dim calibType As Integer = YAPI._decodeCalibrationPoints(calibration_Str, cur_calpar, cur_calraw, cur_calref, colscl(j), colofs(j))
                        caltyp.Add(calibType)
                        calhdl.Add(YAPI._getCalibrationHandler(calibType))
                        calpar.Add(cur_calpar)
                        calraw.Add(cur_calraw)
                        calref.Add(cur_calref)
                    Next j
                ElseIf (name = "data") Then
                    If (colscl.Count <= 0) Then
                        For j = 0 To nCols - 1
                            colscl.Add(1.0 / coldiv(j))
                            If (coltype(j) <> 0) Then
                                colofs.Add(-32767)
                            Else
                                colofs.Add(0)
                            End If
                        Next j
                    End If
                    count = el.itemcount
                    udat.Clear()
                    If (el.recordtype = TJSONRECORDTYPE.JSON_STRING) Then
                        sdat = el.svalue
                        Dim p As Integer = 0
                        While (p < sdat.Length)
                            Dim val As UInteger
                            Dim c As UInteger = CUInt(Asc(sdat.Substring(p, 1)))
                            p += 1
                            If (c >= 97) Then REM 97 ='a'
                                Dim srcpos As Integer = CInt((udat.Count - 1 - (c - 97)))
                                If (srcpos < 0) Then
                                    dataLogger.throw_friend(YAPI_IO_ERROR, "Unexpected JSON reply format")
                                    Return YAPI_IO_ERROR
                                End If
                                val = udat.ElementAt(srcpos)
                            Else
                                If (p + 2 > sdat.Length) Then
                                    dataLogger.throw_friend(YAPI_IO_ERROR, "Unexpected JSON reply format")
                                    Return YAPI_IO_ERROR
                                End If
                                val = CUInt(c - 48) REM 48='0'
                                c = CUInt(Asc(sdat.Substring(p, 1)))
                                p += 1
                                val += CUInt(c - 48) << 5
                                c = CUInt(Asc(sdat.Substring(p, 1)))
                                p += 1
                                If (c = 122) Then REM 122 ='z'
                                    c = 92 REM 92 ='\'
                                End If
                                val += CUInt(c - 48) << 10
                            End If
                            udat.Add(val)
                        End While
                    Else
                        count = el.itemcount
                        For j = 0 To count - 1
                            Dim tmp As UInteger = CUInt(el.items(j).ivalue)
                            udat.Add(tmp)
                        Next
                    End If


                    ReDim values(nRows, nCols)
                    x = 0
                    y = 0
                    For Each uval As Integer In udat
                        If coltype(x) < 2 Then
                            value = (uval + colofs(x)) * colscl(x)
                        Else
                            value = YAPI._decimalToDouble(uval - 32767)
                        End If
                        If (caltyp(x) > 0 And calhdl(x) <> Nothing) Then
                            Dim handler As yCalibrationHandler = calhdl(x)
                            If (caltyp(x) <= 10) Then
                                value = handler((uval + coldiv(x)) / coldiv(x), caltyp(x), calpar(x), calraw(x), calref(x))
                            ElseIf (caltyp(x) > 20) Then
                                value = handler(value, caltyp(x), calpar(x), calraw(x), calref(x))
                            End If
                        End If
                        values(y, x) = value
                        x = x + 1
                        If (x = nCols) Then
                            x = 0
                            y = y + 1
                        End If
                    Next
                End If
            Next i

            json = Nothing

            loadStream = YAPI_SUCCESS
        End Function


  End Class

  REM --- (end of generated code: DataLogger functions declaration)

  REM --- (generated code: YDataLogger implementation)

  Private _DataLoggerCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   Yoctopuce sensors include a non-volatile memory capable of storing ongoing measured
  '''   data automatically, without requiring a permanent connection to a computer.
  ''' <para>
  '''   The Yoctopuce application programming interface includes functions to control
  '''   how this internal data logger works.
  '''   Beacause the sensors do not include a battery, they do not have an absolute time
  '''   reference. Therefore, measures are simply indexed by the absolute run number
  '''   and time relative to the start of the run. Every new power up starts a new run.
  '''   It is however possible to setup an absolute UTC time by software at a given time,
  '''   so that the data logger keeps track of it until it is next powered off.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDataLogger
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const OLDESTRUNINDEX_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const CURRENTRUNINDEX_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const SAMPLINGINTERVAL_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const TIMEUTC_INVALID As Long = YAPI.INVALID_LONG
    Public Const RECORDING_OFF = 0
    Public Const RECORDING_ON = 1
    Public Const RECORDING_INVALID = -1

    Public Const AUTOSTART_OFF = 0
    Public Const AUTOSTART_ON = 1
    Public Const AUTOSTART_INVALID = -1

    Public Const CLEARHISTORY_FALSE = 0
    Public Const CLEARHISTORY_TRUE = 1
    Public Const CLEARHISTORY_INVALID = -1


    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _oldestRunIndex As Long
    Protected _currentRunIndex As Long
    Protected _samplingInterval As Long
    Protected _timeUTC As Long
    Protected _recording As Long
    Protected _autoStart As Long
    Protected _clearHistory As Long

    Public Sub New(ByVal func As String)
      MyBase.new("DataLogger", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _oldestRunIndex = Y_OLDESTRUNINDEX_INVALID
      _currentRunIndex = Y_CURRENTRUNINDEX_INVALID
      _samplingInterval = Y_SAMPLINGINTERVAL_INVALID
      _timeUTC = Y_TIMEUTC_INVALID
      _recording = Y_RECORDING_INVALID
      _autoStart = Y_AUTOSTART_INVALID
      _clearHistory = Y_CLEARHISTORY_INVALID
    End Sub

    Protected Overrides Function _parse(ByRef j As TJSONRECORD) As Integer
      Dim member As TJSONRECORD
      Dim i As Integer
      If (j.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then
        Return -1
      End If
      For i = 0 To j.membercount - 1
        member = j.members(i)
        If (member.name = "logicalName") Then
          _logicalName = member.svalue
        ElseIf (member.name = "advertisedValue") Then
          _advertisedValue = member.svalue
        ElseIf (member.name = "oldestRunIndex") Then
          _oldestRunIndex = CLng(member.ivalue)
        ElseIf (member.name = "currentRunIndex") Then
          _currentRunIndex = CLng(member.ivalue)
        ElseIf (member.name = "samplingInterval") Then
          _samplingInterval = CLng(member.ivalue)
        ElseIf (member.name = "timeUTC") Then
          _timeUTC = CLng(member.ivalue)
        ElseIf (member.name = "recording") Then
          If (member.ivalue > 0) Then _recording = 1 Else _recording = 0
        ElseIf (member.name = "autoStart") Then
          If (member.ivalue > 0) Then _autoStart = 1 Else _autoStart = 0
        ElseIf (member.name = "clearHistory") Then
          If (member.ivalue > 0) Then _clearHistory = 1 Else _clearHistory = 0
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the data logger.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the data logger
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LOGICALNAME_INVALID
        End If
      End If
      Return _logicalName
    End Function

    '''*
    ''' <summary>
    '''   Changes the logical name of the data logger.
    ''' <para>
    '''   You can use <c>yCheckLogicalName()</c>
    '''   prior to this call to make sure that your parameter is valid.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the logical name of the data logger
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_logicalName(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("logicalName", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of the data logger (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the data logger (no more than 6 characters)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ADVERTISEDVALUE_INVALID
        End If
      End If
      Return _advertisedValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the index of the oldest run for which the non-volatile memory still holds recorded data.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the index of the oldest run for which the non-volatile memory still
    '''   holds recorded data
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OLDESTRUNINDEX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_oldestRunIndex() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_OLDESTRUNINDEX_INVALID
        End If
      End If
      Return CType(_oldestRunIndex,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current run number, corresponding to the number of times the module was
    '''   powered on with the dataLogger enabled at some point.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current run number, corresponding to the number of times the module was
    '''   powered on with the dataLogger enabled at some point
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENTRUNINDEX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentRunIndex() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CURRENTRUNINDEX_INVALID
        End If
      End If
      Return CType(_currentRunIndex,Integer)
    End Function

    Public Function get_samplingInterval() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_SAMPLINGINTERVAL_INVALID
        End If
      End If
      Return CType(_samplingInterval,Integer)
    End Function

    Public Function set_samplingInterval(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("samplingInterval", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the Unix timestamp for current UTC time, if known.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the Unix timestamp for current UTC time, if known
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_TIMEUTC_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_timeUTC() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_TIMEUTC_INVALID
        End If
      End If
      Return _timeUTC
    End Function

    '''*
    ''' <summary>
    '''   Changes the current UTC time reference used for recorded data.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current UTC time reference used for recorded data
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_timeUTC(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("timeUTC", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current activation state of the data logger.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_RECORDING_OFF</c> or <c>Y_RECORDING_ON</c>, according to the current activation state
    '''   of the data logger
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RECORDING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_recording() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_RECORDING_INVALID
        End If
      End If
      Return CType(_recording,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the activation state of the data logger to start/stop recording data.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_RECORDING_OFF</c> or <c>Y_RECORDING_ON</c>, according to the activation state of the
    '''   data logger to start/stop recording data
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_recording(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("recording", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the default activation state of the data logger on power up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_AUTOSTART_OFF</c> or <c>Y_AUTOSTART_ON</c>, according to the default activation state
    '''   of the data logger on power up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_AUTOSTART_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_autoStart() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_AUTOSTART_INVALID
        End If
      End If
      Return CType(_autoStart,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the default activation state of the data logger on power up.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_AUTOSTART_OFF</c> or <c>Y_AUTOSTART_ON</c>, according to the default activation state
    '''   of the data logger on power up
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_autoStart(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("autoStart", rest_val)
    End Function

    Public Function get_clearHistory() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CLEARHISTORY_INVALID
        End If
      End If
      Return CType(_clearHistory,Integer)
    End Function

    Public Function set_clearHistory(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("clearHistory", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of data loggers started using <c>yFirstDataLogger()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDataLogger</c> object, corresponding to
    '''   a data logger currently online, or a <c>null</c> pointer
    '''   if there are no more data loggers to enumerate.
    ''' </returns>
    '''/
    Public Function nextDataLogger() as YDataLogger
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindDataLogger(hwid)
    End Function

    '''*
    ''' <summary>
    '''   comment from .
    ''' <para>
    '''   yc definition
    ''' </para>
    ''' </summary>
    '''/
  Public Overloads Sub registerValueCallback(ByVal callback As UpdateCallback)
   If (callback IsNot Nothing) Then
     registerFuncCallback(Me)
   Else
     unregisterFuncCallback(Me)
   End If
   _callback = callback
  End Sub

  Public Sub set_callback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Sub setCallback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Overrides Sub advertiseValue(ByVal value As String)
    If (_callback IsNot Nothing) Then _callback(Me, value)
  End Sub


    '''*
    ''' <summary>
    '''   Retrieves a data logger for a given identifier.
    ''' <para>
    '''   The identifier can be specified using several formats:
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   - FunctionLogicalName
    ''' </para>
    ''' <para>
    '''   - ModuleSerialNumber.FunctionIdentifier
    ''' </para>
    ''' <para>
    '''   - ModuleSerialNumber.FunctionLogicalName
    ''' </para>
    ''' <para>
    '''   - ModuleLogicalName.FunctionIdentifier
    ''' </para>
    ''' <para>
    '''   - ModuleLogicalName.FunctionLogicalName
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   This function does not require that the data logger is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YDataLogger.isOnline()</c> to test if the data logger is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a data logger by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the data logger
    ''' </param>
    ''' <returns>
    '''   a <c>YDataLogger</c> object allowing you to drive the data logger.
    ''' </returns>
    '''/
    Public Shared Function FindDataLogger(ByVal func As String) As YDataLogger
      Dim res As YDataLogger
      If (_DataLoggerCache.ContainsKey(func)) Then
        Return CType(_DataLoggerCache(func), YDataLogger)
      End If
      res = New YDataLogger(func)
      _DataLoggerCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of data loggers currently accessible.
    ''' <para>
    '''   Use the method <c>YDataLogger.nextDataLogger()</c> to iterate on
    '''   next data loggers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDataLogger</c> object, corresponding to
    '''   the first data logger currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstDataLogger() As YDataLogger
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("DataLogger", 0, p, size, neededsize, errmsg)
      Marshal.Copy(p, v_fundescr, 0, 1)
      Marshal.FreeHGlobal(p)

      If (YISERR(err) Or (neededsize = 0)) Then
        Return Nothing
      End If
      serial = ""
      funcId = ""
      funcName = ""
      funcVal = ""
      errmsg = ""
      If (YISERR(yapiGetFunctionInfo(v_fundescr(0), dev, serial, funcId, funcName, funcVal, errmsg))) Then
        Return Nothing
      End If
      Return YDataLogger.FindDataLogger(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YDataLogger implementation)

    Protected _dataLoggerURL As String

    Public Function getData(ByVal runIdx As Long, ByVal timeIdx As Long, ByRef jsondata As TJsonParser) As Integer

      Dim dev As YDevice = Nothing
      Dim errmsg As String = ""
      Dim query As String
      Dim buffer As String = ""
      Dim res As Integer

      If (_dataLoggerURL = "") Then
        _dataLoggerURL = "/logger.json"
      End If
      REM Resolve our reference to our device, load REST API
      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        getData = res
        Exit Function
      End If

      If (timeIdx > 0) Then
        query = "GET " + _dataLoggerURL + "?run=" + LTrim(Str(runIdx)) + "&time=" + LTrim(Str(timeIdx)) + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      Else
        query = "GET " + _dataLoggerURL + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      End If

      res = dev.HTTPRequest(query, buffer, errmsg)
      REM make sure a device scan does not solve the issue
      If (YISERR(res)) Then
      res = yapiUpdateDeviceList(1, errmsg)
      If (YISERR(res)) Then
        getData = res
        Exit Function
      End If

      res = dev.HTTPRequest("GET " + _dataLoggerURL + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10), buffer, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        getData = res
        Exit Function
      End If
      End If

      Try
        jsondata = New TJsonParser(buffer)
      Catch e As Exception
        errmsg = "unexpected JSON structure: " + e.Message
        _throw(YAPI_IO_ERROR, errmsg)
        getData = YAPI_IO_ERROR
        Exit Function
      End Try
      If (jsondata.httpcode = 404 And _dataLoggerURL <> "/dataLogger.json") Then
        REM retry using backward-compatible datalogger URL
        _dataLoggerURL = "/dataLogger.json"
        Return getData(runIdx, timeIdx, jsondata)
      End If
      getData = YAPI_SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Clears the data logger memory and discards all recorded data streams.
    ''' <para>
    '''   This method also resets the current run index to zero.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function forgetAllDataStreams() As Integer
      forgetAllDataStreams = set_clearHistory(Y_CLEARHISTORY_TRUE)
    End Function

    '''*
    ''' <summary>
    '''   Builds a list of all data streams hold by the data logger.
    ''' <para>
    '''   The caller must pass by reference an empty array to hold YDataStream
    '''   objects, and the function fills it with objects describing available
    '''   data sequences.
    ''' </para>
    ''' </summary>
    ''' <param name="v">
    '''   an array of YDataStream objects to be filled in
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function get_dataStreams(ByVal v As List(Of YDataStream)) As Integer

      Dim j As TJsonParser = Nothing
      Dim i, res As Integer
      Dim root, el As TJSONRECORD

      v.Clear()
      res = getData(0, 0, j)
      If (res <> YAPI_SUCCESS) Then
        get_dataStreams = res
        Exit Function
      End If

      root = j.GetRootNode()
      For i = 0 To root.itemcount - 1
        el = root.items(i)
        v.Add(New YDataStream(Me, el.items(0).ivalue, el.items(1).ivalue, el.items(2).ivalue, el.items(3).ivalue))
      Next i

      j = Nothing
      get_dataStreams = YAPI_SUCCESS
    End Function


    Public Sub throw_friend(ByVal errType As System.Int32, ByVal errMsg As String)
      _throw(errType, errMsg)
    End Sub

  End Class

  REM --- (generated code: DataLogger functions)

  '''*
  ''' <summary>
  '''   Retrieves a data logger for a given identifier.
  ''' <para>
  '''   The identifier can be specified using several formats:
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   - FunctionLogicalName
  ''' </para>
  ''' <para>
  '''   - ModuleSerialNumber.FunctionIdentifier
  ''' </para>
  ''' <para>
  '''   - ModuleSerialNumber.FunctionLogicalName
  ''' </para>
  ''' <para>
  '''   - ModuleLogicalName.FunctionIdentifier
  ''' </para>
  ''' <para>
  '''   - ModuleLogicalName.FunctionLogicalName
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   This function does not require that the data logger is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YDataLogger.isOnline()</c> to test if the data logger is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a data logger by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the data logger
  ''' </param>
  ''' <returns>
  '''   a <c>YDataLogger</c> object allowing you to drive the data logger.
  ''' </returns>
  '''/
  Public Function yFindDataLogger(ByVal func As String) As YDataLogger
    Return YDataLogger.FindDataLogger(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of data loggers currently accessible.
  ''' <para>
  '''   Use the method <c>YDataLogger.nextDataLogger()</c> to iterate on
  '''   next data loggers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YDataLogger</c> object, corresponding to
  '''   the first data logger currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstDataLogger() As YDataLogger
    Return YDataLogger.FirstDataLogger()
  End Function

  Private Sub _DataLoggerCleanup()
  End Sub


  REM --- (end of generated code: DataLogger functions)




End Module