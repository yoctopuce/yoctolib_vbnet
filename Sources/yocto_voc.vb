'*********************************************************************
'*
'* $Id: yocto_voc.vb 11112 2013-04-16 14:51:20Z mvuilleu $
'*
'* Implements yFindVoc(), the high-level API for Voc functions
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
'*********************************************************************/


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_voc

  REM --- (YVoc definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YVoc, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_UNIT_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CURRENTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_LOWESTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_HIGHESTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_CURRENTRAWVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_RESOLUTION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING


  REM --- (end of YVoc definitions)

  REM --- (YVoc implementation)

  Private _VocCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YVoc
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const UNIT_INVALID As String = YAPI.INVALID_STRING
    Public Const CURRENTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const LOWESTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const HIGHESTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const CURRENTRAWVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const RESOLUTION_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _unit As String
    Protected _currentValue As Double
    Protected _lowestValue As Double
    Protected _highestValue As Double
    Protected _currentRawValue As Double
    Protected _resolution As Double
    Protected _calibrationParam As String
    Protected _calibrationOffset As Long

    Public Sub New(ByVal func As String)
      MyBase.new("Voc", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _unit = Y_UNIT_INVALID
      _currentValue = Y_CURRENTVALUE_INVALID
      _lowestValue = Y_LOWESTVALUE_INVALID
      _highestValue = Y_HIGHESTVALUE_INVALID
      _currentRawValue = Y_CURRENTRAWVALUE_INVALID
      _resolution = Y_RESOLUTION_INVALID
      _calibrationParam = Y_CALIBRATIONPARAM_INVALID
      _calibrationOffset = 0
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
        ElseIf (member.name = "unit") Then
          _unit = member.svalue
        ElseIf (member.name = "currentValue") Then
          _currentValue = Math.Round(member.ivalue/6553.6) / 10
        ElseIf (member.name = "lowestValue") Then
          _lowestValue = Math.Round(member.ivalue/6553.6) / 10
        ElseIf (member.name = "highestValue") Then
          _highestValue = Math.Round(member.ivalue/6553.6) / 10
        ElseIf (member.name = "currentRawValue") Then
          _currentRawValue = member.ivalue/65536.0
        ElseIf (member.name = "resolution") Then
          _resolution = 1.0 / Math.Round(65536.0/member.ivalue)
        ElseIf (member.name = "calibrationParam") Then
          _calibrationParam = member.svalue
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the Volatile Organic Compound sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the Volatile Organic Compound sensor
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
    '''   Changes the logical name of the Volatile Organic Compound sensor.
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
    '''   a string corresponding to the logical name of the Volatile Organic Compound sensor
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
    '''   Returns the current value of the Volatile Organic Compound sensor (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the Volatile Organic Compound sensor (no more than 6 characters)
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
    '''   Returns the measuring unit for the measured value.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the measuring unit for the measured value
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_UNIT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_unit() As String
      If (_unit = Y_UNIT_INVALID) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_UNIT_INVALID
        End If
      End If
      Return _unit
    End Function

    '''*
    ''' <summary>
    '''   Returns the current measured value.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current measured value
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentValue() As Double
       dim res as double
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CURRENTVALUE_INVALID
        End If
      End If
     res = YAPI._applyCalibration(_currentRawValue, _calibrationParam, _calibrationOffset, _resolution)
     if (res <> CURRENTVALUE_INVALID)  
         Return res
      End If
      Return _currentValue
    End Function

    '''*
    ''' <summary>
    '''   Changes the recorded minimal value observed.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the recorded minimal value observed
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
    Public Function set_lowestValue(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval*65536.0)))
      Return _setAttr("lowestValue", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the minimal value observed.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the minimal value observed
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOWESTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lowestValue() As Double
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LOWESTVALUE_INVALID
        End If
      End If
      Return _lowestValue
    End Function

    '''*
    ''' <summary>
    '''   Changes the recorded maximal value observed.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the recorded maximal value observed
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
    Public Function set_highestValue(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval*65536.0)))
      Return _setAttr("highestValue", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the maximal value observed.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the maximal value observed
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_HIGHESTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_highestValue() As Double
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_HIGHESTVALUE_INVALID
        End If
      End If
      Return _highestValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the unrounded and uncalibrated raw value returned by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the unrounded and uncalibrated raw value returned by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENTRAWVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentRawValue() As Double
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CURRENTRAWVALUE_INVALID
        End If
      End If
      Return _currentRawValue
    End Function

    Public Function set_resolution(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval*65536.0)))
      Return _setAttr("resolution", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the resolution of the measured values.
    ''' <para>
    '''   The resolution corresponds to the numerical precision
    '''   of the values, which is not always the same as the actual precision of the sensor.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the resolution of the measured values
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RESOLUTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_resolution() As Double
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_RESOLUTION_INVALID
        End If
      End If
      Return _resolution
    End Function

    Public Function get_calibrationParam() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CALIBRATIONPARAM_INVALID
        End If
      End If
      Return _calibrationParam
    End Function

    Public Function set_calibrationParam(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("calibrationParam", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Configures error correction data points, in particular to compensate for
    '''   a possible perturbation of the measure caused by an enclosure.
    ''' <para>
    '''   It is possible
    '''   to configure up to five correction points. Correction points must be provided
    '''   in ascending order, and be in the range of the sensor. The device will automatically
    '''   perform a lineat interpolatation of the error correction between specified
    '''   points. Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    '''   For more information on advanced capabilities to refine the calibration of
    '''   sensors, please contact support@yoctopuce.com.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="rawValues">
    '''   array of floating point numbers, corresponding to the raw
    '''   values returned by the sensor for the correction points.
    ''' </param>
    ''' <param name="refValues">
    '''   array of floating point numbers, corresponding to the corrected
    '''   values for the correction points.
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
    Public Function calibrateFromPoints(ByVal rawValues As double(),ByVal refValues As double()) As Integer
      Dim rest_val As String
      rest_val = YAPI._encodeCalibrationPoints(rawValues,refValues,Me._resolution,Me._calibrationOffset,Me._calibrationParam)
      Return _setAttr("calibrationParam", rest_val)
    End Function

    Public Function loadCalibrationPoints(ByRef rawValues As double(),ByRef refValues As double()) As Integer
      Return YAPI._decodeCalibrationPoints(Me._calibrationParam,Nothing,rawValues,refValues, Me._resolution, Me._calibrationOffset)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of Volatile Organic Compound sensors started using <c>yFirstVoc()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVoc</c> object, corresponding to
    '''   a Volatile Organic Compound sensor currently online, or a <c>null</c> pointer
    '''   if there are no more Volatile Organic Compound sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextVoc() as YVoc
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindVoc(hwid)
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
    '''   Retrieves a Volatile Organic Compound sensor for a given identifier.
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
    '''   This function does not require that the Volatile Organic Compound sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YVoc.isOnline()</c> to test if the Volatile Organic Compound sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a Volatile Organic Compound sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the Volatile Organic Compound sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YVoc</c> object allowing you to drive the Volatile Organic Compound sensor.
    ''' </returns>
    '''/
    Public Shared Function FindVoc(ByVal func As String) As YVoc
      Dim res As YVoc
      If (_VocCache.ContainsKey(func)) Then
        Return CType(_VocCache(func), YVoc)
      End If
      res = New YVoc(func)
      _VocCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of Volatile Organic Compound sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YVoc.nextVoc()</c> to iterate on
    '''   next Volatile Organic Compound sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVoc</c> object, corresponding to
    '''   the first Volatile Organic Compound sensor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstVoc() As YVoc
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Voc", 0, p, size, neededsize, errmsg)
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
      Return YVoc.FindVoc(serial + "." + funcId)
    End Function

    REM --- (end of YVoc implementation)

  End Class

  REM --- (Voc functions)

  '''*
  ''' <summary>
  '''   Retrieves a Volatile Organic Compound sensor for a given identifier.
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
  '''   This function does not require that the Volatile Organic Compound sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YVoc.isOnline()</c> to test if the Volatile Organic Compound sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a Volatile Organic Compound sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the Volatile Organic Compound sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YVoc</c> object allowing you to drive the Volatile Organic Compound sensor.
  ''' </returns>
  '''/
  Public Function yFindVoc(ByVal func As String) As YVoc
    Return YVoc.FindVoc(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of Volatile Organic Compound sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YVoc.nextVoc()</c> to iterate on
  '''   next Volatile Organic Compound sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YVoc</c> object, corresponding to
  '''   the first Volatile Organic Compound sensor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstVoc() As YVoc
    Return YVoc.FirstVoc()
  End Function

  Private Sub _VocCleanup()
  End Sub


  REM --- (end of Voc functions)

End Module
