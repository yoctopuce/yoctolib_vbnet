'*********************************************************************
'*
'* $Id: yocto_altitude.vb 17356 2014-08-29 14:38:39Z seb $
'*
'* Implements yFindAltitude(), the high-level API for Altitude functions
'*
'* - - - - - - - - - License information: - - - - - - - - - 
'*
'*  Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'*
'*  Yoctopuce Sarl (hereafter Licensor) grants to you a perpetual
'*  non-exclusive license to use, modify, copy and integrate this
'*  file into your software for the sole purpose of interfacing
'*  with Yoctopuce products.
'*
'*  You may reproduce and distribute copies of this file in
'*  source or object form, as long as the sole purpose of this
'*  code is to interface with Yoctopuce products. You must retain
'*  this notice in the distributed source file.
'*
'*  You should refer to Yoctopuce General Terms and Conditions
'*  for additional information regarding your rights and
'*  obligations.
'*
'*  THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'*  WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING 
'*  WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS
'*  FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'*  EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'*  INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA,
'*  COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR 
'*  SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT 
'*  LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'*  CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'*  BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'*  WARRANTY, OR OTHERWISE.
'*
'*********************************************************************/


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_altitude

    REM --- (YAltitude return codes)
    REM --- (end of YAltitude return codes)
    REM --- (YAltitude dlldef)
    REM --- (end of YAltitude dlldef)
  REM --- (YAltitude globals)

  Public Const Y_QNH_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YAltitudeValueCallback(ByVal func As YAltitude, ByVal value As String)
  Public Delegate Sub YAltitudeTimedReportCallback(ByVal func As YAltitude, ByVal measure As YMeasure)
  REM --- (end of YAltitude globals)

  REM --- (YAltitude class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YAltitude
    Inherits YSensor
    REM --- (end of YAltitude class start)

    REM --- (YAltitude definitions)
    Public Const QNH_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YAltitude definitions)

    REM --- (YAltitude attributes declaration)
    Protected _qnh As Double
    Protected _valueCallbackAltitude As YAltitudeValueCallback
    Protected _timedReportCallbackAltitude As YAltitudeTimedReportCallback
    REM --- (end of YAltitude attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Altitude"
      REM --- (YAltitude attributes initialization)
      _qnh = QNH_INVALID
      _valueCallbackAltitude = Nothing
      _timedReportCallbackAltitude = Nothing
      REM --- (end of YAltitude attributes initialization)
    End Sub

    REM --- (YAltitude private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "qnh") Then
        _qnh = Math.Round(member.ivalue * 1000.0 / 65536.0) / 1000.0
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YAltitude private methods declaration)

    REM --- (YAltitude public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the current estimated altitude.
    ''' <para>
    '''   This allows to compensate for
    '''   ambient pressure variations and to work in relative mode.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the current estimated altitude
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
    Public Function set_currentValue(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("currentValue", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the barometric pressure adjusted to sea level used to compute
    '''   the altitude (QNH).
    ''' <para>
    '''   This enables you to compensate for atmospheric pressure
    '''   changes due to weather conditions.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the barometric pressure adjusted to sea level used to compute
    '''   the altitude (QNH)
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
    Public Function set_qnh(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("qnh", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the barometric pressure adjusted to sea level used to compute
    '''   the altitude (QNH).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the barometric pressure adjusted to sea level used to compute
    '''   the altitude (QNH)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_QNH_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_qnh() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return QNH_INVALID
        End If
      End If
      Return Me._qnh
    End Function

    '''*
    ''' <summary>
    '''   Retrieves an altimeter for a given identifier.
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
    '''   This function does not require that the altimeter is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YAltitude.isOnline()</c> to test if the altimeter is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an altimeter by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the altimeter
    ''' </param>
    ''' <returns>
    '''   a <c>YAltitude</c> object allowing you to drive the altimeter.
    ''' </returns>
    '''/
    Public Shared Function FindAltitude(func As String) As YAltitude
      Dim obj As YAltitude
      obj = CType(YFunction._FindFromCache("Altitude", func), YAltitude)
      If ((obj Is Nothing)) Then
        obj = New YAltitude(func)
        YFunction._AddToCache("Altitude", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerValueCallback(callback As YAltitudeValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackAltitude = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) And Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overrides Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackAltitude Is Nothing)) Then
        Me._valueCallbackAltitude(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every periodic timed notification.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YAltitudeTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me, True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me, False)
      End If
      Me._timedReportCallbackAltitude = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackAltitude Is Nothing)) Then
        Me._timedReportCallbackAltitude(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of altimeters started using <c>yFirstAltitude()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAltitude</c> object, corresponding to
    '''   an altimeter currently online, or a <c>null</c> pointer
    '''   if there are no more altimeters to enumerate.
    ''' </returns>
    '''/
    Public Function nextAltitude() As YAltitude
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YAltitude.FindAltitude(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of altimeters currently accessible.
    ''' <para>
    '''   Use the method <c>YAltitude.nextAltitude()</c> to iterate on
    '''   next altimeters.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAltitude</c> object, corresponding to
    '''   the first altimeter currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstAltitude() As YAltitude
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Altitude", 0, p, size, neededsize, errmsg)
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
      Return YAltitude.FindAltitude(serial + "." + funcId)
    End Function

    REM --- (end of YAltitude public methods declaration)

  End Class

  REM --- (Altitude functions)

  '''*
  ''' <summary>
  '''   Retrieves an altimeter for a given identifier.
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
  '''   This function does not require that the altimeter is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YAltitude.isOnline()</c> to test if the altimeter is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an altimeter by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the altimeter
  ''' </param>
  ''' <returns>
  '''   a <c>YAltitude</c> object allowing you to drive the altimeter.
  ''' </returns>
  '''/
  Public Function yFindAltitude(ByVal func As String) As YAltitude
    Return YAltitude.FindAltitude(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of altimeters currently accessible.
  ''' <para>
  '''   Use the method <c>YAltitude.nextAltitude()</c> to iterate on
  '''   next altimeters.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YAltitude</c> object, corresponding to
  '''   the first altimeter currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstAltitude() As YAltitude
    Return YAltitude.FirstAltitude()
  End Function


  REM --- (end of Altitude functions)

End Module
