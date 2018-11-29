' ********************************************************************
'
'  $Id: yocto_latitude.vb 32908 2018-11-02 10:19:28Z seb $
'
'  Implements yFindLatitude(), the high-level API for Latitude functions
'
'  - - - - - - - - - License information: - - - - - - - - -
'
'  Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'
'  Yoctopuce Sarl (hereafter Licensor) grants to you a perpetual
'  non-exclusive license to use, modify, copy and integrate this
'  file into your software for the sole purpose of interfacing
'  with Yoctopuce products.
'
'  You may reproduce and distribute copies of this file in
'  source or object form, as long as the sole purpose of this
'  code is to interface with Yoctopuce products. You must retain
'  this notice in the distributed source file.
'
'  You should refer to Yoctopuce General Terms and Conditions
'  for additional information regarding your rights and
'  obligations.
'
'  THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'  WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING
'  WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS
'  FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'  EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'  INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA,
'  COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR
'  SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT
'  LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'  CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'  BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'  WARRANTY, OR OTHERWISE.
'
' *********************************************************************


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_latitude

    REM --- (YLatitude return codes)
    REM --- (end of YLatitude return codes)
    REM --- (YLatitude dlldef)
    REM --- (end of YLatitude dlldef)
   REM --- (YLatitude yapiwrapper)
   REM --- (end of YLatitude yapiwrapper)
  REM --- (YLatitude globals)

  Public Delegate Sub YLatitudeValueCallback(ByVal func As YLatitude, ByVal value As String)
  Public Delegate Sub YLatitudeTimedReportCallback(ByVal func As YLatitude, ByVal measure As YMeasure)
  REM --- (end of YLatitude globals)

  REM --- (YLatitude class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce class YLatitude allows you to read the latitude from Yoctopuce
  '''   geolocalization sensors.
  ''' <para>
  '''   It inherits from the YSensor class the core functions to
  '''   read measurements, to register callback functions, to access the autonomous
  '''   datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YLatitude
    Inherits YSensor
    REM --- (end of YLatitude class start)

    REM --- (YLatitude definitions)
    REM --- (end of YLatitude definitions)

    REM --- (YLatitude attributes declaration)
    Protected _valueCallbackLatitude As YLatitudeValueCallback
    Protected _timedReportCallbackLatitude As YLatitudeTimedReportCallback
    REM --- (end of YLatitude attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Latitude"
      REM --- (YLatitude attributes initialization)
      _valueCallbackLatitude = Nothing
      _timedReportCallbackLatitude = Nothing
      REM --- (end of YLatitude attributes initialization)
    End Sub

    REM --- (YLatitude private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YLatitude private methods declaration)

    REM --- (YLatitude public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a latitude sensor for a given identifier.
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
    '''   This function does not require that the latitude sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YLatitude.isOnline()</c> to test if the latitude sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a latitude sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' <para>
    '''   If a call to this object's is_online() method returns FALSE although
    '''   you are certain that the matching device is plugged, make sure that you did
    '''   call registerHub() at application initialization time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the latitude sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YLatitude</c> object allowing you to drive the latitude sensor.
    ''' </returns>
    '''/
    Public Shared Function FindLatitude(func As String) As YLatitude
      Dim obj As YLatitude
      obj = CType(YFunction._FindFromCache("Latitude", func), YLatitude)
      If ((obj Is Nothing)) Then
        obj = New YLatitude(func)
        YFunction._AddToCache("Latitude", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerValueCallback(callback As YLatitudeValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackLatitude = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) AndAlso Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overrides Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackLatitude Is Nothing)) Then
        Me._valueCallbackLatitude(Me, value)
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
    '''   one of these two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YLatitudeTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackLatitude = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackLatitude Is Nothing)) Then
        Me._timedReportCallbackLatitude(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of latitude sensors started using <c>yFirstLatitude()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned latitude sensors order.
    '''   If you want to find a specific a latitude sensor, use <c>Latitude.findLatitude()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLatitude</c> object, corresponding to
    '''   a latitude sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more latitude sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextLatitude() As YLatitude
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YLatitude.FindLatitude(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of latitude sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YLatitude.nextLatitude()</c> to iterate on
    '''   next latitude sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLatitude</c> object, corresponding to
    '''   the first latitude sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstLatitude() As YLatitude
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Latitude", 0, p, size, neededsize, errmsg)
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
      Return YLatitude.FindLatitude(serial + "." + funcId)
    End Function

    REM --- (end of YLatitude public methods declaration)

  End Class

  REM --- (YLatitude functions)

  '''*
  ''' <summary>
  '''   Retrieves a latitude sensor for a given identifier.
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
  '''   This function does not require that the latitude sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YLatitude.isOnline()</c> to test if the latitude sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a latitude sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' <para>
  '''   If a call to this object's is_online() method returns FALSE although
  '''   you are certain that the matching device is plugged, make sure that you did
  '''   call registerHub() at application initialization time.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the latitude sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YLatitude</c> object allowing you to drive the latitude sensor.
  ''' </returns>
  '''/
  Public Function yFindLatitude(ByVal func As String) As YLatitude
    Return YLatitude.FindLatitude(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of latitude sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YLatitude.nextLatitude()</c> to iterate on
  '''   next latitude sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YLatitude</c> object, corresponding to
  '''   the first latitude sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstLatitude() As YLatitude
    Return YLatitude.FirstLatitude()
  End Function


  REM --- (end of YLatitude functions)

End Module
