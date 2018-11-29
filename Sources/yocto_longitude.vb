' ********************************************************************
'
'  $Id: yocto_longitude.vb 32908 2018-11-02 10:19:28Z seb $
'
'  Implements yFindLongitude(), the high-level API for Longitude functions
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

Module yocto_longitude

    REM --- (YLongitude return codes)
    REM --- (end of YLongitude return codes)
    REM --- (YLongitude dlldef)
    REM --- (end of YLongitude dlldef)
   REM --- (YLongitude yapiwrapper)
   REM --- (end of YLongitude yapiwrapper)
  REM --- (YLongitude globals)

  Public Delegate Sub YLongitudeValueCallback(ByVal func As YLongitude, ByVal value As String)
  Public Delegate Sub YLongitudeTimedReportCallback(ByVal func As YLongitude, ByVal measure As YMeasure)
  REM --- (end of YLongitude globals)

  REM --- (YLongitude class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce class YLongitude allows you to read the longitude from Yoctopuce
  '''   geolocalization sensors.
  ''' <para>
  '''   It inherits from the YSensor class the core functions to
  '''   read measurements, register callback functions, access the autonomous
  '''   datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YLongitude
    Inherits YSensor
    REM --- (end of YLongitude class start)

    REM --- (YLongitude definitions)
    REM --- (end of YLongitude definitions)

    REM --- (YLongitude attributes declaration)
    Protected _valueCallbackLongitude As YLongitudeValueCallback
    Protected _timedReportCallbackLongitude As YLongitudeTimedReportCallback
    REM --- (end of YLongitude attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Longitude"
      REM --- (YLongitude attributes initialization)
      _valueCallbackLongitude = Nothing
      _timedReportCallbackLongitude = Nothing
      REM --- (end of YLongitude attributes initialization)
    End Sub

    REM --- (YLongitude private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YLongitude private methods declaration)

    REM --- (YLongitude public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a longitude sensor for a given identifier.
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
    '''   This function does not require that the longitude sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YLongitude.isOnline()</c> to test if the longitude sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a longitude sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the longitude sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YLongitude</c> object allowing you to drive the longitude sensor.
    ''' </returns>
    '''/
    Public Shared Function FindLongitude(func As String) As YLongitude
      Dim obj As YLongitude
      obj = CType(YFunction._FindFromCache("Longitude", func), YLongitude)
      If ((obj Is Nothing)) Then
        obj = New YLongitude(func)
        YFunction._AddToCache("Longitude", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YLongitudeValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackLongitude = callback
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
      If (Not (Me._valueCallbackLongitude Is Nothing)) Then
        Me._valueCallbackLongitude(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YLongitudeTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackLongitude = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackLongitude Is Nothing)) Then
        Me._timedReportCallbackLongitude(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of longitude sensors started using <c>yFirstLongitude()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned longitude sensors order.
    '''   If you want to find a specific a longitude sensor, use <c>Longitude.findLongitude()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLongitude</c> object, corresponding to
    '''   a longitude sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more longitude sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextLongitude() As YLongitude
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YLongitude.FindLongitude(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of longitude sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YLongitude.nextLongitude()</c> to iterate on
    '''   next longitude sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLongitude</c> object, corresponding to
    '''   the first longitude sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstLongitude() As YLongitude
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Longitude", 0, p, size, neededsize, errmsg)
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
      Return YLongitude.FindLongitude(serial + "." + funcId)
    End Function

    REM --- (end of YLongitude public methods declaration)

  End Class

  REM --- (YLongitude functions)

  '''*
  ''' <summary>
  '''   Retrieves a longitude sensor for a given identifier.
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
  '''   This function does not require that the longitude sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YLongitude.isOnline()</c> to test if the longitude sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a longitude sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the longitude sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YLongitude</c> object allowing you to drive the longitude sensor.
  ''' </returns>
  '''/
  Public Function yFindLongitude(ByVal func As String) As YLongitude
    Return YLongitude.FindLongitude(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of longitude sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YLongitude.nextLongitude()</c> to iterate on
  '''   next longitude sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YLongitude</c> object, corresponding to
  '''   the first longitude sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstLongitude() As YLongitude
    Return YLongitude.FirstLongitude()
  End Function


  REM --- (end of YLongitude functions)

End Module
