' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindGroundSpeed(), the high-level API for GroundSpeed functions
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

Module yocto_groundspeed

    REM --- (YGroundSpeed return codes)
    REM --- (end of YGroundSpeed return codes)
    REM --- (YGroundSpeed dlldef)
    REM --- (end of YGroundSpeed dlldef)
   REM --- (YGroundSpeed yapiwrapper)
   REM --- (end of YGroundSpeed yapiwrapper)
  REM --- (YGroundSpeed globals)

  Public Delegate Sub YGroundSpeedValueCallback(ByVal func As YGroundSpeed, ByVal value As String)
  Public Delegate Sub YGroundSpeedTimedReportCallback(ByVal func As YGroundSpeed, ByVal measure As YMeasure)
  REM --- (end of YGroundSpeed globals)

  REM --- (YGroundSpeed class start)

  '''*
  ''' <summary>
  '''   The <c>YGroundSpeed</c> class allows you to read and configure Yoctopuce ground speed sensors.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YGroundSpeed
    Inherits YSensor
    REM --- (end of YGroundSpeed class start)

    REM --- (YGroundSpeed definitions)
    REM --- (end of YGroundSpeed definitions)

    REM --- (YGroundSpeed attributes declaration)
    Protected _valueCallbackGroundSpeed As YGroundSpeedValueCallback
    Protected _timedReportCallbackGroundSpeed As YGroundSpeedTimedReportCallback
    REM --- (end of YGroundSpeed attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "GroundSpeed"
      REM --- (YGroundSpeed attributes initialization)
      _valueCallbackGroundSpeed = Nothing
      _timedReportCallbackGroundSpeed = Nothing
      REM --- (end of YGroundSpeed attributes initialization)
    End Sub

    REM --- (YGroundSpeed private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YGroundSpeed private methods declaration)

    REM --- (YGroundSpeed public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a ground speed sensor for a given identifier.
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
    '''   This function does not require that the ground speed sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YGroundSpeed.isOnline()</c> to test if the ground speed sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a ground speed sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the ground speed sensor, for instance
    '''   <c>YGNSSMK2.groundSpeed</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YGroundSpeed</c> object allowing you to drive the ground speed sensor.
    ''' </returns>
    '''/
    Public Shared Function FindGroundSpeed(func As String) As YGroundSpeed
      Dim obj As YGroundSpeed
      obj = CType(YFunction._FindFromCache("GroundSpeed", func), YGroundSpeed)
      If ((obj Is Nothing)) Then
        obj = New YGroundSpeed(func)
        YFunction._AddToCache("GroundSpeed", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YGroundSpeedValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackGroundSpeed = callback
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
      If (Not (Me._valueCallbackGroundSpeed Is Nothing)) Then
        Me._valueCallbackGroundSpeed(Me, value)
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
    '''   arguments: the function object of which the value has changed, and an <c>YMeasure</c> object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YGroundSpeedTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackGroundSpeed = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackGroundSpeed Is Nothing)) Then
        Me._timedReportCallbackGroundSpeed(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of ground speed sensors started using <c>yFirstGroundSpeed()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned ground speed sensors order.
    '''   If you want to find a specific a ground speed sensor, use <c>GroundSpeed.findGroundSpeed()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGroundSpeed</c> object, corresponding to
    '''   a ground speed sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more ground speed sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextGroundSpeed() As YGroundSpeed
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YGroundSpeed.FindGroundSpeed(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of ground speed sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YGroundSpeed.nextGroundSpeed()</c> to iterate on
    '''   next ground speed sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGroundSpeed</c> object, corresponding to
    '''   the first ground speed sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstGroundSpeed() As YGroundSpeed
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("GroundSpeed", 0, p, size, neededsize, errmsg)
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
      Return YGroundSpeed.FindGroundSpeed(serial + "." + funcId)
    End Function

    REM --- (end of YGroundSpeed public methods declaration)

  End Class

  REM --- (YGroundSpeed functions)

  '''*
  ''' <summary>
  '''   Retrieves a ground speed sensor for a given identifier.
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
  '''   This function does not require that the ground speed sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YGroundSpeed.isOnline()</c> to test if the ground speed sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a ground speed sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the ground speed sensor, for instance
  '''   <c>YGNSSMK2.groundSpeed</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YGroundSpeed</c> object allowing you to drive the ground speed sensor.
  ''' </returns>
  '''/
  Public Function yFindGroundSpeed(ByVal func As String) As YGroundSpeed
    Return YGroundSpeed.FindGroundSpeed(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of ground speed sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YGroundSpeed.nextGroundSpeed()</c> to iterate on
  '''   next ground speed sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YGroundSpeed</c> object, corresponding to
  '''   the first ground speed sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstGroundSpeed() As YGroundSpeed
    Return YGroundSpeed.FirstGroundSpeed()
  End Function


  REM --- (end of YGroundSpeed functions)

End Module
