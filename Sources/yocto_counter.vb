' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindCounter(), the high-level API for Counter functions
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

Module yocto_counter

    REM --- (YCounter return codes)
    REM --- (end of YCounter return codes)
    REM --- (YCounter dlldef)
    REM --- (end of YCounter dlldef)
   REM --- (YCounter yapiwrapper)
   REM --- (end of YCounter yapiwrapper)
  REM --- (YCounter globals)

  Public Delegate Sub YCounterValueCallback(ByVal func As YCounter, ByVal value As String)
  Public Delegate Sub YCounterTimedReportCallback(ByVal func As YCounter, ByVal measure As YMeasure)
  REM --- (end of YCounter globals)

  REM --- (YCounter class start)

  '''*
  ''' <summary>
  '''   The <c>YCounter</c> class allows you to read and configure Yoctopuce gcounters.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YCounter
    Inherits YSensor
    REM --- (end of YCounter class start)

    REM --- (YCounter definitions)
    REM --- (end of YCounter definitions)

    REM --- (YCounter attributes declaration)
    Protected _valueCallbackCounter As YCounterValueCallback
    Protected _timedReportCallbackCounter As YCounterTimedReportCallback
    REM --- (end of YCounter attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Counter"
      REM --- (YCounter attributes initialization)
      _valueCallbackCounter = Nothing
      _timedReportCallbackCounter = Nothing
      REM --- (end of YCounter attributes initialization)
    End Sub

    REM --- (YCounter private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YCounter private methods declaration)

    REM --- (YCounter public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a counter for a given identifier.
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
    '''   This function does not require that the counter is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YCounter.isOnline()</c> to test if the counter is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a counter by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the counter, for instance
    '''   <c>MyDevice.counter</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YCounter</c> object allowing you to drive the counter.
    ''' </returns>
    '''/
    Public Shared Function FindCounter(func As String) As YCounter
      Dim obj As YCounter
      obj = CType(YFunction._FindFromCache("Counter", func), YCounter)
      If ((obj Is Nothing)) Then
        obj = New YCounter(func)
        YFunction._AddToCache("Counter", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YCounterValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackCounter = callback
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
      If (Not (Me._valueCallbackCounter Is Nothing)) Then
        Me._valueCallbackCounter(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YCounterTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackCounter = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackCounter Is Nothing)) Then
        Me._timedReportCallbackCounter(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of gcounters started using <c>yFirstCounter()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned gcounters order.
    '''   If you want to find a specific a counter, use <c>Counter.findCounter()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCounter</c> object, corresponding to
    '''   a counter currently online, or a <c>Nothing</c> pointer
    '''   if there are no more gcounters to enumerate.
    ''' </returns>
    '''/
    Public Function nextCounter() As YCounter
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YCounter.FindCounter(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of gcounters currently accessible.
    ''' <para>
    '''   Use the method <c>YCounter.nextCounter()</c> to iterate on
    '''   next gcounters.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCounter</c> object, corresponding to
    '''   the first counter currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstCounter() As YCounter
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Counter", 0, p, size, neededsize, errmsg)
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
      Return YCounter.FindCounter(serial + "." + funcId)
    End Function

    REM --- (end of YCounter public methods declaration)

  End Class

  REM --- (YCounter functions)

  '''*
  ''' <summary>
  '''   Retrieves a counter for a given identifier.
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
  '''   This function does not require that the counter is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YCounter.isOnline()</c> to test if the counter is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a counter by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the counter, for instance
  '''   <c>MyDevice.counter</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YCounter</c> object allowing you to drive the counter.
  ''' </returns>
  '''/
  Public Function yFindCounter(ByVal func As String) As YCounter
    Return YCounter.FindCounter(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of gcounters currently accessible.
  ''' <para>
  '''   Use the method <c>YCounter.nextCounter()</c> to iterate on
  '''   next gcounters.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YCounter</c> object, corresponding to
  '''   the first counter currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstCounter() As YCounter
    Return YCounter.FirstCounter()
  End Function


  REM --- (end of YCounter functions)

End Module
