'*********************************************************************
'*
'* $Id: yocto_humidity.vb 19575 2015-03-04 10:42:56Z seb $
'*
'* Implements yFindHumidity(), the high-level API for Humidity functions
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

Module yocto_humidity

    REM --- (YHumidity return codes)
    REM --- (end of YHumidity return codes)
    REM --- (YHumidity dlldef)
    REM --- (end of YHumidity dlldef)
  REM --- (YHumidity globals)

  Public Delegate Sub YHumidityValueCallback(ByVal func As YHumidity, ByVal value As String)
  Public Delegate Sub YHumidityTimedReportCallback(ByVal func As YHumidity, ByVal measure As YMeasure)
  REM --- (end of YHumidity globals)

  REM --- (YHumidity class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce class YHumidity allows you to read and configure Yoctopuce humidity
  '''   sensors.
  ''' <para>
  '''   It inherits from YSensor class the core functions to read measurements,
  '''   register callback functions, access to the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YHumidity
    Inherits YSensor
    REM --- (end of YHumidity class start)

    REM --- (YHumidity definitions)
    REM --- (end of YHumidity definitions)

    REM --- (YHumidity attributes declaration)
    Protected _valueCallbackHumidity As YHumidityValueCallback
    Protected _timedReportCallbackHumidity As YHumidityTimedReportCallback
    REM --- (end of YHumidity attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Humidity"
      REM --- (YHumidity attributes initialization)
      _valueCallbackHumidity = Nothing
      _timedReportCallbackHumidity = Nothing
      REM --- (end of YHumidity attributes initialization)
    End Sub

    REM --- (YHumidity private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YHumidity private methods declaration)

    REM --- (YHumidity public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a humidity sensor for a given identifier.
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
    '''   This function does not require that the humidity sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YHumidity.isOnline()</c> to test if the humidity sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a humidity sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the humidity sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YHumidity</c> object allowing you to drive the humidity sensor.
    ''' </returns>
    '''/
    Public Shared Function FindHumidity(func As String) As YHumidity
      Dim obj As YHumidity
      obj = CType(YFunction._FindFromCache("Humidity", func), YHumidity)
      If ((obj Is Nothing)) Then
        obj = New YHumidity(func)
        YFunction._AddToCache("Humidity", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YHumidityValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackHumidity = callback
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
      If (Not (Me._valueCallbackHumidity Is Nothing)) Then
        Me._valueCallbackHumidity(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YHumidityTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me, True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me, False)
      End If
      Me._timedReportCallbackHumidity = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackHumidity Is Nothing)) Then
        Me._timedReportCallbackHumidity(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of humidity sensors started using <c>yFirstHumidity()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHumidity</c> object, corresponding to
    '''   a humidity sensor currently online, or a <c>null</c> pointer
    '''   if there are no more humidity sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextHumidity() As YHumidity
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YHumidity.FindHumidity(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of humidity sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YHumidity.nextHumidity()</c> to iterate on
    '''   next humidity sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHumidity</c> object, corresponding to
    '''   the first humidity sensor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstHumidity() As YHumidity
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Humidity", 0, p, size, neededsize, errmsg)
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
      Return YHumidity.FindHumidity(serial + "." + funcId)
    End Function

    REM --- (end of YHumidity public methods declaration)

  End Class

  REM --- (Humidity functions)

  '''*
  ''' <summary>
  '''   Retrieves a humidity sensor for a given identifier.
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
  '''   This function does not require that the humidity sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YHumidity.isOnline()</c> to test if the humidity sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a humidity sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the humidity sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YHumidity</c> object allowing you to drive the humidity sensor.
  ''' </returns>
  '''/
  Public Function yFindHumidity(ByVal func As String) As YHumidity
    Return YHumidity.FindHumidity(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of humidity sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YHumidity.nextHumidity()</c> to iterate on
  '''   next humidity sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YHumidity</c> object, corresponding to
  '''   the first humidity sensor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstHumidity() As YHumidity
    Return YHumidity.FirstHumidity()
  End Function


  REM --- (end of Humidity functions)

End Module
