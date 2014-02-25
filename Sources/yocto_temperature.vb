'*********************************************************************
'*
'* $Id: yocto_temperature.vb 15039 2014-02-24 11:22:11Z seb $
'*
'* Implements yFindTemperature(), the high-level API for Temperature functions
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

Module yocto_temperature

    REM --- (YTemperature return codes)
    REM --- (end of YTemperature return codes)
  REM --- (YTemperature globals)

  Public Const Y_SENSORTYPE_DIGITAL = 0
  Public Const Y_SENSORTYPE_TYPE_K = 1
  Public Const Y_SENSORTYPE_TYPE_E = 2
  Public Const Y_SENSORTYPE_TYPE_J = 3
  Public Const Y_SENSORTYPE_TYPE_N = 4
  Public Const Y_SENSORTYPE_TYPE_R = 5
  Public Const Y_SENSORTYPE_TYPE_S = 6
  Public Const Y_SENSORTYPE_TYPE_T = 7
  Public Const Y_SENSORTYPE_PT100_4WIRES = 8
  Public Const Y_SENSORTYPE_PT100_3WIRES = 9
  Public Const Y_SENSORTYPE_PT100_2WIRES = 10
  Public Const Y_SENSORTYPE_INVALID = -1

  Public Delegate Sub YTemperatureValueCallback(ByVal func As YTemperature, ByVal value As String)
  Public Delegate Sub YTemperatureTimedReportCallback(ByVal func As YTemperature, ByVal measure As YMeasure)
  REM --- (end of YTemperature globals)

  REM --- (YTemperature class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YTemperature
    Inherits YSensor
    REM --- (end of YTemperature class start)

    REM --- (YTemperature definitions)
    Public Const SENSORTYPE_DIGITAL = 0
    Public Const SENSORTYPE_TYPE_K = 1
    Public Const SENSORTYPE_TYPE_E = 2
    Public Const SENSORTYPE_TYPE_J = 3
    Public Const SENSORTYPE_TYPE_N = 4
    Public Const SENSORTYPE_TYPE_R = 5
    Public Const SENSORTYPE_TYPE_S = 6
    Public Const SENSORTYPE_TYPE_T = 7
    Public Const SENSORTYPE_PT100_4WIRES = 8
    Public Const SENSORTYPE_PT100_3WIRES = 9
    Public Const SENSORTYPE_PT100_2WIRES = 10
    Public Const SENSORTYPE_INVALID = -1

    REM --- (end of YTemperature definitions)

    REM --- (YTemperature attributes declaration)
    Protected _sensorType As Integer
    Protected _valueCallbackTemperature As YTemperatureValueCallback
    Protected _timedReportCallbackTemperature As YTemperatureTimedReportCallback
    REM --- (end of YTemperature attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Temperature"
      REM --- (YTemperature attributes initialization)
      _sensorType = SENSORTYPE_INVALID
      _valueCallbackTemperature = Nothing
      _timedReportCallbackTemperature = Nothing
      REM --- (end of YTemperature attributes initialization)
    End Sub

  REM --- (YTemperature private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "sensorType") Then
        _sensorType = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YTemperature private methods declaration)

    REM --- (YTemperature public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the temperature sensor type.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_SENSORTYPE_DIGITAL</c>, <c>Y_SENSORTYPE_TYPE_K</c>, <c>Y_SENSORTYPE_TYPE_E</c>,
    '''   <c>Y_SENSORTYPE_TYPE_J</c>, <c>Y_SENSORTYPE_TYPE_N</c>, <c>Y_SENSORTYPE_TYPE_R</c>,
    '''   <c>Y_SENSORTYPE_TYPE_S</c>, <c>Y_SENSORTYPE_TYPE_T</c>, <c>Y_SENSORTYPE_PT100_4WIRES</c>,
    '''   <c>Y_SENSORTYPE_PT100_3WIRES</c> and <c>Y_SENSORTYPE_PT100_2WIRES</c> corresponding to the
    '''   temperature sensor type
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SENSORTYPE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_sensorType() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SENSORTYPE_INVALID
        End If
      End If
      Return Me._sensorType
    End Function


    '''*
    ''' <summary>
    '''   Modify the temperature sensor type.
    ''' <para>
    '''   This function is used to
    '''   to define the type of thermocouple (K,E...) used with the device.
    '''   This will have no effect if module is using a digital sensor.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_SENSORTYPE_DIGITAL</c>, <c>Y_SENSORTYPE_TYPE_K</c>, <c>Y_SENSORTYPE_TYPE_E</c>,
    '''   <c>Y_SENSORTYPE_TYPE_J</c>, <c>Y_SENSORTYPE_TYPE_N</c>, <c>Y_SENSORTYPE_TYPE_R</c>,
    '''   <c>Y_SENSORTYPE_TYPE_S</c>, <c>Y_SENSORTYPE_TYPE_T</c>, <c>Y_SENSORTYPE_PT100_4WIRES</c>,
    '''   <c>Y_SENSORTYPE_PT100_3WIRES</c> and <c>Y_SENSORTYPE_PT100_2WIRES</c>
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
    Public Function set_sensorType(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("sensorType", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a temperature sensor for a given identifier.
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
    '''   This function does not require that the temperature sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YTemperature.isOnline()</c> to test if the temperature sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a temperature sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the temperature sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YTemperature</c> object allowing you to drive the temperature sensor.
    ''' </returns>
    '''/
    Public Shared Function FindTemperature(func As String) As YTemperature
      Dim obj As YTemperature
      obj = CType(YFunction._FindFromCache("Temperature", func), YTemperature)
      If ((obj Is Nothing)) Then
        obj = New YTemperature(func)
        YFunction._AddToCache("Temperature", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YTemperatureValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me , True)
      Else
        YFunction._UpdateValueCallbackList(Me , False)
      End If
      Me._valueCallbackTemperature = callback
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
      If (Not (Me._valueCallbackTemperature Is Nothing)) Then
        Me._valueCallbackTemperature(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YTemperatureTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me , True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me , False)
      End If
      Me._timedReportCallbackTemperature = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackTemperature Is Nothing)) Then
        Me._timedReportCallbackTemperature(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of temperature sensors started using <c>yFirstTemperature()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YTemperature</c> object, corresponding to
    '''   a temperature sensor currently online, or a <c>null</c> pointer
    '''   if there are no more temperature sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextTemperature() As YTemperature
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YTemperature.FindTemperature(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of temperature sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YTemperature.nextTemperature()</c> to iterate on
    '''   next temperature sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YTemperature</c> object, corresponding to
    '''   the first temperature sensor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstTemperature() As YTemperature
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Temperature", 0, p, size, neededsize, errmsg)
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
      Return YTemperature.FindTemperature(serial + "." + funcId)
    End Function

    REM --- (end of YTemperature public methods declaration)

  End Class

  REM --- (Temperature functions)

  '''*
  ''' <summary>
  '''   Retrieves a temperature sensor for a given identifier.
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
  '''   This function does not require that the temperature sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YTemperature.isOnline()</c> to test if the temperature sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a temperature sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the temperature sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YTemperature</c> object allowing you to drive the temperature sensor.
  ''' </returns>
  '''/
  Public Function yFindTemperature(ByVal func As String) As YTemperature
    Return YTemperature.FindTemperature(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of temperature sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YTemperature.nextTemperature()</c> to iterate on
  '''   next temperature sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YTemperature</c> object, corresponding to
  '''   the first temperature sensor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstTemperature() As YTemperature
    Return YTemperature.FirstTemperature()
  End Function


  REM --- (end of Temperature functions)

End Module
