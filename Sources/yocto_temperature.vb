'*********************************************************************
'*
'* $Id: yocto_temperature.vb 18322 2014-11-10 10:49:13Z seb $
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
    REM --- (YTemperature dlldef)
    REM --- (end of YTemperature dlldef)
  REM --- (YTemperature globals)

  Public Const Y_SENSORTYPE_DIGITAL As Integer = 0
  Public Const Y_SENSORTYPE_TYPE_K As Integer = 1
  Public Const Y_SENSORTYPE_TYPE_E As Integer = 2
  Public Const Y_SENSORTYPE_TYPE_J As Integer = 3
  Public Const Y_SENSORTYPE_TYPE_N As Integer = 4
  Public Const Y_SENSORTYPE_TYPE_R As Integer = 5
  Public Const Y_SENSORTYPE_TYPE_S As Integer = 6
  Public Const Y_SENSORTYPE_TYPE_T As Integer = 7
  Public Const Y_SENSORTYPE_PT100_4WIRES As Integer = 8
  Public Const Y_SENSORTYPE_PT100_3WIRES As Integer = 9
  Public Const Y_SENSORTYPE_PT100_2WIRES As Integer = 10
  Public Const Y_SENSORTYPE_RES_OHM As Integer = 11
  Public Const Y_SENSORTYPE_RES_NTC As Integer = 12
  Public Const Y_SENSORTYPE_RES_LINEAR As Integer = 13
  Public Const Y_SENSORTYPE_INVALID As Integer = -1
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
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
    Public Const SENSORTYPE_DIGITAL As Integer = 0
    Public Const SENSORTYPE_TYPE_K As Integer = 1
    Public Const SENSORTYPE_TYPE_E As Integer = 2
    Public Const SENSORTYPE_TYPE_J As Integer = 3
    Public Const SENSORTYPE_TYPE_N As Integer = 4
    Public Const SENSORTYPE_TYPE_R As Integer = 5
    Public Const SENSORTYPE_TYPE_S As Integer = 6
    Public Const SENSORTYPE_TYPE_T As Integer = 7
    Public Const SENSORTYPE_PT100_4WIRES As Integer = 8
    Public Const SENSORTYPE_PT100_3WIRES As Integer = 9
    Public Const SENSORTYPE_PT100_2WIRES As Integer = 10
    Public Const SENSORTYPE_RES_OHM As Integer = 11
    Public Const SENSORTYPE_RES_NTC As Integer = 12
    Public Const SENSORTYPE_RES_LINEAR As Integer = 13
    Public Const SENSORTYPE_INVALID As Integer = -1
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YTemperature definitions)

    REM --- (YTemperature attributes declaration)
    Protected _sensorType As Integer
    Protected _command As String
    Protected _valueCallbackTemperature As YTemperatureValueCallback
    Protected _timedReportCallbackTemperature As YTemperatureTimedReportCallback
    REM --- (end of YTemperature attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Temperature"
      REM --- (YTemperature attributes initialization)
      _sensorType = SENSORTYPE_INVALID
      _command = COMMAND_INVALID
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
      If (member.name = "command") Then
        _command = member.svalue
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
    '''   <c>Y_SENSORTYPE_PT100_3WIRES</c>, <c>Y_SENSORTYPE_PT100_2WIRES</c>, <c>Y_SENSORTYPE_RES_OHM</c>,
    '''   <c>Y_SENSORTYPE_RES_NTC</c> and <c>Y_SENSORTYPE_RES_LINEAR</c> corresponding to the temperature sensor type
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
    '''   <c>Y_SENSORTYPE_PT100_3WIRES</c>, <c>Y_SENSORTYPE_PT100_2WIRES</c>, <c>Y_SENSORTYPE_RES_OHM</c>,
    '''   <c>Y_SENSORTYPE_RES_NTC</c> and <c>Y_SENSORTYPE_RES_LINEAR</c>
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
    Public Function get_command() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      Return Me._command
    End Function


    Public Function set_command(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("command", rest_val)
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
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
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
        YFunction._UpdateTimedReportCallbackList(Me, True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me, False)
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
    '''   Record a thermistor response table, for interpolating the temperature from
    '''   the measured resistance.
    ''' <para>
    '''   This function can only be used with temperature
    '''   sensor based on thermistors.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, corresponding to all
    '''   temperatures (in degrees Celcius) for which the resistance of the
    '''   thermistor is specified.
    ''' </param>
    ''' <param name="resValues">
    '''   array of floating point numbers, corresponding to the resistance
    '''   values (in Ohms) for each of the temperature included in the first
    '''   argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_thermistorResponseTable(tempValues As List(Of Double), resValues As List(Of Double)) As Integer
      Dim siz As Integer = 0
      Dim res As Integer = 0
      Dim idx As Integer = 0
      Dim found As Integer = 0
      Dim prev As Double = 0
      Dim curr As Double = 0
      Dim currTemp As Double = 0
      Dim idxres As Double = 0
      siz = tempValues.Count
      If Not(siz >= 2) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "thermistor response table must have at least two points")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not(siz = resValues.Count) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "table sizes mismatch")
        return YAPI.INVALID_ARGUMENT
      end if
      
      REM // may throw an exception
      res = Me.set_command("Z")
      If Not(res=YAPI.SUCCESS) Then
        me._throw( YAPI.IO_ERROR,  "unable to reset thermistor parameters")
        return YAPI.IO_ERROR
      end if
      
      REM // add records in growing resistance value
      found = 1
      prev = 0.0
      While (found > 0)
        found = 0
        curr = 99999999.0
        currTemp = -999999.0
        idx = 0
        While (idx < siz)
          idxres = resValues(idx)
          If ((idxres > prev) And (idxres < curr)) Then
            curr = idxres
            currTemp = tempValues(idx)
            found = 1
          End If
          idx = idx + 1
        End While
        If (found > 0) Then
          res = Me.set_command("m" + Convert.ToString( CType(Math.Round(1000*curr), Integer)) + ":" + Convert.ToString(CType(Math.Round(1000*currTemp), Integer)))
          If Not(res=YAPI.SUCCESS) Then
            me._throw( YAPI.IO_ERROR,  "unable to reset thermistor parameters")
            return YAPI.IO_ERROR
          end if
          prev = curr
        End If
      End While
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the thermistor response table previously configured using function
    '''   <c>set_thermistorResponseTable</c>.
    ''' <para>
    '''   This function can only be used with
    '''   temperature sensor based on thermistors.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, that will be filled by the function
    '''   with all temperatures (in degrees Celcius) for which the resistance
    '''   of the thermistor is specified.
    ''' </param>
    ''' <param name="resValues">
    '''   array of floating point numbers, that will be filled by the function
    '''   with the value (in Ohms) for each of the temperature included in the
    '''   first argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadThermistorResponseTable(tempValues As List(Of Double), resValues As List(Of Double)) As Integer
      Dim id As String
      Dim bin_json As Byte()
      Dim paramlist As List(Of String) = New List(Of String)()
      Dim templist As List(Of Double) = New List(Of Double)()
      Dim siz As Integer = 0
      Dim idx As Integer = 0
      Dim temp As Double = 0
      Dim found As Integer = 0
      Dim prev As Double = 0
      Dim curr As Double = 0
      Dim currRes As Double = 0
      
      tempValues.Clear()
      resValues.Clear()
      
      REM // may throw an exception
      id = Me.get_functionId()
      id = (id).Substring( 11, (id).Length-1)
      bin_json = Me._download("extra.json?page=" + id)
      paramlist = Me._json_get_array(bin_json)
      REM // first convert all temperatures to float
      siz = ((paramlist.Count) >> (1))
      templist.Clear()
      idx = 0
      While (idx < siz)
        temp = Double.Parse(paramlist(2*idx+1))/1000.0
        templist.Add(temp)
        idx = idx + 1
      End While
      REM // then add records in growing temperature value
      tempValues.Clear()
      resValues.Clear()
      found = 1
      prev = -999999.0
      While (found > 0)
        found = 0
        curr = 999999.0
        currRes = -999999.0
        idx = 0
        While (idx < siz)
          temp = templist(idx)
          If ((temp > prev) And (temp < curr)) Then
            curr = temp
            currRes = Double.Parse(paramlist(2*idx))/1000.0
            found = 1
          End If
          idx = idx + 1
        End While
        If (found > 0) Then
          tempValues.Add(curr)
          resValues.Add(currRes)
          prev = curr
        End If
      End While
      
      
      
      Return YAPI.SUCCESS
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
