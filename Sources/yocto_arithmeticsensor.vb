' ********************************************************************
'
'  $Id: yocto_arithmeticsensor.vb 55979 2023-08-11 08:24:13Z seb $
'
'  Implements yFindArithmeticSensor(), the high-level API for ArithmeticSensor functions
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

Module yocto_arithmeticsensor

    REM --- (YArithmeticSensor return codes)
    REM --- (end of YArithmeticSensor return codes)
    REM --- (YArithmeticSensor dlldef)
    REM --- (end of YArithmeticSensor dlldef)
   REM --- (YArithmeticSensor yapiwrapper)
   REM --- (end of YArithmeticSensor yapiwrapper)
  REM --- (YArithmeticSensor globals)

  Public Const Y_DESCRIPTION_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YArithmeticSensorValueCallback(ByVal func As YArithmeticSensor, ByVal value As String)
  Public Delegate Sub YArithmeticSensorTimedReportCallback(ByVal func As YArithmeticSensor, ByVal measure As YMeasure)
  REM --- (end of YArithmeticSensor globals)

  REM --- (YArithmeticSensor class start)

  '''*
  ''' <summary>
  '''   The <c>YArithmeticSensor</c> class allows some Yoctopuce devices to compute in real-time
  '''   values based on an arithmetic formula involving one or more measured signals as
  '''   well as the temperature.
  ''' <para>
  '''   As for any physical sensor, the computed values can be
  '''   read by callback and stored in the built-in datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YArithmeticSensor
    Inherits YSensor
    REM --- (end of YArithmeticSensor class start)

    REM --- (YArithmeticSensor definitions)
    Public Const DESCRIPTION_INVALID As String = YAPI.INVALID_STRING
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YArithmeticSensor definitions)

    REM --- (YArithmeticSensor attributes declaration)
    Protected _description As String
    Protected _command As String
    Protected _valueCallbackArithmeticSensor As YArithmeticSensorValueCallback
    Protected _timedReportCallbackArithmeticSensor As YArithmeticSensorTimedReportCallback
    REM --- (end of YArithmeticSensor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "ArithmeticSensor"
      REM --- (YArithmeticSensor attributes initialization)
      _description = DESCRIPTION_INVALID
      _command = COMMAND_INVALID
      _valueCallbackArithmeticSensor = Nothing
      _timedReportCallbackArithmeticSensor = Nothing
      REM --- (end of YArithmeticSensor attributes initialization)
    End Sub

    REM --- (YArithmeticSensor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("description") Then
        _description = json_val.getString("description")
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YArithmeticSensor private methods declaration)

    REM --- (YArithmeticSensor public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the measuring unit for the arithmetic sensor.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the measuring unit for the arithmetic sensor
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_unit(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("unit", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns a short informative description of the formula.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to a short informative description of the formula
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YArithmeticSensor.DESCRIPTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_description() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DESCRIPTION_INVALID
        End If
      End If
      res = Me._description
      Return res
    End Function

    Public Function get_command() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      res = Me._command
      Return res
    End Function


    Public Function set_command(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("command", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves an arithmetic sensor for a given identifier.
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
    '''   This function does not require that the arithmetic sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YArithmeticSensor.isOnline()</c> to test if the arithmetic sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an arithmetic sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the arithmetic sensor, for instance
    '''   <c>RXUVOLT1.arithmeticSensor1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YArithmeticSensor</c> object allowing you to drive the arithmetic sensor.
    ''' </returns>
    '''/
    Public Shared Function FindArithmeticSensor(func As String) As YArithmeticSensor
      Dim obj As YArithmeticSensor
      obj = CType(YFunction._FindFromCache("ArithmeticSensor", func), YArithmeticSensor)
      If ((obj Is Nothing)) Then
        obj = New YArithmeticSensor(func)
        YFunction._AddToCache("ArithmeticSensor", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YArithmeticSensorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackArithmeticSensor = callback
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
      If (Not (Me._valueCallbackArithmeticSensor Is Nothing)) Then
        Me._valueCallbackArithmeticSensor(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YArithmeticSensorTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackArithmeticSensor = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackArithmeticSensor Is Nothing)) Then
        Me._timedReportCallbackArithmeticSensor(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Defines the arithmetic function by means of an algebraic expression.
    ''' <para>
    '''   The expression
    '''   may include references to device sensors, by their physical or logical name, to
    '''   usual math functions and to auxiliary functions defined separately.
    ''' </para>
    ''' </summary>
    ''' <param name="expr">
    '''   the algebraic expression defining the function.
    ''' </param>
    ''' <param name="descr">
    '''   short informative description of the expression.
    ''' </param>
    ''' <returns>
    '''   the current expression value if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns YAPI.INVALID_DOUBLE.
    ''' </para>
    '''/
    Public Overridable Function defineExpression(expr As String, descr As String) As Double
      Dim id As String
      Dim fname As String
      Dim content As String
      Dim data As Byte() = New Byte(){}
      Dim diags As String
      Dim resval As Double = 0
      id = Me.get_functionId()
      id = (id).Substring( 16, (id).Length - 16)
      fname = "arithmExpr" + id + ".txt"

      content = "// " +  descr + "" + vbLf + "" + expr
      data = Me._uploadEx(fname, YAPI.DefaultEncoding.GetBytes(content))
      diags = YAPI.DefaultEncoding.GetString(data)
      If Not((diags).Substring(0, 8) = "Result: ") Then
        me._throw( YAPI.INVALID_ARGUMENT,  diags)
        return YAPI.INVALID_DOUBLE
      end if
      resval = YAPI._atof((diags).Substring( 8, (diags).Length-8))
      Return resval
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the algebraic expression defining the arithmetic function, as previously
    '''   configured using the <c>defineExpression</c> function.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string containing the mathematical expression.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadExpression() As String
      Dim id As String
      Dim fname As String
      Dim content As String
      Dim idx As Integer = 0
      id = Me.get_functionId()
      id = (id).Substring( 16, (id).Length - 16)
      fname = "arithmExpr" + id + ".txt"

      content = YAPI.DefaultEncoding.GetString(Me._download(fname))
      idx = content.IndexOf("" + vbLf + "")
      If (idx > 0) Then
        content = (content).Substring( idx+1, (content).Length-(idx+1))
      End If
      Return content
    End Function

    '''*
    ''' <summary>
    '''   Defines a auxiliary function by means of a table of reference points.
    ''' <para>
    '''   Intermediate values
    '''   will be interpolated between specified reference points. The reference points are given
    '''   as pairs of floating point numbers.
    '''   The auxiliary function will be available for use by all ArithmeticSensor objects of the
    '''   device. Up to nine auxiliary function can be defined in a device, each containing up to
    '''   96 reference points.
    ''' </para>
    ''' </summary>
    ''' <param name="name">
    '''   auxiliary function name, up to 16 characters.
    ''' </param>
    ''' <param name="inputValues">
    '''   array of floating point numbers, corresponding to the function input value.
    ''' </param>
    ''' <param name="outputValues">
    '''   array of floating point numbers, corresponding to the output value
    '''   desired for each of the input value, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function defineAuxiliaryFunction(name As String, inputValues As List(Of Double), outputValues As List(Of Double)) As Integer
      Dim siz As Integer = 0
      Dim defstr As String
      Dim idx As Integer = 0
      Dim inputVal As Double = 0
      Dim outputVal As Double = 0
      Dim fname As String
      siz = inputValues.Count
      If Not(siz > 1) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "auxiliary function must be defined by at least two points")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not(siz = outputValues.Count) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "table sizes mismatch")
        return YAPI.INVALID_ARGUMENT
      end if
      defstr = ""
      idx = 0
      While (idx < siz)
        inputVal = inputValues(idx)
        outputVal = outputValues(idx)
        defstr = "" +  defstr + "" + YAPI._floatToStr( inputVal) + ":" + YAPI._floatToStr(outputVal) + "" + vbLf + ""
        idx = idx + 1
      End While
      fname = "userMap" + name + ".txt"

      Return Me._upload(fname, YAPI.DefaultEncoding.GetBytes(defstr))
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the reference points table defining an auxiliary function previously
    '''   configured using the <c>defineAuxiliaryFunction</c> function.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="name">
    '''   auxiliary function name, up to 16 characters.
    ''' </param>
    ''' <param name="inputValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with all the function reference input value.
    ''' </param>
    ''' <param name="outputValues">
    '''   array of floating point numbers, that is filled by the function
    '''   output value for each of the input value, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadAuxiliaryFunction(name As String, inputValues As List(Of Double), outputValues As List(Of Double)) As Integer
      Dim fname As String
      Dim defbin As Byte() = New Byte(){}
      Dim siz As Integer = 0

      fname = "userMap" + name + ".txt"
      defbin = Me._download(fname)
      siz = (defbin).Length
      If Not(siz > 0) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "auxiliary function does not exist")
        return YAPI.INVALID_ARGUMENT
      end if
      inputValues.Clear()
      outputValues.Clear()
      REM // FIXME: decode line by line
      Return YAPI.SUCCESS
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of arithmetic sensors started using <c>yFirstArithmeticSensor()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned arithmetic sensors order.
    '''   If you want to find a specific an arithmetic sensor, use <c>ArithmeticSensor.findArithmeticSensor()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YArithmeticSensor</c> object, corresponding to
    '''   an arithmetic sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more arithmetic sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextArithmeticSensor() As YArithmeticSensor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YArithmeticSensor.FindArithmeticSensor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of arithmetic sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YArithmeticSensor.nextArithmeticSensor()</c> to iterate on
    '''   next arithmetic sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YArithmeticSensor</c> object, corresponding to
    '''   the first arithmetic sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstArithmeticSensor() As YArithmeticSensor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("ArithmeticSensor", 0, p, size, neededsize, errmsg)
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
      Return YArithmeticSensor.FindArithmeticSensor(serial + "." + funcId)
    End Function

    REM --- (end of YArithmeticSensor public methods declaration)

  End Class

  REM --- (YArithmeticSensor functions)

  '''*
  ''' <summary>
  '''   Retrieves an arithmetic sensor for a given identifier.
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
  '''   This function does not require that the arithmetic sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YArithmeticSensor.isOnline()</c> to test if the arithmetic sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an arithmetic sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the arithmetic sensor, for instance
  '''   <c>RXUVOLT1.arithmeticSensor1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YArithmeticSensor</c> object allowing you to drive the arithmetic sensor.
  ''' </returns>
  '''/
  Public Function yFindArithmeticSensor(ByVal func As String) As YArithmeticSensor
    Return YArithmeticSensor.FindArithmeticSensor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of arithmetic sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YArithmeticSensor.nextArithmeticSensor()</c> to iterate on
  '''   next arithmetic sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YArithmeticSensor</c> object, corresponding to
  '''   the first arithmetic sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstArithmeticSensor() As YArithmeticSensor
    Return YArithmeticSensor.FirstArithmeticSensor()
  End Function


  REM --- (end of YArithmeticSensor functions)

End Module
