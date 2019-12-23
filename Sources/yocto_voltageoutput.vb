' ********************************************************************
'
'  $Id: yocto_voltageoutput.vb 38899 2019-12-20 17:21:03Z mvuilleu $
'
'  Implements yFindVoltageOutput(), the high-level API for VoltageOutput functions
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

Module yocto_voltageoutput

    REM --- (YVoltageOutput return codes)
    REM --- (end of YVoltageOutput return codes)
    REM --- (YVoltageOutput dlldef)
    REM --- (end of YVoltageOutput dlldef)
   REM --- (YVoltageOutput yapiwrapper)
   REM --- (end of YVoltageOutput yapiwrapper)
  REM --- (YVoltageOutput globals)

  Public Const Y_CURRENTVOLTAGE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_VOLTAGETRANSITION_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_VOLTAGEATSTARTUP_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YVoltageOutputValueCallback(ByVal func As YVoltageOutput, ByVal value As String)
  Public Delegate Sub YVoltageOutputTimedReportCallback(ByVal func As YVoltageOutput, ByVal measure As YMeasure)
  REM --- (end of YVoltageOutput globals)

  REM --- (YVoltageOutput class start)

  '''*
  ''' <summary>
  '''   The <c>YVoltageOutput</c> class allows you to drive a voltage output.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YVoltageOutput
    Inherits YFunction
    REM --- (end of YVoltageOutput class start)

    REM --- (YVoltageOutput definitions)
    Public Const CURRENTVOLTAGE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const VOLTAGETRANSITION_INVALID As String = YAPI.INVALID_STRING
    Public Const VOLTAGEATSTARTUP_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YVoltageOutput definitions)

    REM --- (YVoltageOutput attributes declaration)
    Protected _currentVoltage As Double
    Protected _voltageTransition As String
    Protected _voltageAtStartUp As Double
    Protected _valueCallbackVoltageOutput As YVoltageOutputValueCallback
    REM --- (end of YVoltageOutput attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "VoltageOutput"
      REM --- (YVoltageOutput attributes initialization)
      _currentVoltage = CURRENTVOLTAGE_INVALID
      _voltageTransition = VOLTAGETRANSITION_INVALID
      _voltageAtStartUp = VOLTAGEATSTARTUP_INVALID
      _valueCallbackVoltageOutput = Nothing
      REM --- (end of YVoltageOutput attributes initialization)
    End Sub

    REM --- (YVoltageOutput private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("currentVoltage") Then
        _currentVoltage = Math.Round(json_val.getDouble("currentVoltage") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("voltageTransition") Then
        _voltageTransition = json_val.getString("voltageTransition")
      End If
      If json_val.has("voltageAtStartUp") Then
        _voltageAtStartUp = Math.Round(json_val.getDouble("voltageAtStartUp") * 1000.0 / 65536.0) / 1000.0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YVoltageOutput private methods declaration)

    REM --- (YVoltageOutput public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the output voltage, in V.
    ''' <para>
    '''   Valid range is from 0 to 10V.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the output voltage, in V
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
    Public Function set_currentVoltage(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("currentVoltage", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the output voltage set point, in V.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the output voltage set point, in V
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENTVOLTAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentVoltage() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CURRENTVOLTAGE_INVALID
        End If
      End If
      res = Me._currentVoltage
      Return res
    End Function

    Public Function get_voltageTransition() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return VOLTAGETRANSITION_INVALID
        End If
      End If
      res = Me._voltageTransition
      Return res
    End Function


    Public Function set_voltageTransition(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("voltageTransition", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the output voltage at device start up.
    ''' <para>
    '''   Remember to call the matching
    '''   module <c>saveToFlash()</c> method, otherwise this call has no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the output voltage at device start up
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
    Public Function set_voltageAtStartUp(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("voltageAtStartUp", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the selected voltage output at device startup, in V.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the selected voltage output at device startup, in V
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VOLTAGEATSTARTUP_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_voltageAtStartUp() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return VOLTAGEATSTARTUP_INVALID
        End If
      End If
      res = Me._voltageAtStartUp
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a voltage output for a given identifier.
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
    '''   This function does not require that the voltage output is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YVoltageOutput.isOnline()</c> to test if the voltage output is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a voltage output by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the voltage output, for instance
    '''   <c>TX010V01.voltageOutput1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YVoltageOutput</c> object allowing you to drive the voltage output.
    ''' </returns>
    '''/
    Public Shared Function FindVoltageOutput(func As String) As YVoltageOutput
      Dim obj As YVoltageOutput
      obj = CType(YFunction._FindFromCache("VoltageOutput", func), YVoltageOutput)
      If ((obj Is Nothing)) Then
        obj = New YVoltageOutput(func)
        YFunction._AddToCache("VoltageOutput", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YVoltageOutputValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackVoltageOutput = callback
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
      If (Not (Me._valueCallbackVoltageOutput Is Nothing)) Then
        Me._valueCallbackVoltageOutput(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth transition of output voltage.
    ''' <para>
    '''   Any explicit voltage
    '''   change cancels any ongoing transition process.
    ''' </para>
    ''' </summary>
    ''' <param name="V_target">
    '''   new output voltage value at the end of the transition
    '''   (floating-point number, representing the end voltage in V)
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the transition, in milliseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    '''/
    Public Overridable Function voltageMove(V_target As Double, ms_duration As Integer) As Integer
      Dim newval As String
      If (V_target < 0.0) Then
        V_target  = 0.0
      End If
      If (V_target > 10.0) Then
        V_target = 10.0
      End If
      newval = "" + Convert.ToString( CType(Math.Round(V_target*65536), Integer)) + ":" + Convert.ToString(ms_duration)

      Return Me.set_voltageTransition(newval)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of voltage outputs started using <c>yFirstVoltageOutput()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned voltage outputs order.
    '''   If you want to find a specific a voltage output, use <c>VoltageOutput.findVoltageOutput()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVoltageOutput</c> object, corresponding to
    '''   a voltage output currently online, or a <c>Nothing</c> pointer
    '''   if there are no more voltage outputs to enumerate.
    ''' </returns>
    '''/
    Public Function nextVoltageOutput() As YVoltageOutput
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YVoltageOutput.FindVoltageOutput(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of voltage outputs currently accessible.
    ''' <para>
    '''   Use the method <c>YVoltageOutput.nextVoltageOutput()</c> to iterate on
    '''   next voltage outputs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVoltageOutput</c> object, corresponding to
    '''   the first voltage output currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstVoltageOutput() As YVoltageOutput
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("VoltageOutput", 0, p, size, neededsize, errmsg)
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
      Return YVoltageOutput.FindVoltageOutput(serial + "." + funcId)
    End Function

    REM --- (end of YVoltageOutput public methods declaration)

  End Class

  REM --- (YVoltageOutput functions)

  '''*
  ''' <summary>
  '''   Retrieves a voltage output for a given identifier.
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
  '''   This function does not require that the voltage output is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YVoltageOutput.isOnline()</c> to test if the voltage output is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a voltage output by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the voltage output, for instance
  '''   <c>TX010V01.voltageOutput1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YVoltageOutput</c> object allowing you to drive the voltage output.
  ''' </returns>
  '''/
  Public Function yFindVoltageOutput(ByVal func As String) As YVoltageOutput
    Return YVoltageOutput.FindVoltageOutput(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of voltage outputs currently accessible.
  ''' <para>
  '''   Use the method <c>YVoltageOutput.nextVoltageOutput()</c> to iterate on
  '''   next voltage outputs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YVoltageOutput</c> object, corresponding to
  '''   the first voltage output currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstVoltageOutput() As YVoltageOutput
    Return YVoltageOutput.FirstVoltageOutput()
  End Function


  REM --- (end of YVoltageOutput functions)

End Module
