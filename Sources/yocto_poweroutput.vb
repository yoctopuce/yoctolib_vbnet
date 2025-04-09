' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindPowerOutput(), the high-level API for PowerOutput functions
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

Module yocto_poweroutput

    REM --- (YPowerOutput return codes)
    REM --- (end of YPowerOutput return codes)
    REM --- (YPowerOutput dlldef)
    REM --- (end of YPowerOutput dlldef)
   REM --- (YPowerOutput yapiwrapper)
   REM --- (end of YPowerOutput yapiwrapper)
  REM --- (YPowerOutput globals)

  Public Const Y_VOLTAGE_OFF As Integer = 0
  Public Const Y_VOLTAGE_OUT3V3 As Integer = 1
  Public Const Y_VOLTAGE_OUT5V As Integer = 2
  Public Const Y_VOLTAGE_OUT4V7 As Integer = 3
  Public Const Y_VOLTAGE_OUT1V8 As Integer = 4
  Public Const Y_VOLTAGE_INVALID As Integer = -1
  Public Delegate Sub YPowerOutputValueCallback(ByVal func As YPowerOutput, ByVal value As String)
  Public Delegate Sub YPowerOutputTimedReportCallback(ByVal func As YPowerOutput, ByVal measure As YMeasure)
  REM --- (end of YPowerOutput globals)

  REM --- (YPowerOutput class start)

  '''*
  ''' <summary>
  '''   The <c>YPowerOutput</c> class allows you to control
  '''   the power output featured on some Yoctopuce devices.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YPowerOutput
    Inherits YFunction
    REM --- (end of YPowerOutput class start)

    REM --- (YPowerOutput definitions)
    Public Const VOLTAGE_OFF As Integer = 0
    Public Const VOLTAGE_OUT3V3 As Integer = 1
    Public Const VOLTAGE_OUT5V As Integer = 2
    Public Const VOLTAGE_OUT4V7 As Integer = 3
    Public Const VOLTAGE_OUT1V8 As Integer = 4
    Public Const VOLTAGE_INVALID As Integer = -1
    REM --- (end of YPowerOutput definitions)

    REM --- (YPowerOutput attributes declaration)
    Protected _voltage As Integer
    Protected _valueCallbackPowerOutput As YPowerOutputValueCallback
    REM --- (end of YPowerOutput attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "PowerOutput"
      REM --- (YPowerOutput attributes initialization)
      _voltage = VOLTAGE_INVALID
      _valueCallbackPowerOutput = Nothing
      REM --- (end of YPowerOutput attributes initialization)
    End Sub

    REM --- (YPowerOutput private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("voltage") Then
        _voltage = CInt(json_val.getLong("voltage"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YPowerOutput private methods declaration)

    REM --- (YPowerOutput public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the voltage on the power output featured by the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YPowerOutput.VOLTAGE_OFF</c>, <c>YPowerOutput.VOLTAGE_OUT3V3</c>,
    '''   <c>YPowerOutput.VOLTAGE_OUT5V</c>, <c>YPowerOutput.VOLTAGE_OUT4V7</c> and
    '''   <c>YPowerOutput.VOLTAGE_OUT1V8</c> corresponding to the voltage on the power output featured by the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YPowerOutput.VOLTAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_voltage() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return VOLTAGE_INVALID
        End If
      End If
      res = Me._voltage
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the voltage on the power output provided by the
    '''   module.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YPowerOutput.VOLTAGE_OFF</c>, <c>YPowerOutput.VOLTAGE_OUT3V3</c>,
    '''   <c>YPowerOutput.VOLTAGE_OUT5V</c>, <c>YPowerOutput.VOLTAGE_OUT4V7</c> and
    '''   <c>YPowerOutput.VOLTAGE_OUT1V8</c> corresponding to the voltage on the power output provided by the
    '''   module
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
    Public Function set_voltage(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("voltage", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a power output for a given identifier.
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
    '''   This function does not require that the power output is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YPowerOutput.isOnline()</c> to test if the power output is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a power output by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the power output, for instance
    '''   <c>YI2CMK01.powerOutput</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YPowerOutput</c> object allowing you to drive the power output.
    ''' </returns>
    '''/
    Public Shared Function FindPowerOutput(func As String) As YPowerOutput
      Dim obj As YPowerOutput
      obj = CType(YFunction._FindFromCache("PowerOutput", func), YPowerOutput)
      If ((obj Is Nothing)) Then
        obj = New YPowerOutput(func)
        YFunction._AddToCache("PowerOutput", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YPowerOutputValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackPowerOutput = callback
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
      If (Not (Me._valueCallbackPowerOutput Is Nothing)) Then
        Me._valueCallbackPowerOutput(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of power output started using <c>yFirstPowerOutput()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned power output order.
    '''   If you want to find a specific a power output, use <c>PowerOutput.findPowerOutput()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPowerOutput</c> object, corresponding to
    '''   a power output currently online, or a <c>Nothing</c> pointer
    '''   if there are no more power output to enumerate.
    ''' </returns>
    '''/
    Public Function nextPowerOutput() As YPowerOutput
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YPowerOutput.FindPowerOutput(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of power output currently accessible.
    ''' <para>
    '''   Use the method <c>YPowerOutput.nextPowerOutput()</c> to iterate on
    '''   next power output.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPowerOutput</c> object, corresponding to
    '''   the first power output currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstPowerOutput() As YPowerOutput
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("PowerOutput", 0, p, size, neededsize, errmsg)
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
      Return YPowerOutput.FindPowerOutput(serial + "." + funcId)
    End Function

    REM --- (end of YPowerOutput public methods declaration)

  End Class

  REM --- (YPowerOutput functions)

  '''*
  ''' <summary>
  '''   Retrieves a power output for a given identifier.
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
  '''   This function does not require that the power output is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YPowerOutput.isOnline()</c> to test if the power output is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a power output by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the power output, for instance
  '''   <c>YI2CMK01.powerOutput</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YPowerOutput</c> object allowing you to drive the power output.
  ''' </returns>
  '''/
  Public Function yFindPowerOutput(ByVal func As String) As YPowerOutput
    Return YPowerOutput.FindPowerOutput(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of power output currently accessible.
  ''' <para>
  '''   Use the method <c>YPowerOutput.nextPowerOutput()</c> to iterate on
  '''   next power output.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YPowerOutput</c> object, corresponding to
  '''   the first power output currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstPowerOutput() As YPowerOutput
    Return YPowerOutput.FirstPowerOutput()
  End Function


  REM --- (end of YPowerOutput functions)

End Module
