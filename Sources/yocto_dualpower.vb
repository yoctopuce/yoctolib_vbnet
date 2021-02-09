' ********************************************************************
'
'  $Id: yocto_dualpower.vb 43580 2021-01-26 17:46:01Z mvuilleu $
'
'  Implements yFindDualPower(), the high-level API for DualPower functions
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

Module yocto_dualpower

    REM --- (YDualPower return codes)
    REM --- (end of YDualPower return codes)
    REM --- (YDualPower dlldef)
    REM --- (end of YDualPower dlldef)
   REM --- (YDualPower yapiwrapper)
   REM --- (end of YDualPower yapiwrapper)
  REM --- (YDualPower globals)

  Public Const Y_POWERSTATE_OFF As Integer = 0
  Public Const Y_POWERSTATE_FROM_USB As Integer = 1
  Public Const Y_POWERSTATE_FROM_EXT As Integer = 2
  Public Const Y_POWERSTATE_INVALID As Integer = -1
  Public Const Y_POWERCONTROL_AUTO As Integer = 0
  Public Const Y_POWERCONTROL_FROM_USB As Integer = 1
  Public Const Y_POWERCONTROL_FROM_EXT As Integer = 2
  Public Const Y_POWERCONTROL_OFF As Integer = 3
  Public Const Y_POWERCONTROL_INVALID As Integer = -1
  Public Const Y_EXTVOLTAGE_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YDualPowerValueCallback(ByVal func As YDualPower, ByVal value As String)
  Public Delegate Sub YDualPowerTimedReportCallback(ByVal func As YDualPower, ByVal measure As YMeasure)
  REM --- (end of YDualPower globals)

  REM --- (YDualPower class start)

  '''*
  ''' <summary>
  '''   The <c>YDualPower</c> class allows you to control
  '''   the power source to use for module functions that require high current.
  ''' <para>
  '''   The module can also automatically disconnect the external power
  '''   when a voltage drop is observed on the external power source
  '''   (external battery running out of power).
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDualPower
    Inherits YFunction
    REM --- (end of YDualPower class start)

    REM --- (YDualPower definitions)
    Public Const POWERSTATE_OFF As Integer = 0
    Public Const POWERSTATE_FROM_USB As Integer = 1
    Public Const POWERSTATE_FROM_EXT As Integer = 2
    Public Const POWERSTATE_INVALID As Integer = -1
    Public Const POWERCONTROL_AUTO As Integer = 0
    Public Const POWERCONTROL_FROM_USB As Integer = 1
    Public Const POWERCONTROL_FROM_EXT As Integer = 2
    Public Const POWERCONTROL_OFF As Integer = 3
    Public Const POWERCONTROL_INVALID As Integer = -1
    Public Const EXTVOLTAGE_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YDualPower definitions)

    REM --- (YDualPower attributes declaration)
    Protected _powerState As Integer
    Protected _powerControl As Integer
    Protected _extVoltage As Integer
    Protected _valueCallbackDualPower As YDualPowerValueCallback
    REM --- (end of YDualPower attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "DualPower"
      REM --- (YDualPower attributes initialization)
      _powerState = POWERSTATE_INVALID
      _powerControl = POWERCONTROL_INVALID
      _extVoltage = EXTVOLTAGE_INVALID
      _valueCallbackDualPower = Nothing
      REM --- (end of YDualPower attributes initialization)
    End Sub

    REM --- (YDualPower private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("powerState") Then
        _powerState = CInt(json_val.getLong("powerState"))
      End If
      If json_val.has("powerControl") Then
        _powerControl = CInt(json_val.getLong("powerControl"))
      End If
      If json_val.has("extVoltage") Then
        _extVoltage = CInt(json_val.getLong("extVoltage"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YDualPower private methods declaration)

    REM --- (YDualPower public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current power source for module functions that require lots of current.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YDualPower.POWERSTATE_OFF</c>, <c>YDualPower.POWERSTATE_FROM_USB</c> and
    '''   <c>YDualPower.POWERSTATE_FROM_EXT</c> corresponding to the current power source for module
    '''   functions that require lots of current
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDualPower.POWERSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_powerState() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return POWERSTATE_INVALID
        End If
      End If
      res = Me._powerState
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the selected power source for module functions that require lots of current.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YDualPower.POWERCONTROL_AUTO</c>, <c>YDualPower.POWERCONTROL_FROM_USB</c>,
    '''   <c>YDualPower.POWERCONTROL_FROM_EXT</c> and <c>YDualPower.POWERCONTROL_OFF</c> corresponding to the
    '''   selected power source for module functions that require lots of current
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDualPower.POWERCONTROL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_powerControl() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return POWERCONTROL_INVALID
        End If
      End If
      res = Me._powerControl
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the selected power source for module functions that require lots of current.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YDualPower.POWERCONTROL_AUTO</c>, <c>YDualPower.POWERCONTROL_FROM_USB</c>,
    '''   <c>YDualPower.POWERCONTROL_FROM_EXT</c> and <c>YDualPower.POWERCONTROL_OFF</c> corresponding to the
    '''   selected power source for module functions that require lots of current
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
    Public Function set_powerControl(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("powerControl", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the measured voltage on the external power source, in millivolts.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the measured voltage on the external power source, in millivolts
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDualPower.EXTVOLTAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_extVoltage() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return EXTVOLTAGE_INVALID
        End If
      End If
      res = Me._extVoltage
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a dual power switch for a given identifier.
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
    '''   This function does not require that the dual power switch is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YDualPower.isOnline()</c> to test if the dual power switch is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a dual power switch by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the dual power switch, for instance
    '''   <c>SERVORC1.dualPower</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YDualPower</c> object allowing you to drive the dual power switch.
    ''' </returns>
    '''/
    Public Shared Function FindDualPower(func As String) As YDualPower
      Dim obj As YDualPower
      obj = CType(YFunction._FindFromCache("DualPower", func), YDualPower)
      If ((obj Is Nothing)) Then
        obj = New YDualPower(func)
        YFunction._AddToCache("DualPower", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YDualPowerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackDualPower = callback
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
      If (Not (Me._valueCallbackDualPower Is Nothing)) Then
        Me._valueCallbackDualPower(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of dual power switches started using <c>yFirstDualPower()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned dual power switches order.
    '''   If you want to find a specific a dual power switch, use <c>DualPower.findDualPower()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDualPower</c> object, corresponding to
    '''   a dual power switch currently online, or a <c>Nothing</c> pointer
    '''   if there are no more dual power switches to enumerate.
    ''' </returns>
    '''/
    Public Function nextDualPower() As YDualPower
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YDualPower.FindDualPower(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of dual power switches currently accessible.
    ''' <para>
    '''   Use the method <c>YDualPower.nextDualPower()</c> to iterate on
    '''   next dual power switches.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDualPower</c> object, corresponding to
    '''   the first dual power switch currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstDualPower() As YDualPower
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("DualPower", 0, p, size, neededsize, errmsg)
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
      Return YDualPower.FindDualPower(serial + "." + funcId)
    End Function

    REM --- (end of YDualPower public methods declaration)

  End Class

  REM --- (YDualPower functions)

  '''*
  ''' <summary>
  '''   Retrieves a dual power switch for a given identifier.
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
  '''   This function does not require that the dual power switch is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YDualPower.isOnline()</c> to test if the dual power switch is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a dual power switch by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the dual power switch, for instance
  '''   <c>SERVORC1.dualPower</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YDualPower</c> object allowing you to drive the dual power switch.
  ''' </returns>
  '''/
  Public Function yFindDualPower(ByVal func As String) As YDualPower
    Return YDualPower.FindDualPower(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of dual power switches currently accessible.
  ''' <para>
  '''   Use the method <c>YDualPower.nextDualPower()</c> to iterate on
  '''   next dual power switches.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YDualPower</c> object, corresponding to
  '''   the first dual power switch currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstDualPower() As YDualPower
    Return YDualPower.FirstDualPower()
  End Function


  REM --- (end of YDualPower functions)

End Module
