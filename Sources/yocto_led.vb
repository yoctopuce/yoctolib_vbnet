' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindLed(), the high-level API for Led functions
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

Module yocto_led

    REM --- (YLed return codes)
    REM --- (end of YLed return codes)
    REM --- (YLed dlldef)
    REM --- (end of YLed dlldef)
   REM --- (YLed yapiwrapper)
   REM --- (end of YLed yapiwrapper)
  REM --- (YLed globals)

  Public Const Y_POWER_OFF As Integer = 0
  Public Const Y_POWER_ON As Integer = 1
  Public Const Y_POWER_INVALID As Integer = -1
  Public Const Y_LUMINOSITY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BLINKING_STILL As Integer = 0
  Public Const Y_BLINKING_RELAX As Integer = 1
  Public Const Y_BLINKING_AWARE As Integer = 2
  Public Const Y_BLINKING_RUN As Integer = 3
  Public Const Y_BLINKING_CALL As Integer = 4
  Public Const Y_BLINKING_PANIC As Integer = 5
  Public Const Y_BLINKING_INVALID As Integer = -1
  Public Delegate Sub YLedValueCallback(ByVal func As YLed, ByVal value As String)
  Public Delegate Sub YLedTimedReportCallback(ByVal func As YLed, ByVal measure As YMeasure)
  REM --- (end of YLed globals)

  REM --- (YLed class start)

  '''*
  ''' <summary>
  '''   The <c>YLed</c> class allows you to drive a monocolor LED.
  ''' <para>
  '''   You can not only to drive the intensity of the LED, but also to
  '''   have it blink at various preset frequencies.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YLed
    Inherits YFunction
    REM --- (end of YLed class start)

    REM --- (YLed definitions)
    Public Const POWER_OFF As Integer = 0
    Public Const POWER_ON As Integer = 1
    Public Const POWER_INVALID As Integer = -1
    Public Const LUMINOSITY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BLINKING_STILL As Integer = 0
    Public Const BLINKING_RELAX As Integer = 1
    Public Const BLINKING_AWARE As Integer = 2
    Public Const BLINKING_RUN As Integer = 3
    Public Const BLINKING_CALL As Integer = 4
    Public Const BLINKING_PANIC As Integer = 5
    Public Const BLINKING_INVALID As Integer = -1
    REM --- (end of YLed definitions)

    REM --- (YLed attributes declaration)
    Protected _power As Integer
    Protected _luminosity As Integer
    Protected _blinking As Integer
    Protected _valueCallbackLed As YLedValueCallback
    REM --- (end of YLed attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Led"
      REM --- (YLed attributes initialization)
      _power = POWER_INVALID
      _luminosity = LUMINOSITY_INVALID
      _blinking = BLINKING_INVALID
      _valueCallbackLed = Nothing
      REM --- (end of YLed attributes initialization)
    End Sub

    REM --- (YLed private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("power") Then
        If (json_val.getInt("power") > 0) Then _power = 1 Else _power = 0
      End If
      If json_val.has("luminosity") Then
        _luminosity = CInt(json_val.getLong("luminosity"))
      End If
      If json_val.has("blinking") Then
        _blinking = CInt(json_val.getLong("blinking"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YLed private methods declaration)

    REM --- (YLed public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current LED state.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YLed.POWER_OFF</c> or <c>YLed.POWER_ON</c>, according to the current LED state
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YLed.POWER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_power() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return POWER_INVALID
        End If
      End If
      res = Me._power
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the state of the LED.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YLed.POWER_OFF</c> or <c>YLed.POWER_ON</c>, according to the state of the LED
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
    Public Function set_power(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("power", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current LED intensity (in per cent).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current LED intensity (in per cent)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YLed.LUMINOSITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_luminosity() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LUMINOSITY_INVALID
        End If
      End If
      res = Me._luminosity
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current LED intensity (in per cent).
    ''' <para>
    '''   Remember to call the
    '''   <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current LED intensity (in per cent)
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
    Public Function set_luminosity(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("luminosity", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current LED signaling mode.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YLed.BLINKING_STILL</c>, <c>YLed.BLINKING_RELAX</c>, <c>YLed.BLINKING_AWARE</c>,
    '''   <c>YLed.BLINKING_RUN</c>, <c>YLed.BLINKING_CALL</c> and <c>YLed.BLINKING_PANIC</c> corresponding to
    '''   the current LED signaling mode
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YLed.BLINKING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_blinking() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BLINKING_INVALID
        End If
      End If
      res = Me._blinking
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current LED signaling mode.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YLed.BLINKING_STILL</c>, <c>YLed.BLINKING_RELAX</c>, <c>YLed.BLINKING_AWARE</c>,
    '''   <c>YLed.BLINKING_RUN</c>, <c>YLed.BLINKING_CALL</c> and <c>YLed.BLINKING_PANIC</c> corresponding to
    '''   the current LED signaling mode
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
    Public Function set_blinking(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("blinking", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a monochrome LED for a given identifier.
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
    '''   This function does not require that the monochrome LED is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YLed.isOnline()</c> to test if the monochrome LED is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a monochrome LED by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the monochrome LED, for instance
    '''   <c>YBUZZER2.led1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YLed</c> object allowing you to drive the monochrome LED.
    ''' </returns>
    '''/
    Public Shared Function FindLed(func As String) As YLed
      Dim obj As YLed
      obj = CType(YFunction._FindFromCache("Led", func), YLed)
      If ((obj Is Nothing)) Then
        obj = New YLed(func)
        YFunction._AddToCache("Led", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YLedValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackLed = callback
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
      If (Not (Me._valueCallbackLed Is Nothing)) Then
        Me._valueCallbackLed(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of monochrome LEDs started using <c>yFirstLed()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned monochrome LEDs order.
    '''   If you want to find a specific a monochrome LED, use <c>Led.findLed()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLed</c> object, corresponding to
    '''   a monochrome LED currently online, or a <c>Nothing</c> pointer
    '''   if there are no more monochrome LEDs to enumerate.
    ''' </returns>
    '''/
    Public Function nextLed() As YLed
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YLed.FindLed(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of monochrome LEDs currently accessible.
    ''' <para>
    '''   Use the method <c>YLed.nextLed()</c> to iterate on
    '''   next monochrome LEDs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLed</c> object, corresponding to
    '''   the first monochrome LED currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstLed() As YLed
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Led", 0, p, size, neededsize, errmsg)
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
      Return YLed.FindLed(serial + "." + funcId)
    End Function

    REM --- (end of YLed public methods declaration)

  End Class

  REM --- (YLed functions)

  '''*
  ''' <summary>
  '''   Retrieves a monochrome LED for a given identifier.
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
  '''   This function does not require that the monochrome LED is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YLed.isOnline()</c> to test if the monochrome LED is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a monochrome LED by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the monochrome LED, for instance
  '''   <c>YBUZZER2.led1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YLed</c> object allowing you to drive the monochrome LED.
  ''' </returns>
  '''/
  Public Function yFindLed(ByVal func As String) As YLed
    Return YLed.FindLed(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of monochrome LEDs currently accessible.
  ''' <para>
  '''   Use the method <c>YLed.nextLed()</c> to iterate on
  '''   next monochrome LEDs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YLed</c> object, corresponding to
  '''   the first monochrome LED currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstLed() As YLed
    Return YLed.FirstLed()
  End Function


  REM --- (end of YLed functions)

End Module
