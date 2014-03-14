'*********************************************************************
'*
'* $Id: yocto_colorled.vb 15259 2014-03-06 10:21:05Z seb $
'*
'* Implements yFindColorLed(), the high-level API for ColorLed functions
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

Module yocto_colorled

    REM --- (YColorLed return codes)
    REM --- (end of YColorLed return codes)
  REM --- (YColorLed globals)

Public Class YColorLedMove
  Public target As Integer = YAPI.INVALID_INT
  Public ms As Integer = YAPI.INVALID_INT
  Public moving As Integer = YAPI.INVALID_UINT
End Class

  Public Const Y_RGBCOLOR_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_HSLCOLOR_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_RGBCOLORATPOWERON_INVALID As Integer = YAPI.INVALID_UINT
  Public ReadOnly Y_RGBMOVE_INVALID As YColorLedMove = Nothing
  Public ReadOnly Y_HSLMOVE_INVALID As YColorLedMove = Nothing
  Public Delegate Sub YColorLedValueCallback(ByVal func As YColorLed, ByVal value As String)
  Public Delegate Sub YColorLedTimedReportCallback(ByVal func As YColorLed, ByVal measure As YMeasure)
  REM --- (end of YColorLed globals)

  REM --- (YColorLed class start)

  '''*
  ''' <summary>
  '''   Yoctopuce application programming interface
  '''   allows you to drive a color led using RGB coordinates as well as HSL coordinates.
  ''' <para>
  '''   The module performs all conversions form RGB to HSL automatically. It is then
  '''   self-evident to turn on a led with a given hue and to progressively vary its
  '''   saturation or lightness. If needed, you can find more information on the
  '''   difference between RGB and HSL in the section following this one.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YColorLed
    Inherits YFunction
    REM --- (end of YColorLed class start)

    REM --- (YColorLed definitions)
    Public Const RGBCOLOR_INVALID As Integer = YAPI.INVALID_UINT
    Public Const HSLCOLOR_INVALID As Integer = YAPI.INVALID_UINT
    Public ReadOnly RGBMOVE_INVALID As YColorLedMove = Nothing
    Public ReadOnly HSLMOVE_INVALID As YColorLedMove = Nothing
    Public Const RGBCOLORATPOWERON_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YColorLed definitions)

    REM --- (YColorLed attributes declaration)
    Protected _rgbColor As Integer
    Protected _hslColor As Integer
    Protected _rgbMove As YColorLedMove
    Protected _hslMove As YColorLedMove
    Protected _rgbColorAtPowerOn As Integer
    Protected _valueCallbackColorLed As YColorLedValueCallback
    REM --- (end of YColorLed attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "ColorLed"
      REM --- (YColorLed attributes initialization)
      _rgbColor = RGBCOLOR_INVALID
      _hslColor = HSLCOLOR_INVALID
      _rgbMove = New YColorLedMove()
      _hslMove = New YColorLedMove()
      _rgbColorAtPowerOn = RGBCOLORATPOWERON_INVALID
      _valueCallbackColorLed = Nothing
      REM --- (end of YColorLed attributes initialization)
    End Sub

    REM --- (YColorLed private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "rgbColor") Then
        _rgbColor = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "hslColor") Then
        _hslColor = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "rgbMove") Then
        If (member.recordtype = TJSONRECORDTYPE.JSON_STRUCT) Then
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l = 0 To member.membercount - 1
            submemb = member.members(l)
            If (submemb.name = "moving") Then
              _rgbMove.moving = CInt(submemb.ivalue)
            ElseIf (submemb.name = "target") Then
              _rgbMove.target = CInt(submemb.ivalue)
            ElseIf (submemb.name = "ms") Then
              _rgbMove.ms = CInt(submemb.ivalue)
            End If
          Next l
        End If
        Return 1
      End If
      If (member.name = "hslMove") Then
        If (member.recordtype = TJSONRECORDTYPE.JSON_STRUCT) Then
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l = 0 To member.membercount - 1
            submemb = member.members(l)
            If (submemb.name = "moving") Then
              _hslMove.moving = CInt(submemb.ivalue)
            ElseIf (submemb.name = "target") Then
              _hslMove.target = CInt(submemb.ivalue)
            ElseIf (submemb.name = "ms") Then
              _hslMove.ms = CInt(submemb.ivalue)
            End If
          Next l
        End If
        Return 1
      End If
      If (member.name = "rgbColorAtPowerOn") Then
        _rgbColorAtPowerOn = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YColorLed private methods declaration)

    REM --- (YColorLed public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current RGB color of the led.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current RGB color of the led
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RGBCOLOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rgbColor() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RGBCOLOR_INVALID
        End If
      End If
      Return Me._rgbColor
    End Function


    '''*
    ''' <summary>
    '''   Changes the current color of the led, using a RGB color.
    ''' <para>
    '''   Encoding is done as follows: 0xRRGGBB.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current color of the led, using a RGB color
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
    Public Function set_rgbColor(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = "0x" + Hex(newval)
      Return _setAttr("rgbColor", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current HSL color of the led.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current HSL color of the led
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_HSLCOLOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_hslColor() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return HSLCOLOR_INVALID
        End If
      End If
      Return Me._hslColor
    End Function


    '''*
    ''' <summary>
    '''   Changes the current color of the led, using a color HSL.
    ''' <para>
    '''   Encoding is done as follows: 0xHHSSLL.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current color of the led, using a color HSL
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
    Public Function set_hslColor(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = "0x" + Hex(newval)
      Return _setAttr("hslColor", rest_val)
    End Function
    Public Function get_rgbMove() As YColorLedMove
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RGBMOVE_INVALID
        End If
      End If
      Return Me._rgbMove
    End Function


    Public Function set_rgbMove(ByVal newval As YColorLedMove) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target)) + ":" + Ltrim(Str(newval.ms))
      Return _setAttr("rgbMove", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth transition in the RGB color space between the current color and a target color.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="rgb_target">
    '''   desired RGB color at the end of the transition
    ''' </param>
    ''' <param name="ms_duration">
    '''   duration of the transition, in millisecond
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
    Public Function rgbMove(ByVal rgb_target As Integer, ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(rgb_target)) + ":" + Ltrim(Str(ms_duration))
      Return _setAttr("rgbMove", rest_val)
    End Function
    Public Function get_hslMove() As YColorLedMove
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return HSLMOVE_INVALID
        End If
      End If
      Return Me._hslMove
    End Function


    Public Function set_hslMove(ByVal newval As YColorLedMove) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target)) + ":" + Ltrim(Str(newval.ms))
      Return _setAttr("hslMove", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth transition in the HSL color space between the current color and a target color.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="hsl_target">
    '''   desired HSL color at the end of the transition
    ''' </param>
    ''' <param name="ms_duration">
    '''   duration of the transition, in millisecond
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
    Public Function hslMove(ByVal hsl_target As Integer, ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(hsl_target)) + ":" + Ltrim(Str(ms_duration))
      Return _setAttr("hslMove", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the configured color to be displayed when the module is turned on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the configured color to be displayed when the module is turned on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RGBCOLORATPOWERON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rgbColorAtPowerOn() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RGBCOLORATPOWERON_INVALID
        End If
      End If
      Return Me._rgbColorAtPowerOn
    End Function


    '''*
    ''' <summary>
    '''   Changes the color that the led will display by default when the module is turned on.
    ''' <para>
    '''   This color will be displayed as soon as the module is powered on.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   change should be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the color that the led will display by default when the module is turned on
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
    Public Function set_rgbColorAtPowerOn(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = "0x" + Hex(newval)
      Return _setAttr("rgbColorAtPowerOn", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves an RGB led for a given identifier.
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
    '''   This function does not require that the RGB led is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YColorLed.isOnline()</c> to test if the RGB led is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an RGB led by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the RGB led
    ''' </param>
    ''' <returns>
    '''   a <c>YColorLed</c> object allowing you to drive the RGB led.
    ''' </returns>
    '''/
    Public Shared Function FindColorLed(func As String) As YColorLed
      Dim obj As YColorLed
      obj = CType(YFunction._FindFromCache("ColorLed", func), YColorLed)
      If ((obj Is Nothing)) Then
        obj = New YColorLed(func)
        YFunction._AddToCache("ColorLed", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YColorLedValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackColorLed = callback
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
      If (Not (Me._valueCallbackColorLed Is Nothing)) Then
        Me._valueCallbackColorLed(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of RGB leds started using <c>yFirstColorLed()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YColorLed</c> object, corresponding to
    '''   an RGB led currently online, or a <c>null</c> pointer
    '''   if there are no more RGB leds to enumerate.
    ''' </returns>
    '''/
    Public Function nextColorLed() As YColorLed
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YColorLed.FindColorLed(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of RGB leds currently accessible.
    ''' <para>
    '''   Use the method <c>YColorLed.nextColorLed()</c> to iterate on
    '''   next RGB leds.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YColorLed</c> object, corresponding to
    '''   the first RGB led currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstColorLed() As YColorLed
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("ColorLed", 0, p, size, neededsize, errmsg)
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
      Return YColorLed.FindColorLed(serial + "." + funcId)
    End Function

    REM --- (end of YColorLed public methods declaration)

  End Class

  REM --- (ColorLed functions)

  '''*
  ''' <summary>
  '''   Retrieves an RGB led for a given identifier.
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
  '''   This function does not require that the RGB led is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YColorLed.isOnline()</c> to test if the RGB led is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an RGB led by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the RGB led
  ''' </param>
  ''' <returns>
  '''   a <c>YColorLed</c> object allowing you to drive the RGB led.
  ''' </returns>
  '''/
  Public Function yFindColorLed(ByVal func As String) As YColorLed
    Return YColorLed.FindColorLed(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of RGB leds currently accessible.
  ''' <para>
  '''   Use the method <c>YColorLed.nextColorLed()</c> to iterate on
  '''   next RGB leds.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YColorLed</c> object, corresponding to
  '''   the first RGB led currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstColorLed() As YColorLed
    Return YColorLed.FirstColorLed()
  End Function


  REM --- (end of ColorLed functions)

End Module
