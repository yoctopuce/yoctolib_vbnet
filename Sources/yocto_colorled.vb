'*********************************************************************
'*
'* $Id: yocto_colorled.vb 25275 2016-08-24 13:42:24Z mvuilleu $
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
    REM --- (YColorLed dlldef)
    REM --- (end of YColorLed dlldef)
  REM --- (YColorLed globals)

Public Class YColorLedMove
  Public target As Integer = YAPI.INVALID_INT
  Public ms As Integer = YAPI.INVALID_INT
  Public moving As Integer = YAPI.INVALID_UINT
End Class

  Public Const Y_RGBCOLOR_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_HSLCOLOR_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_RGBCOLORATPOWERON_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BLINKSEQSIZE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BLINKSEQMAXSIZE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BLINKSEQSIGNATURE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public ReadOnly Y_RGBMOVE_INVALID As YColorLedMove = Nothing
  Public ReadOnly Y_HSLMOVE_INVALID As YColorLedMove = Nothing
  Public Delegate Sub YColorLedValueCallback(ByVal func As YColorLed, ByVal value As String)
  Public Delegate Sub YColorLedTimedReportCallback(ByVal func As YColorLed, ByVal measure As YMeasure)
  REM --- (end of YColorLed globals)

  REM --- (YColorLed class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface
  '''   allows you to drive a color LED using RGB coordinates as well as HSL coordinates.
  ''' <para>
  '''   The module performs all conversions form RGB to HSL automatically. It is then
  '''   self-evident to turn on a LED with a given hue and to progressively vary its
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
    Public Const BLINKSEQSIZE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BLINKSEQMAXSIZE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BLINKSEQSIGNATURE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YColorLed definitions)

    REM --- (YColorLed attributes declaration)
    Protected _rgbColor As Integer
    Protected _hslColor As Integer
    Protected _rgbMove As YColorLedMove
    Protected _hslMove As YColorLedMove
    Protected _rgbColorAtPowerOn As Integer
    Protected _blinkSeqSize As Integer
    Protected _blinkSeqMaxSize As Integer
    Protected _blinkSeqSignature As Integer
    Protected _command As String
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
      _blinkSeqSize = BLINKSEQSIZE_INVALID
      _blinkSeqMaxSize = BLINKSEQMAXSIZE_INVALID
      _blinkSeqSignature = BLINKSEQSIGNATURE_INVALID
      _command = COMMAND_INVALID
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
      If (member.name = "blinkSeqSize") Then
        _blinkSeqSize = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "blinkSeqMaxSize") Then
        _blinkSeqMaxSize = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "blinkSeqSignature") Then
        _blinkSeqSignature = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "command") Then
        _command = member.svalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YColorLed private methods declaration)

    REM --- (YColorLed public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current RGB color of the LED.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current RGB color of the LED
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
    '''   Changes the current color of the LED, using an RGB color.
    ''' <para>
    '''   Encoding is done as follows: 0xRRGGBB.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current color of the LED, using an RGB color
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
    '''   Returns the current HSL color of the LED.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current HSL color of the LED
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
    '''   Changes the current color of the LED, using a color HSL.
    ''' <para>
    '''   Encoding is done as follows: 0xHHSSLL.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current color of the LED, using a color HSL
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
    '''   Changes the color that the LED will display by default when the module is turned on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the color that the LED will display by default when the module is turned on
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
    '''   Returns the current length of the blinking sequence.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current length of the blinking sequence
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BLINKSEQSIZE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_blinkSeqSize() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BLINKSEQSIZE_INVALID
        End If
      End If
      Return Me._blinkSeqSize
    End Function

    '''*
    ''' <summary>
    '''   Returns the maximum length of the blinking sequence.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximum length of the blinking sequence
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BLINKSEQMAXSIZE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_blinkSeqMaxSize() As Integer
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BLINKSEQMAXSIZE_INVALID
        End If
      End If
      Return Me._blinkSeqMaxSize
    End Function

    '''*
    ''' <summary>
    '''   Return the blinking sequence signature.
    ''' <para>
    '''   Since blinking
    '''   sequences cannot be read from the device, this can be used
    '''   to detect if a specific blinking sequence is already
    '''   programmed.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BLINKSEQSIGNATURE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_blinkSeqSignature() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BLINKSEQSIGNATURE_INVALID
        End If
      End If
      Return Me._blinkSeqSignature
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
    '''   Retrieves an RGB LED for a given identifier.
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
    '''   This function does not require that the RGB LED is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YColorLed.isOnline()</c> to test if the RGB LED is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an RGB LED by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the RGB LED
    ''' </param>
    ''' <returns>
    '''   a <c>YColorLed</c> object allowing you to drive the RGB LED.
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

    Public Overridable Function sendCommand(command As String) As Integer
      REM //may throw an exception
      Return Me.set_command(command)
    End Function

    '''*
    ''' <summary>
    '''   Add a new transition to the blinking sequence, the move will
    '''   be performed in the HSL space.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="HSLcolor">
    '''   desired HSL color when the traisntion is completed
    ''' </param>
    ''' <param name="msDelay">
    '''   duration of the color transition, in milliseconds.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function addHslMoveToBlinkSeq(HSLcolor As Integer, msDelay As Integer) As Integer
      Return Me.sendCommand("H" + Convert.ToString(HSLcolor) + "," + Convert.ToString(msDelay))
    End Function

    '''*
    ''' <summary>
    '''   Adds a new transition to the blinking sequence, the move is
    '''   performed in the RGB space.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="RGBcolor">
    '''   desired RGB color when the transition is completed
    ''' </param>
    ''' <param name="msDelay">
    '''   duration of the color transition, in milliseconds.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function addRgbMoveToBlinkSeq(RGBcolor As Integer, msDelay As Integer) As Integer
      Return Me.sendCommand("R" + Convert.ToString(RGBcolor) + "," + Convert.ToString(msDelay))
    End Function

    '''*
    ''' <summary>
    '''   Starts the preprogrammed blinking sequence.
    ''' <para>
    '''   The sequence is
    '''   run in a loop until it is stopped by stopBlinkSeq or an explicit
    '''   change.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function startBlinkSeq() As Integer
      Return Me.sendCommand("S")
    End Function

    '''*
    ''' <summary>
    '''   Stops the preprogrammed blinking sequence.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function stopBlinkSeq() As Integer
      Return Me.sendCommand("X")
    End Function

    '''*
    ''' <summary>
    '''   Resets the preprogrammed blinking sequence.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function resetBlinkSeq() As Integer
      Return Me.sendCommand("Z")
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of RGB LEDs started using <c>yFirstColorLed()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YColorLed</c> object, corresponding to
    '''   an RGB LED currently online, or a <c>Nothing</c> pointer
    '''   if there are no more RGB LEDs to enumerate.
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
    '''   Starts the enumeration of RGB LEDs currently accessible.
    ''' <para>
    '''   Use the method <c>YColorLed.nextColorLed()</c> to iterate on
    '''   next RGB LEDs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YColorLed</c> object, corresponding to
    '''   the first RGB LED currently online, or a <c>Nothing</c> pointer
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
  '''   Retrieves an RGB LED for a given identifier.
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
  '''   This function does not require that the RGB LED is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YColorLed.isOnline()</c> to test if the RGB LED is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an RGB LED by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the RGB LED
  ''' </param>
  ''' <returns>
  '''   a <c>YColorLed</c> object allowing you to drive the RGB LED.
  ''' </returns>
  '''/
  Public Function yFindColorLed(ByVal func As String) As YColorLed
    Return YColorLed.FindColorLed(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of RGB LEDs currently accessible.
  ''' <para>
  '''   Use the method <c>YColorLed.nextColorLed()</c> to iterate on
  '''   next RGB LEDs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YColorLed</c> object, corresponding to
  '''   the first RGB LED currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstColorLed() As YColorLed
    Return YColorLed.FirstColorLed()
  End Function


  REM --- (end of ColorLed functions)

End Module
