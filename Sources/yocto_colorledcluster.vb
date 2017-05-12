'*********************************************************************
'*
'* $Id: yocto_colorledcluster.vb 27282 2017-04-25 15:44:42Z seb $
'*
'* Implements yFindColorLedCluster(), the high-level API for ColorLedCluster functions
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

Module yocto_colorledcluster

    REM --- (YColorLedCluster return codes)
    REM --- (end of YColorLedCluster return codes)
    REM --- (YColorLedCluster dlldef)
    REM --- (end of YColorLedCluster dlldef)
  REM --- (YColorLedCluster globals)

  Public Const Y_ACTIVELEDCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_MAXLEDCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BLINKSEQMAXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BLINKSEQMAXSIZE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YColorLedClusterValueCallback(ByVal func As YColorLedCluster, ByVal value As String)
  Public Delegate Sub YColorLedClusterTimedReportCallback(ByVal func As YColorLedCluster, ByVal measure As YMeasure)
  REM --- (end of YColorLedCluster globals)

  REM --- (YColorLedCluster class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface
  '''   allows you to drive a color LED cluster.
  ''' <para>
  '''   Unlike the ColorLed class, the ColorLedCluster
  '''   allows to handle several LEDs at one. Color changes can be done   using RGB coordinates as well as
  '''   HSL coordinates.
  '''   The module performs all conversions form RGB to HSL automatically. It is then
  '''   self-evident to turn on a LED with a given hue and to progressively vary its
  '''   saturation or lightness. If needed, you can find more information on the
  '''   difference between RGB and HSL in the section following this one.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YColorLedCluster
    Inherits YFunction
    REM --- (end of YColorLedCluster class start)

    REM --- (YColorLedCluster definitions)
    Public Const ACTIVELEDCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const MAXLEDCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BLINKSEQMAXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BLINKSEQMAXSIZE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YColorLedCluster definitions)

    REM --- (YColorLedCluster attributes declaration)
    Protected _activeLedCount As Integer
    Protected _maxLedCount As Integer
    Protected _blinkSeqMaxCount As Integer
    Protected _blinkSeqMaxSize As Integer
    Protected _command As String
    Protected _valueCallbackColorLedCluster As YColorLedClusterValueCallback
    REM --- (end of YColorLedCluster attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "ColorLedCluster"
      REM --- (YColorLedCluster attributes initialization)
      _activeLedCount = ACTIVELEDCOUNT_INVALID
      _maxLedCount = MAXLEDCOUNT_INVALID
      _blinkSeqMaxCount = BLINKSEQMAXCOUNT_INVALID
      _blinkSeqMaxSize = BLINKSEQMAXSIZE_INVALID
      _command = COMMAND_INVALID
      _valueCallbackColorLedCluster = Nothing
      REM --- (end of YColorLedCluster attributes initialization)
    End Sub

    REM --- (YColorLedCluster private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("activeLedCount") Then
        _activeLedCount = CInt(json_val.getLong("activeLedCount"))
      End If
      If json_val.has("maxLedCount") Then
        _maxLedCount = CInt(json_val.getLong("maxLedCount"))
      End If
      If json_val.has("blinkSeqMaxCount") Then
        _blinkSeqMaxCount = CInt(json_val.getLong("blinkSeqMaxCount"))
      End If
      If json_val.has("blinkSeqMaxSize") Then
        _blinkSeqMaxSize = CInt(json_val.getLong("blinkSeqMaxSize"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YColorLedCluster private methods declaration)

    REM --- (YColorLedCluster public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the number of LEDs currently handled by the device.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of LEDs currently handled by the device
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ACTIVELEDCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_activeLedCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ACTIVELEDCOUNT_INVALID
        End If
      End If
      res = Me._activeLedCount
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of LEDs currently handled by the device.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the number of LEDs currently handled by the device
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
    Public Function set_activeLedCount(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("activeLedCount", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the maximum number of LEDs that the device can handle.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximum number of LEDs that the device can handle
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MAXLEDCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_maxLedCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MAXLEDCOUNT_INVALID
        End If
      End If
      res = Me._maxLedCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the maximum number of sequences that the device can handle.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximum number of sequences that the device can handle
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BLINKSEQMAXCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_blinkSeqMaxCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BLINKSEQMAXCOUNT_INVALID
        End If
      End If
      res = Me._blinkSeqMaxCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the maximum length of sequences.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximum length of sequences
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BLINKSEQMAXSIZE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_blinkSeqMaxSize() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BLINKSEQMAXSIZE_INVALID
        End If
      End If
      res = Me._blinkSeqMaxSize
      Return res
    End Function

    Public Function get_command() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   Retrieves a RGB LED cluster for a given identifier.
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
    '''   This function does not require that the RGB LED cluster is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YColorLedCluster.isOnline()</c> to test if the RGB LED cluster is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a RGB LED cluster by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the RGB LED cluster
    ''' </param>
    ''' <returns>
    '''   a <c>YColorLedCluster</c> object allowing you to drive the RGB LED cluster.
    ''' </returns>
    '''/
    Public Shared Function FindColorLedCluster(func As String) As YColorLedCluster
      Dim obj As YColorLedCluster
      obj = CType(YFunction._FindFromCache("ColorLedCluster", func), YColorLedCluster)
      If ((obj Is Nothing)) Then
        obj = New YColorLedCluster(func)
        YFunction._AddToCache("ColorLedCluster", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YColorLedClusterValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackColorLedCluster = callback
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
      If (Not (Me._valueCallbackColorLedCluster Is Nothing)) Then
        Me._valueCallbackColorLedCluster(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function sendCommand(command As String) As Integer
      Return Me.set_command(command)
    End Function

    '''*
    ''' <summary>
    '''   Changes the current color of consecutve LEDs in the cluster, using a RGB color.
    ''' <para>
    '''   Encoding is done as follows: 0xRRGGBB.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <param name="rgbValue">
    '''   new color.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_rgbColor(ledIndex As Integer, count As Integer, rgbValue As Integer) As Integer
      Return Me.sendCommand("SR" + Convert.ToString(ledIndex) + "," + Convert.ToString(count) + "," + YAPI._intToHex(rgbValue,1))
    End Function

    '''*
    ''' <summary>
    '''   Changes the  color at device startup of consecutve LEDs in the cluster, using a RGB color.
    ''' <para>
    '''   Encoding is done as follows: 0xRRGGBB.
    '''   Don't forget to call <c>saveLedsConfigAtPowerOn()</c> to make sure the modification is saved in the
    '''   device flash memory.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <param name="rgbValue">
    '''   new color.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_rgbColorAtPowerOn(ledIndex As Integer, count As Integer, rgbValue As Integer) As Integer
      Return Me.sendCommand("SC" + Convert.ToString(ledIndex) + "," + Convert.ToString(count) + "," + YAPI._intToHex(rgbValue,1))
    End Function

    '''*
    ''' <summary>
    '''   Changes the current color of consecutive LEDs in the cluster, using a HSL color.
    ''' <para>
    '''   Encoding is done as follows: 0xHHSSLL.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <param name="hslValue">
    '''   new color.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_hslColor(ledIndex As Integer, count As Integer, hslValue As Integer) As Integer
      Return Me.sendCommand("SH" + Convert.ToString(ledIndex) + "," + Convert.ToString(count) + "," + YAPI._intToHex(hslValue,1))
    End Function

    '''*
    ''' <summary>
    '''   Allows you to modify the current color of a group of adjacent LEDs to another color, in a seamless and
    '''   autonomous manner.
    ''' <para>
    '''   The transition is performed in the RGB space.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <param name="rgbValue">
    '''   new color (0xRRGGBB).
    ''' </param>
    ''' <param name="delay">
    '''   transition duration in ms
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function rgb_move(ledIndex As Integer, count As Integer, rgbValue As Integer, delay As Integer) As Integer
      Return Me.sendCommand("MR" + Convert.ToString(ledIndex) + "," + Convert.ToString(count) + "," + YAPI._intToHex(rgbValue,1) + "," + Convert.ToString(delay))
    End Function

    '''*
    ''' <summary>
    '''   Allows you to modify the current color of a group of adjacent LEDs  to another color, in a seamless and
    '''   autonomous manner.
    ''' <para>
    '''   The transition is performed in the HSL space. In HSL, hue is a circular
    '''   value (0..360°). There are always two paths to perform the transition: by increasing
    '''   or by decreasing the hue. The module selects the shortest transition.
    '''   If the difference is exactly 180°, the module selects the transition which increases
    '''   the hue.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the fisrt affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <param name="hslValue">
    '''   new color (0xHHSSLL).
    ''' </param>
    ''' <param name="delay">
    '''   transition duration in ms
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function hsl_move(ledIndex As Integer, count As Integer, hslValue As Integer, delay As Integer) As Integer
      Return Me.sendCommand("MH" + Convert.ToString(ledIndex) + "," + Convert.ToString(count) + "," + YAPI._intToHex(hslValue,1) + "," + Convert.ToString(delay))
    End Function

    '''*
    ''' <summary>
    '''   Adds an RGB transition to a sequence.
    ''' <para>
    '''   A sequence is a transition list, which can
    '''   be executed in loop by a group of LEDs.  Sequences are persistent and are saved
    '''   in the device flash memory as soon as the <c>saveBlinkSeq()</c> method is called.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   sequence index.
    ''' </param>
    ''' <param name="rgbValue">
    '''   target color (0xRRGGBB)
    ''' </param>
    ''' <param name="delay">
    '''   transition duration in ms
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function addRgbMoveToBlinkSeq(seqIndex As Integer, rgbValue As Integer, delay As Integer) As Integer
      Return Me.sendCommand("AR" + Convert.ToString(seqIndex) + "," + YAPI._intToHex(rgbValue,1) + "," + Convert.ToString(delay))
    End Function

    '''*
    ''' <summary>
    '''   Adds an HSL transition to a sequence.
    ''' <para>
    '''   A sequence is a transition list, which can
    '''   be executed in loop by an group of LEDs.  Sequences are persistant and are saved
    '''   in the device flash memory as soon as the <c>saveBlinkSeq()</c> method is called.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   sequence index.
    ''' </param>
    ''' <param name="hslValue">
    '''   target color (0xHHSSLL)
    ''' </param>
    ''' <param name="delay">
    '''   transition duration in ms
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function addHslMoveToBlinkSeq(seqIndex As Integer, hslValue As Integer, delay As Integer) As Integer
      Return Me.sendCommand("AH" + Convert.ToString(seqIndex) + "," + YAPI._intToHex(hslValue,1) + "," + Convert.ToString(delay))
    End Function

    '''*
    ''' <summary>
    '''   Adds a mirror ending to a sequence.
    ''' <para>
    '''   When the sequence will reach the end of the last
    '''   transition, its running speed will automatically be reversed so that the sequence plays
    '''   in the reverse direction, like in a mirror. After the first transition of the sequence
    '''   is played at the end of the reverse execution, the sequence starts again in
    '''   the initial direction.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   sequence index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function addMirrorToBlinkSeq(seqIndex As Integer) As Integer
      Return Me.sendCommand("AC" + Convert.ToString(seqIndex) + ",0,0")
    End Function

    '''*
    ''' <summary>
    '''   Links adjacent LEDs to a specific sequence.
    ''' <para>
    '''   These LEDs start to execute
    '''   the sequence as soon as  startBlinkSeq is called. It is possible to add an offset
    '''   in the execution: that way we  can have several groups of LED executing the same
    '''   sequence, with a  temporal offset. A LED cannot be linked to more than one sequence.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <param name="seqIndex">
    '''   sequence index.
    ''' </param>
    ''' <param name="offset">
    '''   execution offset in ms.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function linkLedToBlinkSeq(ledIndex As Integer, count As Integer, seqIndex As Integer, offset As Integer) As Integer
      Return Me.sendCommand("LS" + Convert.ToString(ledIndex) + "," + Convert.ToString(count) + "," + Convert.ToString(seqIndex) + "," + Convert.ToString(offset))
    End Function

    '''*
    ''' <summary>
    '''   Links adjacent LEDs to a specific sequence at device poweron.
    ''' <para>
    '''   Don't forget to configure
    '''   the sequence auto start flag as well and call <c>saveLedsConfigAtPowerOn()</c>. It is possible to add an offset
    '''   in the execution: that way we  can have several groups of LEDs executing the same
    '''   sequence, with a  temporal offset. A LED cannot be linked to more than one sequence.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <param name="seqIndex">
    '''   sequence index.
    ''' </param>
    ''' <param name="offset">
    '''   execution offset in ms.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function linkLedToBlinkSeqAtPowerOn(ledIndex As Integer, count As Integer, seqIndex As Integer, offset As Integer) As Integer
      Return Me.sendCommand("LO" + Convert.ToString(ledIndex) + "," + Convert.ToString(count) + "," + Convert.ToString(seqIndex) + "," + Convert.ToString(offset))
    End Function

    '''*
    ''' <summary>
    '''   Links adjacent LEDs to a specific sequence.
    ''' <para>
    '''   These LED start to execute
    '''   the sequence as soon as  startBlinkSeq is called. This function automatically
    '''   introduces a shift between LEDs so that the specified number of sequence periods
    '''   appears on the group of LEDs (wave effect).
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <param name="seqIndex">
    '''   sequence index.
    ''' </param>
    ''' <param name="periods">
    '''   number of periods to show on LEDs.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function linkLedToPeriodicBlinkSeq(ledIndex As Integer, count As Integer, seqIndex As Integer, periods As Integer) As Integer
      Return Me.sendCommand("LP" + Convert.ToString(ledIndex) + "," + Convert.ToString(count) + "," + Convert.ToString(seqIndex) + "," + Convert.ToString(periods))
    End Function

    '''*
    ''' <summary>
    '''   Unlinks adjacent LEDs from a  sequence.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first affected LED.
    ''' </param>
    ''' <param name="count">
    '''   affected LED count.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function unlinkLedFromBlinkSeq(ledIndex As Integer, count As Integer) As Integer
      Return Me.sendCommand("US" + Convert.ToString(ledIndex) + "," + Convert.ToString(count))
    End Function

    '''*
    ''' <summary>
    '''   Starts a sequence execution: every LED linked to that sequence starts to
    '''   run it in a loop.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the sequence to start.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function startBlinkSeq(seqIndex As Integer) As Integer
      Return Me.sendCommand("SS" + Convert.ToString(seqIndex))
    End Function

    '''*
    ''' <summary>
    '''   Stops a sequence execution.
    ''' <para>
    '''   If started again, the execution
    '''   restarts from the beginning.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the sequence to stop.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function stopBlinkSeq(seqIndex As Integer) As Integer
      Return Me.sendCommand("XS" + Convert.ToString(seqIndex))
    End Function

    '''*
    ''' <summary>
    '''   Stops a sequence execution and resets its contents.
    ''' <para>
    '''   Leds linked to this
    '''   sequence are not automatically updated anymore.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the sequence to reset
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function resetBlinkSeq(seqIndex As Integer) As Integer
      Return Me.sendCommand("ZS" + Convert.ToString(seqIndex))
    End Function

    '''*
    ''' <summary>
    '''   Configures a sequence to make it start automatically at device
    '''   startup.
    ''' <para>
    '''   Don't forget to call <c>saveBlinkSeq()</c> to make sure the
    '''   modification is saved in the device flash memory.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the sequence to reset.
    ''' </param>
    ''' <param name="autostart">
    '''   0 to keep the sequence turned off and 1 to start it automatically.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_blinkSeqStateAtPowerOn(seqIndex As Integer, autostart As Integer) As Integer
      Return Me.sendCommand("AS" + Convert.ToString(seqIndex) + "," + Convert.ToString(autostart))
    End Function

    '''*
    ''' <summary>
    '''   Changes the execution speed of a sequence.
    ''' <para>
    '''   The natural execution speed is 1000 per
    '''   thousand. If you configure a slower speed, you can play the sequence in slow-motion.
    '''   If you set a negative speed, you can play the sequence in reverse direction.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the sequence to start.
    ''' </param>
    ''' <param name="speed">
    '''   sequence running speed (-1000...1000).
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_blinkSeqSpeed(seqIndex As Integer, speed As Integer) As Integer
      Return Me.sendCommand("CS" + Convert.ToString(seqIndex) + "," + Convert.ToString(speed))
    End Function

    '''*
    ''' <summary>
    '''   Saves the LEDs power-on configuration.
    ''' <para>
    '''   This includes the start-up color or
    '''   sequence binding for all LEDs. Warning: if some LEDs are linked to a sequence, the
    '''   method <c>saveBlinkSeq()</c> must also be called to save the sequence definition.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function saveLedsConfigAtPowerOn() As Integer
      Return Me.sendCommand("WL")
    End Function

    Public Overridable Function saveLedsState() As Integer
      Return Me.sendCommand("WL")
    End Function

    '''*
    ''' <summary>
    '''   Saves the definition of a sequence.
    ''' <para>
    '''   Warning: only sequence steps and flags are saved.
    '''   to save the LEDs startup bindings, the method <c>saveLedsConfigAtPowerOn()</c>
    '''   must be called.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the sequence to start.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function saveBlinkSeq(seqIndex As Integer) As Integer
      Return Me.sendCommand("WS" + Convert.ToString(seqIndex))
    End Function

    '''*
    ''' <summary>
    '''   Sends a binary buffer to the LED RGB buffer, as is.
    ''' <para>
    '''   First three bytes are RGB components for LED specified as parameter, the
    '''   next three bytes for the next LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first LED which should be updated
    ''' </param>
    ''' <param name="buff">
    '''   the binary buffer to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_rgbColorBuffer(ledIndex As Integer, buff As Byte()) As Integer
      Return Me._upload("rgb:0:" + Convert.ToString(ledIndex), buff)
    End Function

    '''*
    ''' <summary>
    '''   Sends 24bit RGB colors (provided as a list of integers) to the LED RGB buffer, as is.
    ''' <para>
    '''   The first number represents the RGB value of the LED specified as parameter, the second
    '''   number represents the RGB value of the next LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first LED which should be updated
    ''' </param>
    ''' <param name="rgbList">
    '''   a list of 24bit RGB codes, in the form 0xRRGGBB
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_rgbColorArray(ledIndex As Integer, rgbList As List(Of Integer)) As Integer
      Dim listlen As Integer = 0
      Dim buff As Byte()
      Dim idx As Integer = 0
      Dim rgb As Integer = 0
      Dim res As Integer = 0
      listlen = rgbList.Count
      ReDim buff(3*listlen-1)
      idx = 0
      While (idx < listlen)
        rgb = rgbList(idx)
        buff( 3*idx) = Convert.ToByte(((((rgb) >> (16))) And (255)) And &HFF)
        buff( 3*idx+1) = Convert.ToByte(((((rgb) >> (8))) And (255)) And &HFF)
        buff( 3*idx+2) = Convert.ToByte(((rgb) And (255)) And &HFF)
        idx = idx + 1
      End While

      res = Me._upload("rgb:0:" + Convert.ToString(ledIndex), buff)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sets up a smooth RGB color transition to the specified pixel-by-pixel list of RGB
    '''   color codes.
    ''' <para>
    '''   The first color code represents the target RGB value of the first LED,
    '''   the next color code represents the target value of the next LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="rgbList">
    '''   a list of target 24bit RGB codes, in the form 0xRRGGBB
    ''' </param>
    ''' <param name="delay">
    '''   transition duration in ms
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function rgbArray_move(rgbList As List(Of Integer), delay As Integer) As Integer
      Dim listlen As Integer = 0
      Dim buff As Byte()
      Dim idx As Integer = 0
      Dim rgb As Integer = 0
      Dim res As Integer = 0
      listlen = rgbList.Count
      ReDim buff(3*listlen-1)
      idx = 0
      While (idx < listlen)
        rgb = rgbList(idx)
        buff( 3*idx) = Convert.ToByte(((((rgb) >> (16))) And (255)) And &HFF)
        buff( 3*idx+1) = Convert.ToByte(((((rgb) >> (8))) And (255)) And &HFF)
        buff( 3*idx+2) = Convert.ToByte(((rgb) And (255)) And &HFF)
        idx = idx + 1
      End While

      res = Me._upload("rgb:" + Convert.ToString(delay), buff)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends a binary buffer to the LED HSL buffer, as is.
    ''' <para>
    '''   First three bytes are HSL components for the LED specified as parameter, the
    '''   next three bytes for the second LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first LED which should be updated
    ''' </param>
    ''' <param name="buff">
    '''   the binary buffer to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_hslColorBuffer(ledIndex As Integer, buff As Byte()) As Integer
      Return Me._upload("hsl:0:" + Convert.ToString(ledIndex), buff)
    End Function

    '''*
    ''' <summary>
    '''   Sends 24bit HSL colors (provided as a list of integers) to the LED HSL buffer, as is.
    ''' <para>
    '''   The first number represents the HSL value of the LED specified as parameter, the second number represents
    '''   the HSL value of the second LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first LED which should be updated
    ''' </param>
    ''' <param name="hslList">
    '''   a list of 24bit HSL codes, in the form 0xHHSSLL
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_hslColorArray(ledIndex As Integer, hslList As List(Of Integer)) As Integer
      Dim listlen As Integer = 0
      Dim buff As Byte()
      Dim idx As Integer = 0
      Dim hsl As Integer = 0
      Dim res As Integer = 0
      listlen = hslList.Count
      ReDim buff(3*listlen-1)
      idx = 0
      While (idx < listlen)
        hsl = hslList(idx)
        buff( 3*idx) = Convert.ToByte(((((hsl) >> (16))) And (255)) And &HFF)
        buff( 3*idx+1) = Convert.ToByte(((((hsl) >> (8))) And (255)) And &HFF)
        buff( 3*idx+2) = Convert.ToByte(((hsl) And (255)) And &HFF)
        idx = idx + 1
      End While

      res = Me._upload("hsl:0:" + Convert.ToString(ledIndex), buff)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sets up a smooth HSL color transition to the specified pixel-by-pixel list of HSL
    '''   color codes.
    ''' <para>
    '''   The first color code represents the target HSL value of the first LED,
    '''   the second color code represents the target value of the second LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="hslList">
    '''   a list of target 24bit HSL codes, in the form 0xHHSSLL
    ''' </param>
    ''' <param name="delay">
    '''   transition duration in ms
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function hslArray_move(hslList As List(Of Integer), delay As Integer) As Integer
      Dim listlen As Integer = 0
      Dim buff As Byte()
      Dim idx As Integer = 0
      Dim hsl As Integer = 0
      Dim res As Integer = 0
      listlen = hslList.Count
      ReDim buff(3*listlen-1)
      idx = 0
      While (idx < listlen)
        hsl = hslList(idx)
        buff( 3*idx) = Convert.ToByte(((((hsl) >> (16))) And (255)) And &HFF)
        buff( 3*idx+1) = Convert.ToByte(((((hsl) >> (8))) And (255)) And &HFF)
        buff( 3*idx+2) = Convert.ToByte(((hsl) And (255)) And &HFF)
        idx = idx + 1
      End While

      res = Me._upload("hsl:" + Convert.ToString(delay), buff)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns a binary buffer with content from the LED RGB buffer, as is.
    ''' <para>
    '''   First three bytes are RGB components for the first LED in the interval,
    '''   the next three bytes for the second LED in the interval, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first LED which should be returned
    ''' </param>
    ''' <param name="count">
    '''   number of LEDs which should be returned
    ''' </param>
    ''' <returns>
    '''   a binary buffer with RGB components of selected LEDs.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty binary buffer.
    ''' </para>
    '''/
    Public Overridable Function get_rgbColorBuffer(ledIndex As Integer, count As Integer) As Byte()
      Return Me._download("rgb.bin?typ=0&pos=" + Convert.ToString(3*ledIndex) + "&len=" + Convert.ToString(3*count))
    End Function

    '''*
    ''' <summary>
    '''   Returns a list on 24bit RGB color values with the current colors displayed on
    '''   the RGB leds.
    ''' <para>
    '''   The first number represents the RGB value of the first LED,
    '''   the second number represents the RGB value of the second LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first LED which should be returned
    ''' </param>
    ''' <param name="count">
    '''   number of LEDs which should be returned
    ''' </param>
    ''' <returns>
    '''   a list of 24bit color codes with RGB components of selected LEDs, as 0xRRGGBB.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_rgbColorArray(ledIndex As Integer, count As Integer) As List(Of Integer)
      Dim buff As Byte()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim idx As Integer = 0
      Dim r As Integer = 0
      Dim g As Integer = 0
      Dim b As Integer = 0

      buff = Me._download("rgb.bin?typ=0&pos=" + Convert.ToString(3*ledIndex) + "&len=" + Convert.ToString(3*count))
      res.Clear()

      idx = 0
      While (idx < count)
        r = buff(3*idx)
        g = buff(3*idx+1)
        b = buff(3*idx+2)
        res.Add(r*65536+g*256+b)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns a list on 24bit RGB color values with the RGB LEDs startup colors.
    ''' <para>
    '''   The first number represents the startup RGB value of the first LED,
    '''   the second number represents the RGB value of the second LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first LED  which should be returned
    ''' </param>
    ''' <param name="count">
    '''   number of LEDs which should be returned
    ''' </param>
    ''' <returns>
    '''   a list of 24bit color codes with RGB components of selected LEDs, as 0xRRGGBB.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_rgbColorArrayAtPowerOn(ledIndex As Integer, count As Integer) As List(Of Integer)
      Dim buff As Byte()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim idx As Integer = 0
      Dim r As Integer = 0
      Dim g As Integer = 0
      Dim b As Integer = 0

      buff = Me._download("rgb.bin?typ=4&pos=" + Convert.ToString(3*ledIndex) + "&len=" + Convert.ToString(3*count))
      res.Clear()

      idx = 0
      While (idx < count)
        r = buff(3*idx)
        g = buff(3*idx+1)
        b = buff(3*idx+2)
        res.Add(r*65536+g*256+b)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns a list on sequence index for each RGB LED.
    ''' <para>
    '''   The first number represents the
    '''   sequence index for the the first LED, the second number represents the sequence
    '''   index for the second LED, etc.
    ''' </para>
    ''' </summary>
    ''' <param name="ledIndex">
    '''   index of the first LED which should be returned
    ''' </param>
    ''' <param name="count">
    '''   number of LEDs which should be returned
    ''' </param>
    ''' <returns>
    '''   a list of integers with sequence index
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_linkedSeqArray(ledIndex As Integer, count As Integer) As List(Of Integer)
      Dim buff As Byte()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim idx As Integer = 0
      Dim seq As Integer = 0

      buff = Me._download("rgb.bin?typ=1&pos=" + Convert.ToString(ledIndex) + "&len=" + Convert.ToString(count))
      res.Clear()

      idx = 0
      While (idx < count)
        seq = buff(idx)
        res.Add(seq)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns a list on 32 bit signatures for specified blinking sequences.
    ''' <para>
    '''   Since blinking sequences cannot be read from the device, this can be used
    '''   to detect if a specific blinking sequence is already programmed.
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the first blinking sequence which should be returned
    ''' </param>
    ''' <param name="count">
    '''   number of blinking sequences which should be returned
    ''' </param>
    ''' <returns>
    '''   a list of 32 bit integer signatures
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_blinkSeqSignatures(seqIndex As Integer, count As Integer) As List(Of Integer)
      Dim buff As Byte()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim idx As Integer = 0
      Dim hh As Integer = 0
      Dim hl As Integer = 0
      Dim lh As Integer = 0
      Dim ll As Integer = 0

      buff = Me._download("rgb.bin?typ=2&pos=" + Convert.ToString(4*seqIndex) + "&len=" + Convert.ToString(4*count))
      res.Clear()

      idx = 0
      While (idx < count)
        hh = buff(4*idx)
        hl = buff(4*idx+1)
        lh = buff(4*idx+2)
        ll = buff(4*idx+3)
        res.Add(((hh) << (24))+((hl) << (16))+((lh) << (8))+ll)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of integers with the current speed for specified blinking sequences.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the first sequence speed which should be returned
    ''' </param>
    ''' <param name="count">
    '''   number of sequence speeds which should be returned
    ''' </param>
    ''' <returns>
    '''   a list of integers, 0 for sequences turned off and 1 for sequences running
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_blinkSeqStateSpeed(seqIndex As Integer, count As Integer) As List(Of Integer)
      Dim buff As Byte()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim idx As Integer = 0
      Dim lh As Integer = 0
      Dim ll As Integer = 0

      buff = Me._download("rgb.bin?typ=6&pos=" + Convert.ToString(seqIndex) + "&len=" + Convert.ToString(count))
      res.Clear()

      idx = 0
      While (idx < count)
        lh = buff(2*idx)
        ll = buff(2*idx+1)
        res.Add(((lh) << (8))+ll)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of integers with the "auto-start at power on" flag state for specified blinking sequences.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the first blinking sequence which should be returned
    ''' </param>
    ''' <param name="count">
    '''   number of blinking sequences which should be returned
    ''' </param>
    ''' <returns>
    '''   a list of integers, 0 for sequences turned off and 1 for sequences running
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_blinkSeqStateAtPowerOn(seqIndex As Integer, count As Integer) As List(Of Integer)
      Dim buff As Byte()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim idx As Integer = 0
      Dim started As Integer = 0

      buff = Me._download("rgb.bin?typ=5&pos=" + Convert.ToString(seqIndex) + "&len=" + Convert.ToString(count))
      res.Clear()

      idx = 0
      While (idx < count)
        started = buff(idx)
        res.Add(started)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of integers with the started state for specified blinking sequences.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="seqIndex">
    '''   index of the first blinking sequence which should be returned
    ''' </param>
    ''' <param name="count">
    '''   number of blinking sequences which should be returned
    ''' </param>
    ''' <returns>
    '''   a list of integers, 0 for sequences turned off and 1 for sequences running
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_blinkSeqState(seqIndex As Integer, count As Integer) As List(Of Integer)
      Dim buff As Byte()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim idx As Integer = 0
      Dim started As Integer = 0

      buff = Me._download("rgb.bin?typ=3&pos=" + Convert.ToString(seqIndex) + "&len=" + Convert.ToString(count))
      res.Clear()

      idx = 0
      While (idx < count)
        started = buff(idx)
        res.Add(started)
        idx = idx + 1
      End While

      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of RGB LED clusters started using <c>yFirstColorLedCluster()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YColorLedCluster</c> object, corresponding to
    '''   a RGB LED cluster currently online, or a <c>Nothing</c> pointer
    '''   if there are no more RGB LED clusters to enumerate.
    ''' </returns>
    '''/
    Public Function nextColorLedCluster() As YColorLedCluster
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YColorLedCluster.FindColorLedCluster(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of RGB LED clusters currently accessible.
    ''' <para>
    '''   Use the method <c>YColorLedCluster.nextColorLedCluster()</c> to iterate on
    '''   next RGB LED clusters.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YColorLedCluster</c> object, corresponding to
    '''   the first RGB LED cluster currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstColorLedCluster() As YColorLedCluster
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("ColorLedCluster", 0, p, size, neededsize, errmsg)
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
      Return YColorLedCluster.FindColorLedCluster(serial + "." + funcId)
    End Function

    REM --- (end of YColorLedCluster public methods declaration)

  End Class

  REM --- (ColorLedCluster functions)

  '''*
  ''' <summary>
  '''   Retrieves a RGB LED cluster for a given identifier.
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
  '''   This function does not require that the RGB LED cluster is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YColorLedCluster.isOnline()</c> to test if the RGB LED cluster is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a RGB LED cluster by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the RGB LED cluster
  ''' </param>
  ''' <returns>
  '''   a <c>YColorLedCluster</c> object allowing you to drive the RGB LED cluster.
  ''' </returns>
  '''/
  Public Function yFindColorLedCluster(ByVal func As String) As YColorLedCluster
    Return YColorLedCluster.FindColorLedCluster(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of RGB LED clusters currently accessible.
  ''' <para>
  '''   Use the method <c>YColorLedCluster.nextColorLedCluster()</c> to iterate on
  '''   next RGB LED clusters.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YColorLedCluster</c> object, corresponding to
  '''   the first RGB LED cluster currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstColorLedCluster() As YColorLedCluster
    Return YColorLedCluster.FirstColorLedCluster()
  End Function


  REM --- (end of ColorLedCluster functions)

End Module
