'*********************************************************************
'*
'* $Id: pic24config.php 18250 2014-11-03 16:54:15Z mvuilleu $
'*
'* Implements yFindBuzzer(), the high-level API for Buzzer functions
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

Module yocto_buzzer

    REM --- (YBuzzer return codes)
    REM --- (end of YBuzzer return codes)
    REM --- (YBuzzer dlldef)
    REM --- (end of YBuzzer dlldef)
  REM --- (YBuzzer globals)

  Public Const Y_FREQUENCY_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_VOLUME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PLAYSEQSIZE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PLAYSEQMAXSIZE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PLAYSEQSIGNATURE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YBuzzerValueCallback(ByVal func As YBuzzer, ByVal value As String)
  Public Delegate Sub YBuzzerTimedReportCallback(ByVal func As YBuzzer, ByVal measure As YMeasure)
  REM --- (end of YBuzzer globals)

  REM --- (YBuzzer class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to
  '''   choose the frequency and volume at which the buzzer must sound.
  ''' <para>
  '''   You can also pre-program a play sequence.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YBuzzer
    Inherits YFunction
    REM --- (end of YBuzzer class start)

    REM --- (YBuzzer definitions)
    Public Const FREQUENCY_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const VOLUME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PLAYSEQSIZE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PLAYSEQMAXSIZE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PLAYSEQSIGNATURE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YBuzzer definitions)

    REM --- (YBuzzer attributes declaration)
    Protected _frequency As Double
    Protected _volume As Integer
    Protected _playSeqSize As Integer
    Protected _playSeqMaxSize As Integer
    Protected _playSeqSignature As Integer
    Protected _command As String
    Protected _valueCallbackBuzzer As YBuzzerValueCallback
    REM --- (end of YBuzzer attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Buzzer"
      REM --- (YBuzzer attributes initialization)
      _frequency = FREQUENCY_INVALID
      _volume = VOLUME_INVALID
      _playSeqSize = PLAYSEQSIZE_INVALID
      _playSeqMaxSize = PLAYSEQMAXSIZE_INVALID
      _playSeqSignature = PLAYSEQSIGNATURE_INVALID
      _command = COMMAND_INVALID
      _valueCallbackBuzzer = Nothing
      REM --- (end of YBuzzer attributes initialization)
    End Sub

    REM --- (YBuzzer private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "frequency") Then
        _frequency = Math.Round(member.ivalue * 1000.0 / 65536.0) / 1000.0
        Return 1
      End If
      If (member.name = "volume") Then
        _volume = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "playSeqSize") Then
        _playSeqSize = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "playSeqMaxSize") Then
        _playSeqMaxSize = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "playSeqSignature") Then
        _playSeqSignature = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "command") Then
        _command = member.svalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YBuzzer private methods declaration)

    REM --- (YBuzzer public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the frequency of the signal sent to the buzzer.
    ''' <para>
    '''   A zero value stops the buzzer.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the frequency of the signal sent to the buzzer
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
    Public Function set_frequency(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("frequency", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the  frequency of the signal sent to the buzzer/speaker.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the  frequency of the signal sent to the buzzer/speaker
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_FREQUENCY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_frequency() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return FREQUENCY_INVALID
        End If
      End If
      Return Me._frequency
    End Function

    '''*
    ''' <summary>
    '''   Returns the volume of the signal sent to the buzzer/speaker.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the volume of the signal sent to the buzzer/speaker
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VOLUME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_volume() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return VOLUME_INVALID
        End If
      End If
      Return Me._volume
    End Function


    '''*
    ''' <summary>
    '''   Changes the volume of the signal sent to the buzzer/speaker.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the volume of the signal sent to the buzzer/speaker
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
    Public Function set_volume(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("volume", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current length of the playing sequence
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current length of the playing sequence
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PLAYSEQSIZE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_playSeqSize() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PLAYSEQSIZE_INVALID
        End If
      End If
      Return Me._playSeqSize
    End Function

    '''*
    ''' <summary>
    '''   Returns the maximum length of the playing sequence
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximum length of the playing sequence
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PLAYSEQMAXSIZE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_playSeqMaxSize() As Integer
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PLAYSEQMAXSIZE_INVALID
        End If
      End If
      Return Me._playSeqMaxSize
    End Function

    '''*
    ''' <summary>
    '''   Returns the playing sequence signature.
    ''' <para>
    '''   As playing
    '''   sequences cannot be read from the device, this can be used
    '''   to detect if a specific playing sequence is already
    '''   programmed.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the playing sequence signature
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PLAYSEQSIGNATURE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_playSeqSignature() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PLAYSEQSIGNATURE_INVALID
        End If
      End If
      Return Me._playSeqSignature
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
    '''   Retrieves a buzzer for a given identifier.
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
    '''   This function does not require that the buzzer is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YBuzzer.isOnline()</c> to test if the buzzer is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a buzzer by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the buzzer
    ''' </param>
    ''' <returns>
    '''   a <c>YBuzzer</c> object allowing you to drive the buzzer.
    ''' </returns>
    '''/
    Public Shared Function FindBuzzer(func As String) As YBuzzer
      Dim obj As YBuzzer
      obj = CType(YFunction._FindFromCache("Buzzer", func), YBuzzer)
      If ((obj Is Nothing)) Then
        obj = New YBuzzer(func)
        YFunction._AddToCache("Buzzer", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YBuzzerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackBuzzer = callback
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
      If (Not (Me._valueCallbackBuzzer Is Nothing)) Then
        Me._valueCallbackBuzzer(Me, value)
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
    '''   Adds a new frequency transition to the playing sequence.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="freq">
    '''   desired frequency when the transition is completed, in Hz
    ''' </param>
    ''' <param name="msDelay">
    '''   duration of the frequency transition, in milliseconds.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function addFreqMoveToPlaySeq(freq As Integer, msDelay As Integer) As Integer
      Return Me.sendCommand("A" + Convert.ToString(freq) + "," + Convert.ToString(msDelay))
    End Function

    '''*
    ''' <summary>
    '''   Adds a pulse to the playing sequence.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="freq">
    '''   pulse frequency, in Hz
    ''' </param>
    ''' <param name="msDuration">
    '''   pulse duration, in milliseconds.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function addPulseToPlaySeq(freq As Integer, msDuration As Integer) As Integer
      Return Me.sendCommand("B" + Convert.ToString(freq) + "," + Convert.ToString(msDuration))
    End Function

    '''*
    ''' <summary>
    '''   Adds a new volume transition to the playing sequence.
    ''' <para>
    '''   Frequency stays untouched:
    '''   if frequency is at zero, the transition has no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="volume">
    '''   desired volume when the transition is completed, as a percentage.
    ''' </param>
    ''' <param name="msDuration">
    '''   duration of the volume transition, in milliseconds.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function addVolMoveToPlaySeq(volume As Integer, msDuration As Integer) As Integer
      Return Me.sendCommand("C" + Convert.ToString(volume) + "," + Convert.ToString(msDuration))
    End Function

    '''*
    ''' <summary>
    '''   Starts the preprogrammed playing sequence.
    ''' <para>
    '''   The sequence
    '''   runs in loop until it is stopped by stopPlaySeq or an explicit
    '''   change.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function startPlaySeq() As Integer
      Return Me.sendCommand("S")
    End Function

    '''*
    ''' <summary>
    '''   Stops the preprogrammed playing sequence and sets the frequency to zero.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function stopPlaySeq() As Integer
      Return Me.sendCommand("X")
    End Function

    '''*
    ''' <summary>
    '''   Resets the preprogrammed playing sequence and sets the frequency to zero.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function resetPlaySeq() As Integer
      Return Me.sendCommand("Z")
    End Function

    '''*
    ''' <summary>
    '''   Activates the buzzer for a short duration.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="frequency">
    '''   pulse frequency, in hertz
    ''' </param>
    ''' <param name="duration">
    '''   pulse duration in millseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function pulse(frequency As Integer, duration As Integer) As Integer
      Return Me.set_command("P" + Convert.ToString(frequency) + "," + Convert.ToString(duration))
    End Function

    '''*
    ''' <summary>
    '''   Makes the buzzer frequency change over a period of time.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="frequency">
    '''   frequency to reach, in hertz. A frequency under 25Hz stops the buzzer.
    ''' </param>
    ''' <param name="duration">
    '''   pulse duration in millseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function freqMove(frequency As Integer, duration As Integer) As Integer
      Return Me.set_command("F" + Convert.ToString(frequency) + "," + Convert.ToString(duration))
    End Function

    '''*
    ''' <summary>
    '''   Makes the buzzer volume change over a period of time, frequency  stays untouched.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="volume">
    '''   volume to reach in %
    ''' </param>
    ''' <param name="duration">
    '''   change duration in millseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function volumeMove(volume As Integer, duration As Integer) As Integer
      Return Me.set_command("V" + Convert.ToString(volume) + "," + Convert.ToString(duration))
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of buzzers started using <c>yFirstBuzzer()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YBuzzer</c> object, corresponding to
    '''   a buzzer currently online, or a <c>null</c> pointer
    '''   if there are no more buzzers to enumerate.
    ''' </returns>
    '''/
    Public Function nextBuzzer() As YBuzzer
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YBuzzer.FindBuzzer(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of buzzers currently accessible.
    ''' <para>
    '''   Use the method <c>YBuzzer.nextBuzzer()</c> to iterate on
    '''   next buzzers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YBuzzer</c> object, corresponding to
    '''   the first buzzer currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstBuzzer() As YBuzzer
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Buzzer", 0, p, size, neededsize, errmsg)
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
      Return YBuzzer.FindBuzzer(serial + "." + funcId)
    End Function

    REM --- (end of YBuzzer public methods declaration)

  End Class

  REM --- (Buzzer functions)

  '''*
  ''' <summary>
  '''   Retrieves a buzzer for a given identifier.
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
  '''   This function does not require that the buzzer is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YBuzzer.isOnline()</c> to test if the buzzer is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a buzzer by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the buzzer
  ''' </param>
  ''' <returns>
  '''   a <c>YBuzzer</c> object allowing you to drive the buzzer.
  ''' </returns>
  '''/
  Public Function yFindBuzzer(ByVal func As String) As YBuzzer
    Return YBuzzer.FindBuzzer(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of buzzers currently accessible.
  ''' <para>
  '''   Use the method <c>YBuzzer.nextBuzzer()</c> to iterate on
  '''   next buzzers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YBuzzer</c> object, corresponding to
  '''   the first buzzer currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstBuzzer() As YBuzzer
    Return YBuzzer.FirstBuzzer()
  End Function


  REM --- (end of Buzzer functions)

End Module
