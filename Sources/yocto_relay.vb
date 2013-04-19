'*********************************************************************
'*
'* $Id: yocto_relay.vb 10401 2013-03-17 13:17:31Z martinm $
'*
'* Implements yFindRelay(), the high-level API for Relay functions
'*
'* - - - - - - - - - License information: - - - - - - - - - 
'*
'* Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'*
'* 1) If you have obtained this file from www.yoctopuce.com,
'*    Yoctopuce Sarl licenses to you (hereafter Licensee) the
'*    right to use, modify, copy, and integrate this source file
'*    into your own solution for the sole purpose of interfacing
'*    a Yoctopuce product with Licensee's solution.
'*
'*    The use of this file and all relationship between Yoctopuce 
'*    and Licensee are governed by Yoctopuce General Terms and 
'*    Conditions.
'*
'*    THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'*    WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING 
'*    WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS 
'*    FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'*    EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'*    INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA, 
'*    COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR 
'*    SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT 
'*    LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'*    CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'*    BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'*    WARRANTY, OR OTHERWISE.
'*
'* 2) If your intent is not to interface with Yoctopuce products,
'*    you are not entitled to use, read or create any derived
'*    material from this source file.
'*
'*********************************************************************/


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_relay

  REM --- (YRelay definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YRelay, ByVal value As String)

Public Class YRelayDelayedPulse
  Public target As System.Int64 = YAPI.INVALID_LONG
  Public ms As System.Int64 = YAPI.INVALID_LONG
  Public moving As System.Int64 = YAPI.INVALID_LONG
End Class


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  REM Y_STATE is defined in yocto_api.vb
  REM Y_OUTPUT is defined in yocto_api.vb
  Public Const Y_PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_COUNTDOWN_INVALID As Long = YAPI.INVALID_LONG

  Public Y_DELAYEDPULSETIMER_INVALID As YRelayDelayedPulse

  REM --- (end of YRelay definitions)

  REM --- (YRelay implementation)

  Private _RelayCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to switch the relay state.
  ''' <para>
  '''   This change is not persistent: the relay will automatically return to its idle position
  '''   whenever power is lost or if the module is restarted.
  '''   The library can also generate automatically short pulses of determined duration.
  '''   On devices with two output for each relay (double throw), the two outputs are named A and B,
  '''   with output A corresponding to the idle position (at power off) and the output B corresponding to the
  '''   active state. If you prefer the alternate default state, simply switch your cables on the board.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YRelay
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const STATE_A = 0
    Public Const STATE_B = 1
    Public Const STATE_INVALID = -1

    Public Const OUTPUT_OFF = 0
    Public Const OUTPUT_ON = 1
    Public Const OUTPUT_INVALID = -1

    Public Const PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
    Public Const COUNTDOWN_INVALID As Long = YAPI.INVALID_LONG

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _state As Long
    Protected _output As Long
    Protected _pulseTimer As Long
    Protected _delayedPulseTimer As YRelayDelayedPulse
    Protected _countdown As Long

    Public Sub New(ByVal func As String)
      MyBase.new("Relay", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _state = Y_STATE_INVALID
      _output = Y_OUTPUT_INVALID
      _pulseTimer = Y_PULSETIMER_INVALID
      _delayedPulseTimer = New YRelayDelayedPulse()
      _countdown = Y_COUNTDOWN_INVALID
    End Sub

    Protected Overrides Function _parse(ByRef j As TJSONRECORD) As Integer
      Dim member As TJSONRECORD
      Dim i As Integer
      If (j.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then
        Return -1
      End If
      For i = 0 To j.membercount - 1
        member = j.members(i)
        If (member.name = "logicalName") Then
          _logicalName = member.svalue
        ElseIf (member.name = "advertisedValue") Then
          _advertisedValue = member.svalue
        ElseIf (member.name = "state") Then
          If (member.ivalue > 0) Then _state = 1 Else _state = 0
        ElseIf (member.name = "output") Then
          If (member.ivalue > 0) Then _output = 1 Else _output = 0
        ElseIf (member.name = "pulseTimer") Then
          _pulseTimer = CLng(member.ivalue)
        ElseIf (member.name = "delayedPulseTimer") Then
          If (member.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then 
             _parse = -1
             Exit Function
          End If
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l=0 To member.membercount-1
             submemb = member.members(l)
             If (submemb.name = "moving") Then
                _delayedPulseTimer.moving = submemb.ivalue
             ElseIf (submemb.name = "target") Then
                _delayedPulseTimer.target = submemb.ivalue
             ElseIf (submemb.name = "ms") Then
                _delayedPulseTimer.ms = submemb.ivalue
             End If
          Next l
        ElseIf (member.name = "countdown") Then
          _countdown = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the relay.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the relay
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LOGICALNAME_INVALID
        End If
      End If
      Return _logicalName
    End Function

    '''*
    ''' <summary>
    '''   Changes the logical name of the relay.
    ''' <para>
    '''   You can use <c>yCheckLogicalName()</c>
    '''   prior to this call to make sure that your parameter is valid.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the logical name of the relay
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
    Public Function set_logicalName(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("logicalName", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of the relay (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the relay (no more than 6 characters)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ADVERTISEDVALUE_INVALID
        End If
      End If
      Return _advertisedValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of the relays (A for the idle position, B for the active position).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_STATE_A</c> or <c>Y_STATE_B</c>, according to the state of the relays (A for the idle
    '''   position, B for the active position)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_STATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_state() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_STATE_INVALID
        End If
      End If
      Return CType(_state,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the state of the relays (A for the idle position, B for the active position).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_STATE_A</c> or <c>Y_STATE_B</c>, according to the state of the relays (A for the idle
    '''   position, B for the active position)
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
    Public Function set_state(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("state", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the output state of the relays, when used as a simple switch (single throw).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_OUTPUT_OFF</c> or <c>Y_OUTPUT_ON</c>, according to the output state of the relays, when
    '''   used as a simple switch (single throw)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OUTPUT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_output() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_OUTPUT_INVALID
        End If
      End If
      Return CType(_output,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the output state of the relays, when used as a simple switch (single throw).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_OUTPUT_OFF</c> or <c>Y_OUTPUT_ON</c>, according to the output state of the relays, when
    '''   used as a simple switch (single throw)
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
    Public Function set_output(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("output", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of milliseconds remaining before the relays is returned to idle position
    '''   (state A), during a measured pulse generation.
    ''' <para>
    '''   When there is no ongoing pulse, returns zero.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of milliseconds remaining before the relays is returned to idle position
    '''   (state A), during a measured pulse generation
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PULSETIMER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pulseTimer() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PULSETIMER_INVALID
        End If
      End If
      Return _pulseTimer
    End Function

    Public Function set_pulseTimer(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pulseTimer", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Sets the relay to output B (active) for a specified duration, then brings it
    '''   automatically back to output A (idle state).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="ms_duration">
    '''   pulse duration, in millisecondes
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
    Public Function pulse(ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(ms_duration))
      Return _setAttr("pulseTimer", rest_val)
    End Function

    Public Function get_delayedPulseTimer() As YRelayDelayedPulse
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_DELAYEDPULSETIMER_INVALID
        End If
      End If
      Return _delayedPulseTimer
    End Function

    Public Function set_delayedPulseTimer(ByVal newval As YRelayDelayedPulse) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target))+":"+Ltrim(Str(newval.ms))
      Return _setAttr("delayedPulseTimer", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Schedules a pulse.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="ms_delay">
    '''   waiting time before the pulse, in millisecondes
    ''' </param>
    ''' <param name="ms_duration">
    '''   pulse duration, in millisecondes
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
    Public Function delayedPulse(ByVal ms_delay As Integer,ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(ms_delay))+":"+Ltrim(Str(ms_duration))
      Return _setAttr("delayedPulseTimer", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of milliseconds remaining before a pulse (delayedPulse() call)
    '''   When there is no scheduled pulse, returns zero.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of milliseconds remaining before a pulse (delayedPulse() call)
    '''   When there is no scheduled pulse, returns zero
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_COUNTDOWN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_countdown() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_COUNTDOWN_INVALID
        End If
      End If
      Return _countdown
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of relays started using <c>yFirstRelay()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRelay</c> object, corresponding to
    '''   a relay currently online, or a <c>null</c> pointer
    '''   if there are no more relays to enumerate.
    ''' </returns>
    '''/
    Public Function nextRelay() as YRelay
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindRelay(hwid)
    End Function

    '''*
    ''' <summary>
    '''   comment from .
    ''' <para>
    '''   yc definition
    ''' </para>
    ''' </summary>
    '''/
  Public Overloads Sub registerValueCallback(ByVal callback As UpdateCallback)
   If (callback IsNot Nothing) Then
     registerFuncCallback(Me)
   Else
     unregisterFuncCallback(Me)
   End If
   _callback = callback
  End Sub

  Public Sub set_callback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Sub setCallback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Overrides Sub advertiseValue(ByVal value As String)
    If (_callback IsNot Nothing) Then _callback(Me, value)
  End Sub


    '''*
    ''' <summary>
    '''   Retrieves a relay for a given identifier.
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
    '''   This function does not require that the relay is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YRelay.isOnline()</c> to test if the relay is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a relay by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the relay
    ''' </param>
    ''' <returns>
    '''   a <c>YRelay</c> object allowing you to drive the relay.
    ''' </returns>
    '''/
    Public Shared Function FindRelay(ByVal func As String) As YRelay
      Dim res As YRelay
      If (_RelayCache.ContainsKey(func)) Then
        Return CType(_RelayCache(func), YRelay)
      End If
      res = New YRelay(func)
      _RelayCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of relays currently accessible.
    ''' <para>
    '''   Use the method <c>YRelay.nextRelay()</c> to iterate on
    '''   next relays.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRelay</c> object, corresponding to
    '''   the first relay currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstRelay() As YRelay
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Relay", 0, p, size, neededsize, errmsg)
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
      Return YRelay.FindRelay(serial + "." + funcId)
    End Function

    REM --- (end of YRelay implementation)

  End Class

  REM --- (Relay functions)

  '''*
  ''' <summary>
  '''   Retrieves a relay for a given identifier.
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
  '''   This function does not require that the relay is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YRelay.isOnline()</c> to test if the relay is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a relay by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the relay
  ''' </param>
  ''' <returns>
  '''   a <c>YRelay</c> object allowing you to drive the relay.
  ''' </returns>
  '''/
  Public Function yFindRelay(ByVal func As String) As YRelay
    Return YRelay.FindRelay(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of relays currently accessible.
  ''' <para>
  '''   Use the method <c>YRelay.nextRelay()</c> to iterate on
  '''   next relays.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YRelay</c> object, corresponding to
  '''   the first relay currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstRelay() As YRelay
    Return YRelay.FirstRelay()
  End Function

  Private Sub _RelayCleanup()
  End Sub


  REM --- (end of Relay functions)

End Module
