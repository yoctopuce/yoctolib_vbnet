'*********************************************************************
'*
'* $Id: yocto_vsource.vb 10263 2013-03-11 17:25:38Z seb $
'*
'* Implements yFindVSource(), the high-level API for VSource functions
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

Module yocto_vsource

  REM --- (YVSource definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YVSource, ByVal value As String)

Public Class YVSourceMove
  Public target As System.Int64 = YAPI.INVALID_LONG
  Public ms As System.Int64 = YAPI.INVALID_LONG
  Public moving As System.Int64 = YAPI.INVALID_LONG
End Class

Public Class YVSourcePulse
  Public target As System.Int64 = YAPI.INVALID_LONG
  Public ms As System.Int64 = YAPI.INVALID_LONG
  Public moving As System.Int64 = YAPI.INVALID_LONG
End Class


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_UNIT_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_VOLTAGE_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_FAILURE_FALSE = 0
  Public Const Y_FAILURE_TRUE = 1
  Public Const Y_FAILURE_INVALID = -1

  Public Const Y_OVERHEAT_FALSE = 0
  Public Const Y_OVERHEAT_TRUE = 1
  Public Const Y_OVERHEAT_INVALID = -1

  Public Const Y_OVERCURRENT_FALSE = 0
  Public Const Y_OVERCURRENT_TRUE = 1
  Public Const Y_OVERCURRENT_INVALID = -1

  Public Const Y_OVERLOAD_FALSE = 0
  Public Const Y_OVERLOAD_TRUE = 1
  Public Const Y_OVERLOAD_INVALID = -1

  Public Const Y_REGULATIONFAILURE_FALSE = 0
  Public Const Y_REGULATIONFAILURE_TRUE = 1
  Public Const Y_REGULATIONFAILURE_INVALID = -1

  Public Const Y_EXTPOWERFAILURE_FALSE = 0
  Public Const Y_EXTPOWERFAILURE_TRUE = 1
  Public Const Y_EXTPOWERFAILURE_INVALID = -1


  Public Y_MOVE_INVALID As YVSourceMove
  Public Y_PULSETIMER_INVALID As YVSourcePulse

  REM --- (end of YVSource definitions)

  REM --- (YVSource implementation)

  Private _VSourceCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   Yoctopuce application programming interface allows you to control
  '''   the module voltage output.
  ''' <para>
  '''   You affect absolute output values or make
  '''   transitions
  ''' </para>
  ''' </summary>
  '''/
  Public Class YVSource
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const UNIT_INVALID As String = YAPI.INVALID_STRING
    Public Const VOLTAGE_INVALID As Integer = YAPI.INVALID_INT
    Public Const FAILURE_FALSE = 0
    Public Const FAILURE_TRUE = 1
    Public Const FAILURE_INVALID = -1

    Public Const OVERHEAT_FALSE = 0
    Public Const OVERHEAT_TRUE = 1
    Public Const OVERHEAT_INVALID = -1

    Public Const OVERCURRENT_FALSE = 0
    Public Const OVERCURRENT_TRUE = 1
    Public Const OVERCURRENT_INVALID = -1

    Public Const OVERLOAD_FALSE = 0
    Public Const OVERLOAD_TRUE = 1
    Public Const OVERLOAD_INVALID = -1

    Public Const REGULATIONFAILURE_FALSE = 0
    Public Const REGULATIONFAILURE_TRUE = 1
    Public Const REGULATIONFAILURE_INVALID = -1

    Public Const EXTPOWERFAILURE_FALSE = 0
    Public Const EXTPOWERFAILURE_TRUE = 1
    Public Const EXTPOWERFAILURE_INVALID = -1


    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _unit As String
    Protected _voltage As Long
    Protected _failure As Long
    Protected _overHeat As Long
    Protected _overCurrent As Long
    Protected _overLoad As Long
    Protected _regulationFailure As Long
    Protected _extPowerFailure As Long
    Protected _move As YVSourceMove
    Protected _pulseTimer As YVSourcePulse

    Public Sub New(ByVal func As String)
      MyBase.new("VSource", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _unit = Y_UNIT_INVALID
      _voltage = Y_VOLTAGE_INVALID
      _failure = Y_FAILURE_INVALID
      _overHeat = Y_OVERHEAT_INVALID
      _overCurrent = Y_OVERCURRENT_INVALID
      _overLoad = Y_OVERLOAD_INVALID
      _regulationFailure = Y_REGULATIONFAILURE_INVALID
      _extPowerFailure = Y_EXTPOWERFAILURE_INVALID
      _move = New YVSourceMove()
      _pulseTimer = New YVSourcePulse()
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
        ElseIf (member.name = "unit") Then
          _unit = member.svalue
        ElseIf (member.name = "voltage") Then
          _voltage = member.ivalue
        ElseIf (member.name = "failure") Then
          If (member.ivalue > 0) Then _failure = 1 Else _failure = 0
        ElseIf (member.name = "overHeat") Then
          If (member.ivalue > 0) Then _overHeat = 1 Else _overHeat = 0
        ElseIf (member.name = "overCurrent") Then
          If (member.ivalue > 0) Then _overCurrent = 1 Else _overCurrent = 0
        ElseIf (member.name = "overLoad") Then
          If (member.ivalue > 0) Then _overLoad = 1 Else _overLoad = 0
        ElseIf (member.name = "regulationFailure") Then
          If (member.ivalue > 0) Then _regulationFailure = 1 Else _regulationFailure = 0
        ElseIf (member.name = "extPowerFailure") Then
          If (member.ivalue > 0) Then _extPowerFailure = 1 Else _extPowerFailure = 0
        ElseIf (member.name = "move") Then
          If (member.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then 
             _parse = -1
             Exit Function
          End If
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l=0 To member.membercount-1
             submemb = member.members(l)
             If (submemb.name = "moving") Then
                _move.moving = submemb.ivalue
             ElseIf (submemb.name = "target") Then
                _move.target = submemb.ivalue
             ElseIf (submemb.name = "ms") Then
                _move.ms = submemb.ivalue
             End If
          Next l
        ElseIf (member.name = "pulseTimer") Then
          If (member.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then 
             _parse = -1
             Exit Function
          End If
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l=0 To member.membercount-1
             submemb = member.members(l)
             If (submemb.name = "moving") Then
                _pulseTimer.moving = submemb.ivalue
             ElseIf (submemb.name = "target") Then
                _pulseTimer.target = submemb.ivalue
             ElseIf (submemb.name = "ms") Then
                _pulseTimer.ms = submemb.ivalue
             End If
          Next l
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the voltage source.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the voltage source
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
    '''   Changes the logical name of the voltage source.
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
    '''   a string corresponding to the logical name of the voltage source
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
    '''   Returns the current value of the voltage source (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the voltage source (no more than 6 characters)
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
    '''   Returns the measuring unit for the voltage.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the measuring unit for the voltage
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_UNIT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_unit() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_UNIT_INVALID
        End If
      End If
      Return _unit
    End Function

    '''*
    ''' <summary>
    '''   Returns the voltage output command (mV)
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the voltage output command (mV)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VOLTAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_voltage() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_VOLTAGE_INVALID
        End If
      End If
      Return CType(_voltage,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Tunes the device output voltage (milliVolts).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
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
    Public Function set_voltage(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("voltage", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns true if the  module is in failure mode.
    ''' <para>
    '''   More information can be obtained by testing
    '''   get_overheat, get_overcurrent etc... When a error condition is met, the output voltage is
    '''   set to zéro and cannot be changed until the reset() function is called.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_FAILURE_FALSE</c> or <c>Y_FAILURE_TRUE</c>, according to true if the  module is in failure mode
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_FAILURE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_failure() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_FAILURE_INVALID
        End If
      End If
      Return CType(_failure,Integer)
    End Function

    Public Function set_failure(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("failure", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns TRUE if the  module is overheating.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_OVERHEAT_FALSE</c> or <c>Y_OVERHEAT_TRUE</c>, according to TRUE if the  module is overheating
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OVERHEAT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_overHeat() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_OVERHEAT_INVALID
        End If
      End If
      Return CType(_overHeat,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns true if the appliance connected to the device is too greedy .
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_OVERCURRENT_FALSE</c> or <c>Y_OVERCURRENT_TRUE</c>, according to true if the appliance
    '''   connected to the device is too greedy
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OVERCURRENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_overCurrent() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_OVERCURRENT_INVALID
        End If
      End If
      Return CType(_overCurrent,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns true if the device is not able to maintaint the requested voltage output  .
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_OVERLOAD_FALSE</c> or <c>Y_OVERLOAD_TRUE</c>, according to true if the device is not
    '''   able to maintaint the requested voltage output
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OVERLOAD_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_overLoad() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_OVERLOAD_INVALID
        End If
      End If
      Return CType(_overLoad,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns true if the voltage output is too high regarding the requested voltage  .
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_REGULATIONFAILURE_FALSE</c> or <c>Y_REGULATIONFAILURE_TRUE</c>, according to true if
    '''   the voltage output is too high regarding the requested voltage
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_REGULATIONFAILURE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_regulationFailure() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_REGULATIONFAILURE_INVALID
        End If
      End If
      Return CType(_regulationFailure,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns true if external power supply voltage is too low.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_EXTPOWERFAILURE_FALSE</c> or <c>Y_EXTPOWERFAILURE_TRUE</c>, according to true if
    '''   external power supply voltage is too low
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_EXTPOWERFAILURE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_extPowerFailure() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_EXTPOWERFAILURE_INVALID
        End If
      End If
      Return CType(_extPowerFailure,Integer)
    End Function

    Public Function get_move() As YVSourceMove
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_MOVE_INVALID
        End If
      End If
      Return _move
    End Function

    Public Function set_move(ByVal newval As YVSourceMove) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target))+":"+Ltrim(Str(newval.ms))
      Return _setAttr("move", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth move at constant speed toward a given value.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   new output value at end of transition, in milliVolts.
    ''' </param>
    ''' <param name="ms_duration">
    '''   transition duration, in milliseconds
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
    Public Function voltageMove(ByVal target As Integer,ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(target))+":"+Ltrim(Str(ms_duration))
      Return _setAttr("move", rest_val)
    End Function

    Public Function get_pulseTimer() As YVSourcePulse
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PULSETIMER_INVALID
        End If
      End If
      Return _pulseTimer
    End Function

    Public Function set_pulseTimer(ByVal newval As YVSourcePulse) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target))+":"+Ltrim(Str(newval.ms))
      Return _setAttr("pulseTimer", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Sets device output to a specific volatage, for a specified duration, then brings it
    '''   automatically to 0V.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="voltage">
    '''   pulse voltage, in millivolts
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
    Public Function pulse(ByVal voltage As Integer,ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(voltage))+":"+Ltrim(Str(ms_duration))
      Return _setAttr("pulseTimer", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of voltage sources started using <c>yFirstVSource()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVSource</c> object, corresponding to
    '''   a voltage source currently online, or a <c>null</c> pointer
    '''   if there are no more voltage sources to enumerate.
    ''' </returns>
    '''/
    Public Function nextVSource() as YVSource
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindVSource(hwid)
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
    '''   Retrieves a voltage source for a given identifier.
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
    '''   This function does not require that the voltage source is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YVSource.isOnline()</c> to test if the voltage source is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a voltage source by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the voltage source
    ''' </param>
    ''' <returns>
    '''   a <c>YVSource</c> object allowing you to drive the voltage source.
    ''' </returns>
    '''/
    Public Shared Function FindVSource(ByVal func As String) As YVSource
      Dim res As YVSource
      If (_VSourceCache.ContainsKey(func)) Then
        Return CType(_VSourceCache(func), YVSource)
      End If
      res = New YVSource(func)
      _VSourceCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of voltage sources currently accessible.
    ''' <para>
    '''   Use the method <c>YVSource.nextVSource()</c> to iterate on
    '''   next voltage sources.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVSource</c> object, corresponding to
    '''   the first voltage source currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstVSource() As YVSource
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("VSource", 0, p, size, neededsize, errmsg)
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
      Return YVSource.FindVSource(serial + "." + funcId)
    End Function

    REM --- (end of YVSource implementation)

  End Class

  REM --- (VSource functions)

  '''*
  ''' <summary>
  '''   Retrieves a voltage source for a given identifier.
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
  '''   This function does not require that the voltage source is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YVSource.isOnline()</c> to test if the voltage source is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a voltage source by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the voltage source
  ''' </param>
  ''' <returns>
  '''   a <c>YVSource</c> object allowing you to drive the voltage source.
  ''' </returns>
  '''/
  Public Function yFindVSource(ByVal func As String) As YVSource
    Return YVSource.FindVSource(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of voltage sources currently accessible.
  ''' <para>
  '''   Use the method <c>YVSource.nextVSource()</c> to iterate on
  '''   next voltage sources.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YVSource</c> object, corresponding to
  '''   the first voltage source currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstVSource() As YVSource
    Return YVSource.FirstVSource()
  End Function

  Private Sub _VSourceCleanup()
  End Sub


  REM --- (end of VSource functions)

End Module
