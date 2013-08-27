'*********************************************************************
'*
'* $Id: yocto_led.vb 12324 2013-08-13 15:10:31Z mvuilleu $
'*
'* Implements yFindLed(), the high-level API for Led functions
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

Module yocto_led

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YLed definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YLed, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_POWER_OFF = 0
  Public Const Y_POWER_ON = 1
  Public Const Y_POWER_INVALID = -1

  Public Const Y_LUMINOSITY_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_BLINKING_STILL = 0
  Public Const Y_BLINKING_RELAX = 1
  Public Const Y_BLINKING_AWARE = 2
  Public Const Y_BLINKING_RUN = 3
  Public Const Y_BLINKING_CALL = 4
  Public Const Y_BLINKING_PANIC = 5
  Public Const Y_BLINKING_INVALID = -1



  REM --- (end of YLed definitions)

  REM --- (YLed implementation)

  Private _LedCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   Yoctopuce application programming interface
  '''   allows you not only to drive the intensity of the led, but also to
  '''   have it blink at various preset frequencies.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YLed
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const POWER_OFF = 0
    Public Const POWER_ON = 1
    Public Const POWER_INVALID = -1

    Public Const LUMINOSITY_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const BLINKING_STILL = 0
    Public Const BLINKING_RELAX = 1
    Public Const BLINKING_AWARE = 2
    Public Const BLINKING_RUN = 3
    Public Const BLINKING_CALL = 4
    Public Const BLINKING_PANIC = 5
    Public Const BLINKING_INVALID = -1


    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _power As Long
    Protected _luminosity As Long
    Protected _blinking As Long

    Public Sub New(ByVal func As String)
      MyBase.new("Led", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _power = Y_POWER_INVALID
      _luminosity = Y_LUMINOSITY_INVALID
      _blinking = Y_BLINKING_INVALID
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
        ElseIf (member.name = "power") Then
          If (member.ivalue > 0) Then _power = 1 Else _power = 0
        ElseIf (member.name = "luminosity") Then
          _luminosity = CLng(member.ivalue)
        ElseIf (member.name = "blinking") Then
          _blinking = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the led.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the led
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
    '''   Changes the logical name of the led.
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
    '''   a string corresponding to the logical name of the led
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
    '''   Returns the current value of the led (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the led (no more than 6 characters)
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
    '''   Returns the current led state.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_POWER_OFF</c> or <c>Y_POWER_ON</c>, according to the current led state
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_POWER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_power() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_POWER_INVALID
        End If
      End If
      Return CType(_power,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the state of the led.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_POWER_OFF</c> or <c>Y_POWER_ON</c>, according to the state of the led
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
    Public Function set_power(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("power", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current led intensity (in per cent).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current led intensity (in per cent)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LUMINOSITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_luminosity() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LUMINOSITY_INVALID
        End If
      End If
      Return CType(_luminosity,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the current led intensity (in per cent).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current led intensity (in per cent)
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
    Public Function set_luminosity(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("luminosity", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current led signaling mode.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_BLINKING_STILL</c>, <c>Y_BLINKING_RELAX</c>, <c>Y_BLINKING_AWARE</c>,
    '''   <c>Y_BLINKING_RUN</c>, <c>Y_BLINKING_CALL</c> and <c>Y_BLINKING_PANIC</c> corresponding to the
    '''   current led signaling mode
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BLINKING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_blinking() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_BLINKING_INVALID
        End If
      End If
      Return CType(_blinking,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the current led signaling mode.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_BLINKING_STILL</c>, <c>Y_BLINKING_RELAX</c>, <c>Y_BLINKING_AWARE</c>,
    '''   <c>Y_BLINKING_RUN</c>, <c>Y_BLINKING_CALL</c> and <c>Y_BLINKING_PANIC</c> corresponding to the
    '''   current led signaling mode
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
    Public Function set_blinking(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("blinking", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of leds started using <c>yFirstLed()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLed</c> object, corresponding to
    '''   a led currently online, or a <c>null</c> pointer
    '''   if there are no more leds to enumerate.
    ''' </returns>
    '''/
    Public Function nextLed() as YLed
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindLed(hwid)
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
    '''   Retrieves a led for a given identifier.
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
    '''   This function does not require that the led is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YLed.isOnline()</c> to test if the led is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a led by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the led
    ''' </param>
    ''' <returns>
    '''   a <c>YLed</c> object allowing you to drive the led.
    ''' </returns>
    '''/
    Public Shared Function FindLed(ByVal func As String) As YLed
      Dim res As YLed
      If (_LedCache.ContainsKey(func)) Then
        Return CType(_LedCache(func), YLed)
      End If
      res = New YLed(func)
      _LedCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of leds currently accessible.
    ''' <para>
    '''   Use the method <c>YLed.nextLed()</c> to iterate on
    '''   next leds.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YLed</c> object, corresponding to
    '''   the first led currently online, or a <c>null</c> pointer
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

    REM --- (end of YLed implementation)

  End Class

  REM --- (Led functions)

  '''*
  ''' <summary>
  '''   Retrieves a led for a given identifier.
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
  '''   This function does not require that the led is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YLed.isOnline()</c> to test if the led is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a led by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the led
  ''' </param>
  ''' <returns>
  '''   a <c>YLed</c> object allowing you to drive the led.
  ''' </returns>
  '''/
  Public Function yFindLed(ByVal func As String) As YLed
    Return YLed.FindLed(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of leds currently accessible.
  ''' <para>
  '''   Use the method <c>YLed.nextLed()</c> to iterate on
  '''   next leds.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YLed</c> object, corresponding to
  '''   the first led currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstLed() As YLed
    Return YLed.FirstLed()
  End Function

  Private Sub _LedCleanup()
  End Sub


  REM --- (end of Led functions)

End Module
