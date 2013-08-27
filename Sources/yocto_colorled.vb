'*********************************************************************
'*
'* $Id: yocto_colorled.vb 12324 2013-08-13 15:10:31Z mvuilleu $
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

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YColorLed definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YColorLed, ByVal value As String)

Public Class YColorLedMove
  Public target As System.Int64 = YAPI.INVALID_LONG
  Public ms As System.Int64 = YAPI.INVALID_LONG
  Public moving As System.Int64 = YAPI.INVALID_LONG
End Class


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_RGBCOLOR_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_HSLCOLOR_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_RGBCOLORATPOWERON_INVALID As Long = YAPI.INVALID_LONG

  Public Y_RGBMOVE_INVALID As YColorLedMove
  Public Y_HSLMOVE_INVALID As YColorLedMove

  REM --- (end of YColorLed definitions)

  REM --- (YColorLed implementation)

  Private _ColorLedCache As New Hashtable()
  Private _callback As UpdateCallback

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
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const RGBCOLOR_INVALID As Long = YAPI.INVALID_LONG
    Public Const HSLCOLOR_INVALID As Long = YAPI.INVALID_LONG
    Public Const RGBCOLORATPOWERON_INVALID As Long = YAPI.INVALID_LONG

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _rgbColor As Long
    Protected _hslColor As Long
    Protected _rgbMove As YColorLedMove
    Protected _hslMove As YColorLedMove
    Protected _rgbColorAtPowerOn As Long

    Public Sub New(ByVal func As String)
      MyBase.new("ColorLed", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _rgbColor = Y_RGBCOLOR_INVALID
      _hslColor = Y_HSLCOLOR_INVALID
      _rgbMove = New YColorLedMove()
      _hslMove = New YColorLedMove()
      _rgbColorAtPowerOn = Y_RGBCOLORATPOWERON_INVALID
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
        ElseIf (member.name = "rgbColor") Then
          _rgbColor = CLng(member.ivalue)
        ElseIf (member.name = "hslColor") Then
          _hslColor = CLng(member.ivalue)
        ElseIf (member.name = "rgbMove") Then
          If (member.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then 
             _parse = -1
             Exit Function
          End If
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l=0 To member.membercount-1
             submemb = member.members(l)
             If (submemb.name = "moving") Then
                _rgbMove.moving = submemb.ivalue
             ElseIf (submemb.name = "target") Then
                _rgbMove.target = submemb.ivalue
             ElseIf (submemb.name = "ms") Then
                _rgbMove.ms = submemb.ivalue
             End If
          Next l
        ElseIf (member.name = "hslMove") Then
          If (member.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then 
             _parse = -1
             Exit Function
          End If
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l=0 To member.membercount-1
             submemb = member.members(l)
             If (submemb.name = "moving") Then
                _hslMove.moving = submemb.ivalue
             ElseIf (submemb.name = "target") Then
                _hslMove.target = submemb.ivalue
             ElseIf (submemb.name = "ms") Then
                _hslMove.ms = submemb.ivalue
             End If
          Next l
        ElseIf (member.name = "rgbColorAtPowerOn") Then
          _rgbColorAtPowerOn = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the RGB led.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the RGB led
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
    '''   Changes the logical name of the RGB led.
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
    '''   a string corresponding to the logical name of the RGB led
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
    '''   Returns the current value of the RGB led (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the RGB led (no more than 6 characters)
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
    Public Function get_rgbColor() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_RGBCOLOR_INVALID
        End If
      End If
      Return _rgbColor
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
    Public Function set_rgbColor(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = "0x"+Hex(newval)
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
    Public Function get_hslColor() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_HSLCOLOR_INVALID
        End If
      End If
      Return _hslColor
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
    Public Function set_hslColor(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = "0x"+Hex(newval)
      Return _setAttr("hslColor", rest_val)
    End Function

    Public Function get_rgbMove() As YColorLedMove
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_RGBMOVE_INVALID
        End If
      End If
      Return _rgbMove
    End Function

    Public Function set_rgbMove(ByVal newval As YColorLedMove) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target))+":"+Ltrim(Str(newval.ms))
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
    Public Function rgbMove(ByVal rgb_target As Integer,ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(rgb_target))+":"+Ltrim(Str(ms_duration))
      Return _setAttr("rgbMove", rest_val)
    End Function

    Public Function get_hslMove() As YColorLedMove
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_HSLMOVE_INVALID
        End If
      End If
      Return _hslMove
    End Function

    Public Function set_hslMove(ByVal newval As YColorLedMove) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target))+":"+Ltrim(Str(newval.ms))
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
    Public Function hslMove(ByVal hsl_target As Integer,ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(hsl_target))+":"+Ltrim(Str(ms_duration))
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
    Public Function get_rgbColorAtPowerOn() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_RGBCOLORATPOWERON_INVALID
        End If
      End If
      Return _rgbColorAtPowerOn
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
    Public Function set_rgbColorAtPowerOn(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = "0x"+Hex(newval)
      Return _setAttr("rgbColorAtPowerOn", rest_val)
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
    Public Function nextColorLed() as YColorLed
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindColorLed(hwid)
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
    Public Shared Function FindColorLed(ByVal func As String) As YColorLed
      Dim res As YColorLed
      If (_ColorLedCache.ContainsKey(func)) Then
        Return CType(_ColorLedCache(func), YColorLed)
      End If
      res = New YColorLed(func)
      _ColorLedCache.Add(func, res)
      Return res
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

    REM --- (end of YColorLed implementation)

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

  Private Sub _ColorLedCleanup()
  End Sub


  REM --- (end of ColorLed functions)

End Module
