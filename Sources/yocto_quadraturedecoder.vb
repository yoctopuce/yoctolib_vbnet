'*********************************************************************
'*
'* $Id: yocto_quadraturedecoder.vb 28740 2017-10-03 08:09:13Z seb $
'*
'* Implements yFindQuadratureDecoder(), the high-level API for QuadratureDecoder functions
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

Module yocto_quadraturedecoder

    REM --- (YQuadratureDecoder return codes)
    REM --- (end of YQuadratureDecoder return codes)
    REM --- (YQuadratureDecoder dlldef)
    REM --- (end of YQuadratureDecoder dlldef)
  REM --- (YQuadratureDecoder globals)

  Public Const Y_SPEED_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_DECODING_OFF As Integer = 0
  Public Const Y_DECODING_ON As Integer = 1
  Public Const Y_DECODING_INVALID As Integer = -1
  Public Delegate Sub YQuadratureDecoderValueCallback(ByVal func As YQuadratureDecoder, ByVal value As String)
  Public Delegate Sub YQuadratureDecoderTimedReportCallback(ByVal func As YQuadratureDecoder, ByVal measure As YMeasure)
  REM --- (end of YQuadratureDecoder globals)

  REM --- (YQuadratureDecoder class start)

  '''*
  ''' <summary>
  '''   The class YQuadratureDecoder allows you to decode a two-wire signal produced by a
  '''   quadrature encoder.
  ''' <para>
  '''   It inherits from YSensor class the core functions to read measurements,
  '''   to register callback functions, to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YQuadratureDecoder
    Inherits YSensor
    REM --- (end of YQuadratureDecoder class start)

    REM --- (YQuadratureDecoder definitions)
    Public Const SPEED_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const DECODING_OFF As Integer = 0
    Public Const DECODING_ON As Integer = 1
    Public Const DECODING_INVALID As Integer = -1
    REM --- (end of YQuadratureDecoder definitions)

    REM --- (YQuadratureDecoder attributes declaration)
    Protected _speed As Double
    Protected _decoding As Integer
    Protected _valueCallbackQuadratureDecoder As YQuadratureDecoderValueCallback
    Protected _timedReportCallbackQuadratureDecoder As YQuadratureDecoderTimedReportCallback
    REM --- (end of YQuadratureDecoder attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "QuadratureDecoder"
      REM --- (YQuadratureDecoder attributes initialization)
      _speed = SPEED_INVALID
      _decoding = DECODING_INVALID
      _valueCallbackQuadratureDecoder = Nothing
      _timedReportCallbackQuadratureDecoder = Nothing
      REM --- (end of YQuadratureDecoder attributes initialization)
    End Sub

    REM --- (YQuadratureDecoder private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("speed") Then
        _speed = Math.Round(json_val.getDouble("speed") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("decoding") Then
        If (json_val.getInt("decoding") > 0) Then _decoding = 1 Else _decoding = 0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YQuadratureDecoder private methods declaration)

    REM --- (YQuadratureDecoder public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the current expected position of the quadrature decoder.
    ''' <para>
    '''   Invoking this function implicitely activates the quadrature decoder.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the current expected position of the quadrature decoder
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
    Public Function set_currentValue(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("currentValue", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the increments frequency, in Hz.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the increments frequency, in Hz
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SPEED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_speed() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SPEED_INVALID
        End If
      End If
      res = Me._speed
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current activation state of the quadrature decoder.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_DECODING_OFF</c> or <c>Y_DECODING_ON</c>, according to the current activation state of
    '''   the quadrature decoder
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DECODING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_decoding() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return DECODING_INVALID
        End If
      End If
      res = Me._decoding
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the activation state of the quadrature decoder.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_DECODING_OFF</c> or <c>Y_DECODING_ON</c>, according to the activation state of the
    '''   quadrature decoder
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
    Public Function set_decoding(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("decoding", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a quadrature decoder for a given identifier.
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
    '''   This function does not require that the quadrature decoder is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YQuadratureDecoder.isOnline()</c> to test if the quadrature decoder is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a quadrature decoder by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the quadrature decoder
    ''' </param>
    ''' <returns>
    '''   a <c>YQuadratureDecoder</c> object allowing you to drive the quadrature decoder.
    ''' </returns>
    '''/
    Public Shared Function FindQuadratureDecoder(func As String) As YQuadratureDecoder
      Dim obj As YQuadratureDecoder
      obj = CType(YFunction._FindFromCache("QuadratureDecoder", func), YQuadratureDecoder)
      If ((obj Is Nothing)) Then
        obj = New YQuadratureDecoder(func)
        YFunction._AddToCache("QuadratureDecoder", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YQuadratureDecoderValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackQuadratureDecoder = callback
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
      If (Not (Me._valueCallbackQuadratureDecoder Is Nothing)) Then
        Me._valueCallbackQuadratureDecoder(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every periodic timed notification.
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
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YQuadratureDecoderTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackQuadratureDecoder = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackQuadratureDecoder Is Nothing)) Then
        Me._timedReportCallbackQuadratureDecoder(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of quadrature decoders started using <c>yFirstQuadratureDecoder()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YQuadratureDecoder</c> object, corresponding to
    '''   a quadrature decoder currently online, or a <c>Nothing</c> pointer
    '''   if there are no more quadrature decoders to enumerate.
    ''' </returns>
    '''/
    Public Function nextQuadratureDecoder() As YQuadratureDecoder
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YQuadratureDecoder.FindQuadratureDecoder(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of quadrature decoders currently accessible.
    ''' <para>
    '''   Use the method <c>YQuadratureDecoder.nextQuadratureDecoder()</c> to iterate on
    '''   next quadrature decoders.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YQuadratureDecoder</c> object, corresponding to
    '''   the first quadrature decoder currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstQuadratureDecoder() As YQuadratureDecoder
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("QuadratureDecoder", 0, p, size, neededsize, errmsg)
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
      Return YQuadratureDecoder.FindQuadratureDecoder(serial + "." + funcId)
    End Function

    REM --- (end of YQuadratureDecoder public methods declaration)

  End Class

  REM --- (YQuadratureDecoder functions)

  '''*
  ''' <summary>
  '''   Retrieves a quadrature decoder for a given identifier.
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
  '''   This function does not require that the quadrature decoder is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YQuadratureDecoder.isOnline()</c> to test if the quadrature decoder is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a quadrature decoder by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the quadrature decoder
  ''' </param>
  ''' <returns>
  '''   a <c>YQuadratureDecoder</c> object allowing you to drive the quadrature decoder.
  ''' </returns>
  '''/
  Public Function yFindQuadratureDecoder(ByVal func As String) As YQuadratureDecoder
    Return YQuadratureDecoder.FindQuadratureDecoder(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of quadrature decoders currently accessible.
  ''' <para>
  '''   Use the method <c>YQuadratureDecoder.nextQuadratureDecoder()</c> to iterate on
  '''   next quadrature decoders.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YQuadratureDecoder</c> object, corresponding to
  '''   the first quadrature decoder currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstQuadratureDecoder() As YQuadratureDecoder
    Return YQuadratureDecoder.FirstQuadratureDecoder()
  End Function


  REM --- (end of YQuadratureDecoder functions)

End Module
