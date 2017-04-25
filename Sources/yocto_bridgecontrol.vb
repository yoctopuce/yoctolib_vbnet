'*********************************************************************
'*
'* $Id: yocto_bridgecontrol.vb 27237 2017-04-21 16:36:03Z seb $
'*
'* Implements yFindBridgeControl(), the high-level API for BridgeControl functions
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

Module yocto_bridgecontrol

    REM --- (YBridgeControl return codes)
    REM --- (end of YBridgeControl return codes)
    REM --- (YBridgeControl dlldef)
    REM --- (end of YBridgeControl dlldef)
  REM --- (YBridgeControl globals)

  Public Const Y_EXCITATIONMODE_INTERNAL_AC As Integer = 0
  Public Const Y_EXCITATIONMODE_INTERNAL_DC As Integer = 1
  Public Const Y_EXCITATIONMODE_EXTERNAL_DC As Integer = 2
  Public Const Y_EXCITATIONMODE_INVALID As Integer = -1
  Public Const Y_BRIDGELATENCY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ADVALUE_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_ADGAIN_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YBridgeControlValueCallback(ByVal func As YBridgeControl, ByVal value As String)
  Public Delegate Sub YBridgeControlTimedReportCallback(ByVal func As YBridgeControl, ByVal measure As YMeasure)
  REM --- (end of YBridgeControl globals)

  REM --- (YBridgeControl class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce class YBridgeControl allows you to control bridge excitation parameters
  '''   and measure parameters for a Wheatstone bridge sensor.
  ''' <para>
  '''   To read the measurements, it
  '''   is best to use the GenericSensor calss, which will compute the measured value
  '''   in the optimal way.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YBridgeControl
    Inherits YFunction
    REM --- (end of YBridgeControl class start)

    REM --- (YBridgeControl definitions)
    Public Const EXCITATIONMODE_INTERNAL_AC As Integer = 0
    Public Const EXCITATIONMODE_INTERNAL_DC As Integer = 1
    Public Const EXCITATIONMODE_EXTERNAL_DC As Integer = 2
    Public Const EXCITATIONMODE_INVALID As Integer = -1
    Public Const BRIDGELATENCY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ADVALUE_INVALID As Integer = YAPI.INVALID_INT
    Public Const ADGAIN_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YBridgeControl definitions)

    REM --- (YBridgeControl attributes declaration)
    Protected _excitationMode As Integer
    Protected _bridgeLatency As Integer
    Protected _adValue As Integer
    Protected _adGain As Integer
    Protected _valueCallbackBridgeControl As YBridgeControlValueCallback
    REM --- (end of YBridgeControl attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "BridgeControl"
      REM --- (YBridgeControl attributes initialization)
      _excitationMode = EXCITATIONMODE_INVALID
      _bridgeLatency = BRIDGELATENCY_INVALID
      _adValue = ADVALUE_INVALID
      _adGain = ADGAIN_INVALID
      _valueCallbackBridgeControl = Nothing
      REM --- (end of YBridgeControl attributes initialization)
    End Sub

    REM --- (YBridgeControl private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("excitationMode") Then
        _excitationMode = CInt(json_val.getLong("excitationMode"))
      End If
      If json_val.has("bridgeLatency") Then
        _bridgeLatency = CInt(json_val.getLong("bridgeLatency"))
      End If
      If json_val.has("adValue") Then
        _adValue = CInt(json_val.getLong("adValue"))
      End If
      If json_val.has("adGain") Then
        _adGain = CInt(json_val.getLong("adGain"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YBridgeControl private methods declaration)

    REM --- (YBridgeControl public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current Wheatstone bridge excitation method.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_EXCITATIONMODE_INTERNAL_AC</c>, <c>Y_EXCITATIONMODE_INTERNAL_DC</c> and
    '''   <c>Y_EXCITATIONMODE_EXTERNAL_DC</c> corresponding to the current Wheatstone bridge excitation method
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_EXCITATIONMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_excitationMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return EXCITATIONMODE_INVALID
        End If
      End If
      res = Me._excitationMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current Wheatstone bridge excitation method.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_EXCITATIONMODE_INTERNAL_AC</c>, <c>Y_EXCITATIONMODE_INTERNAL_DC</c> and
    '''   <c>Y_EXCITATIONMODE_EXTERNAL_DC</c> corresponding to the current Wheatstone bridge excitation method
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
    Public Function set_excitationMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("excitationMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current Wheatstone bridge excitation method.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current Wheatstone bridge excitation method
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BRIDGELATENCY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bridgeLatency() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BRIDGELATENCY_INVALID
        End If
      End If
      res = Me._bridgeLatency
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current Wheatstone bridge excitation method.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current Wheatstone bridge excitation method
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
    Public Function set_bridgeLatency(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("bridgeLatency", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the raw value returned by the ratiometric A/D converter
    '''   during last read.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the raw value returned by the ratiometric A/D converter
    '''   during last read
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_adValue() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ADVALUE_INVALID
        End If
      End If
      res = Me._adValue
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current ratiometric A/D converter gain.
    ''' <para>
    '''   The gain is automatically
    '''   configured according to the signalRange set in the corresponding genericSensor.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current ratiometric A/D converter gain
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADGAIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_adGain() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ADGAIN_INVALID
        End If
      End If
      res = Me._adGain
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a Wheatstone bridge controller for a given identifier.
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
    '''   This function does not require that the Wheatstone bridge controller is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YBridgeControl.isOnline()</c> to test if the Wheatstone bridge controller is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a Wheatstone bridge controller by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the Wheatstone bridge controller
    ''' </param>
    ''' <returns>
    '''   a <c>YBridgeControl</c> object allowing you to drive the Wheatstone bridge controller.
    ''' </returns>
    '''/
    Public Shared Function FindBridgeControl(func As String) As YBridgeControl
      Dim obj As YBridgeControl
      obj = CType(YFunction._FindFromCache("BridgeControl", func), YBridgeControl)
      If ((obj Is Nothing)) Then
        obj = New YBridgeControl(func)
        YFunction._AddToCache("BridgeControl", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YBridgeControlValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackBridgeControl = callback
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
      If (Not (Me._valueCallbackBridgeControl Is Nothing)) Then
        Me._valueCallbackBridgeControl(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of Wheatstone bridge controllers started using <c>yFirstBridgeControl()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YBridgeControl</c> object, corresponding to
    '''   a Wheatstone bridge controller currently online, or a <c>Nothing</c> pointer
    '''   if there are no more Wheatstone bridge controllers to enumerate.
    ''' </returns>
    '''/
    Public Function nextBridgeControl() As YBridgeControl
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YBridgeControl.FindBridgeControl(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of Wheatstone bridge controllers currently accessible.
    ''' <para>
    '''   Use the method <c>YBridgeControl.nextBridgeControl()</c> to iterate on
    '''   next Wheatstone bridge controllers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YBridgeControl</c> object, corresponding to
    '''   the first Wheatstone bridge controller currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstBridgeControl() As YBridgeControl
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("BridgeControl", 0, p, size, neededsize, errmsg)
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
      Return YBridgeControl.FindBridgeControl(serial + "." + funcId)
    End Function

    REM --- (end of YBridgeControl public methods declaration)

  End Class

  REM --- (BridgeControl functions)

  '''*
  ''' <summary>
  '''   Retrieves a Wheatstone bridge controller for a given identifier.
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
  '''   This function does not require that the Wheatstone bridge controller is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YBridgeControl.isOnline()</c> to test if the Wheatstone bridge controller is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a Wheatstone bridge controller by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the Wheatstone bridge controller
  ''' </param>
  ''' <returns>
  '''   a <c>YBridgeControl</c> object allowing you to drive the Wheatstone bridge controller.
  ''' </returns>
  '''/
  Public Function yFindBridgeControl(ByVal func As String) As YBridgeControl
    Return YBridgeControl.FindBridgeControl(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of Wheatstone bridge controllers currently accessible.
  ''' <para>
  '''   Use the method <c>YBridgeControl.nextBridgeControl()</c> to iterate on
  '''   next Wheatstone bridge controllers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YBridgeControl</c> object, corresponding to
  '''   the first Wheatstone bridge controller currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstBridgeControl() As YBridgeControl
    Return YBridgeControl.FirstBridgeControl()
  End Function


  REM --- (end of BridgeControl functions)

End Module
