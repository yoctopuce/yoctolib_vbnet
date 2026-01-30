' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindSoundLevel(), the high-level API for SoundLevel functions
'
'  - - - - - - - - - License information: - - - - - - - - -
'
'  Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'
'  Yoctopuce Sarl (hereafter Licensor) grants to you a perpetual
'  non-exclusive license to use, modify, copy and integrate this
'  file into your software for the sole purpose of interfacing
'  with Yoctopuce products.
'
'  You may reproduce and distribute copies of this file in
'  source or object form, as long as the sole purpose of this
'  code is to interface with Yoctopuce products. You must retain
'  this notice in the distributed source file.
'
'  You should refer to Yoctopuce General Terms and Conditions
'  for additional information regarding your rights and
'  obligations.
'
'  THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'  WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING
'  WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS
'  FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'  EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'  INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA,
'  COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR
'  SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT
'  LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'  CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'  BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'  WARRANTY, OR OTHERWISE.
'
' *********************************************************************


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_soundlevel

    REM --- (YSoundLevel return codes)
    REM --- (end of YSoundLevel return codes)
    REM --- (YSoundLevel dlldef)
    REM --- (end of YSoundLevel dlldef)
   REM --- (YSoundLevel yapiwrapper)
   REM --- (end of YSoundLevel yapiwrapper)
  REM --- (YSoundLevel globals)

  Public Const Y_LABEL_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_INTEGRATIONTIME_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YSoundLevelValueCallback(ByVal func As YSoundLevel, ByVal value As String)
  Public Delegate Sub YSoundLevelTimedReportCallback(ByVal func As YSoundLevel, ByVal measure As YMeasure)
  REM --- (end of YSoundLevel globals)

  REM --- (YSoundLevel class start)

  '''*
  ''' <summary>
  '''   The <c>YSoundLevel</c> class allows you to read and configure Yoctopuce sound pressure level meters.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSoundLevel
    Inherits YSensor
    REM --- (end of YSoundLevel class start)

    REM --- (YSoundLevel definitions)
    Public Const LABEL_INVALID As String = YAPI.INVALID_STRING
    Public Const INTEGRATIONTIME_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YSoundLevel definitions)

    REM --- (YSoundLevel attributes declaration)
    Protected _label As String
    Protected _integrationTime As Integer
    Protected _valueCallbackSoundLevel As YSoundLevelValueCallback
    Protected _timedReportCallbackSoundLevel As YSoundLevelTimedReportCallback
    REM --- (end of YSoundLevel attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "SoundLevel"
      REM --- (YSoundLevel attributes initialization)
      _label = LABEL_INVALID
      _integrationTime = INTEGRATIONTIME_INVALID
      _valueCallbackSoundLevel = Nothing
      _timedReportCallbackSoundLevel = Nothing
      REM --- (end of YSoundLevel attributes initialization)
    End Sub

    REM --- (YSoundLevel private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("label") Then
        _label = json_val.getString("label")
      End If
      If json_val.has("integrationTime") Then
        _integrationTime = CInt(json_val.getLong("integrationTime"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YSoundLevel private methods declaration)

    REM --- (YSoundLevel public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the measuring unit for the sound pressure level (dBA, dBC or dBZ).
    ''' <para>
    '''   That unit will directly determine frequency weighting to be used to compute
    '''   the measured value. Remember to call the <c>saveToFlash()</c> method of the
    '''   module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the measuring unit for the sound pressure level (dBA, dBC or dBZ)
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_unit(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("unit", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the label for the sound pressure level measurement, as per
    '''   IEC standard 61672-1:2013.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the label for the sound pressure level measurement, as per
    '''   IEC standard 61672-1:2013
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSoundLevel.LABEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_label() As String
      Dim res As String
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LABEL_INVALID
        End If
      End If
      res = Me._label
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the integration time in milliseconds for measuring the sound pressure level.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the integration time in milliseconds for measuring the sound pressure level
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSoundLevel.INTEGRATIONTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_integrationTime() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return INTEGRATIONTIME_INVALID
        End If
      End If
      res = Me._integrationTime
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a sound pressure level meter for a given identifier.
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
    '''   This function does not require that the sound pressure level meter is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSoundLevel.isOnline()</c> to test if the sound pressure level meter is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a sound pressure level meter by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the sound pressure level meter, for instance
    '''   <c>MyDevice.soundLevel1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YSoundLevel</c> object allowing you to drive the sound pressure level meter.
    ''' </returns>
    '''/
    Public Shared Function FindSoundLevel(func As String) As YSoundLevel
      Dim obj As YSoundLevel
      obj = CType(YFunction._FindFromCache("SoundLevel", func), YSoundLevel)
      If ((obj Is Nothing)) Then
        obj = New YSoundLevel(func)
        YFunction._AddToCache("SoundLevel", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YSoundLevelValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSoundLevel = callback
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
      If (Not (Me._valueCallbackSoundLevel Is Nothing)) Then
        Me._valueCallbackSoundLevel(Me, value)
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
    '''   arguments: the function object of which the value has changed, and an <c>YMeasure</c> object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YSoundLevelTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackSoundLevel = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackSoundLevel Is Nothing)) Then
        Me._timedReportCallbackSoundLevel(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of sound pressure level meters started using <c>yFirstSoundLevel()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned sound pressure level meters order.
    '''   If you want to find a specific a sound pressure level meter, use <c>SoundLevel.findSoundLevel()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSoundLevel</c> object, corresponding to
    '''   a sound pressure level meter currently online, or a <c>Nothing</c> pointer
    '''   if there are no more sound pressure level meters to enumerate.
    ''' </returns>
    '''/
    Public Function nextSoundLevel() As YSoundLevel
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSoundLevel.FindSoundLevel(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of sound pressure level meters currently accessible.
    ''' <para>
    '''   Use the method <c>YSoundLevel.nextSoundLevel()</c> to iterate on
    '''   next sound pressure level meters.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSoundLevel</c> object, corresponding to
    '''   the first sound pressure level meter currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstSoundLevel() As YSoundLevel
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("SoundLevel", 0, p, size, neededsize, errmsg)
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
      Return YSoundLevel.FindSoundLevel(serial + "." + funcId)
    End Function

    REM --- (end of YSoundLevel public methods declaration)

  End Class

  REM --- (YSoundLevel functions)

  '''*
  ''' <summary>
  '''   Retrieves a sound pressure level meter for a given identifier.
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
  '''   This function does not require that the sound pressure level meter is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSoundLevel.isOnline()</c> to test if the sound pressure level meter is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a sound pressure level meter by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the sound pressure level meter, for instance
  '''   <c>MyDevice.soundLevel1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YSoundLevel</c> object allowing you to drive the sound pressure level meter.
  ''' </returns>
  '''/
  Public Function yFindSoundLevel(ByVal func As String) As YSoundLevel
    Return YSoundLevel.FindSoundLevel(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of sound pressure level meters currently accessible.
  ''' <para>
  '''   Use the method <c>YSoundLevel.nextSoundLevel()</c> to iterate on
  '''   next sound pressure level meters.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YSoundLevel</c> object, corresponding to
  '''   the first sound pressure level meter currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSoundLevel() As YSoundLevel
    Return YSoundLevel.FirstSoundLevel()
  End Function


  REM --- (end of YSoundLevel functions)

End Module
