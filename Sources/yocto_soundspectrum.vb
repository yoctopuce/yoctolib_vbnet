' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindSoundSpectrum(), the high-level API for SoundSpectrum functions
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

Module yocto_soundspectrum

    REM --- (YSoundSpectrum return codes)
    REM --- (end of YSoundSpectrum return codes)
    REM --- (YSoundSpectrum dlldef)
    REM --- (end of YSoundSpectrum dlldef)
   REM --- (YSoundSpectrum yapiwrapper)
   REM --- (end of YSoundSpectrum yapiwrapper)
  REM --- (YSoundSpectrum globals)

  Public Const Y_INTEGRATIONTIME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_SPECTRUMDATA_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YSoundSpectrumValueCallback(ByVal func As YSoundSpectrum, ByVal value As String)
  Public Delegate Sub YSoundSpectrumTimedReportCallback(ByVal func As YSoundSpectrum, ByVal measure As YMeasure)
  REM --- (end of YSoundSpectrum globals)

  REM --- (YSoundSpectrum class start)

  '''*
  ''' <summary>
  '''   The <c>YSoundSpectrum</c> class allows you to read and configure Yoctopuce sound spectrum analyzers.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSoundSpectrum
    Inherits YFunction
    REM --- (end of YSoundSpectrum class start)

    REM --- (YSoundSpectrum definitions)
    Public Const INTEGRATIONTIME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const SPECTRUMDATA_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YSoundSpectrum definitions)

    REM --- (YSoundSpectrum attributes declaration)
    Protected _integrationTime As Integer
    Protected _spectrumData As String
    Protected _valueCallbackSoundSpectrum As YSoundSpectrumValueCallback
    REM --- (end of YSoundSpectrum attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "SoundSpectrum"
      REM --- (YSoundSpectrum attributes initialization)
      _integrationTime = INTEGRATIONTIME_INVALID
      _spectrumData = SPECTRUMDATA_INVALID
      _valueCallbackSoundSpectrum = Nothing
      REM --- (end of YSoundSpectrum attributes initialization)
    End Sub

    REM --- (YSoundSpectrum private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("integrationTime") Then
        _integrationTime = CInt(json_val.getLong("integrationTime"))
      End If
      If json_val.has("spectrumData") Then
        _spectrumData = json_val.getString("spectrumData")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YSoundSpectrum private methods declaration)

    REM --- (YSoundSpectrum public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the integration time in milliseconds for calculating time
    '''   weighted spectrum data.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the integration time in milliseconds for calculating time
    '''   weighted spectrum data
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSoundSpectrum.INTEGRATIONTIME_INVALID</c>.
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
    '''   Changes the integration time in milliseconds for computing time weighted
    '''   spectrum data.
    ''' <para>
    '''   Be aware that on some devices, changing the integration
    '''   time for time-weighted spectrum data may also affect the integration
    '''   period for one or more sound pressure level measurements.
    '''   Remember to call the <c>saveToFlash()</c> method of the
    '''   module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the integration time in milliseconds for computing time weighted
    '''   spectrum data
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
    Public Function set_integrationTime(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("integrationTime", rest_val)
    End Function
    Public Function get_spectrumData() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SPECTRUMDATA_INVALID
        End If
      End If
      res = Me._spectrumData
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a sound spectrum analyzer for a given identifier.
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
    '''   This function does not require that the sound spectrum analyzer is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSoundSpectrum.isOnline()</c> to test if the sound spectrum analyzer is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a sound spectrum analyzer by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the sound spectrum analyzer, for instance
    '''   <c>MyDevice.soundSpectrum</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YSoundSpectrum</c> object allowing you to drive the sound spectrum analyzer.
    ''' </returns>
    '''/
    Public Shared Function FindSoundSpectrum(func As String) As YSoundSpectrum
      Dim obj As YSoundSpectrum
      obj = CType(YFunction._FindFromCache("SoundSpectrum", func), YSoundSpectrum)
      If ((obj Is Nothing)) Then
        obj = New YSoundSpectrum(func)
        YFunction._AddToCache("SoundSpectrum", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YSoundSpectrumValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSoundSpectrum = callback
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
      If (Not (Me._valueCallbackSoundSpectrum Is Nothing)) Then
        Me._valueCallbackSoundSpectrum(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   c
    ''' <para>
    '''   omment from .yc definition
    ''' </para>
    ''' </summary>
    '''/
    Public Function nextSoundSpectrum() As YSoundSpectrum
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSoundSpectrum.FindSoundSpectrum(hwid)
    End Function

    '''*
    ''' <summary>
    '''   c
    ''' <para>
    '''   omment from .yc definition
    ''' </para>
    ''' </summary>
    '''/
    Public Shared Function FirstSoundSpectrum() As YSoundSpectrum
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("SoundSpectrum", 0, p, size, neededsize, errmsg)
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
      Return YSoundSpectrum.FindSoundSpectrum(serial + "." + funcId)
    End Function

    REM --- (end of YSoundSpectrum public methods declaration)

  End Class

  REM --- (YSoundSpectrum functions)

  '''*
  ''' <summary>
  '''   Retrieves a sound spectrum analyzer for a given identifier.
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
  '''   This function does not require that the sound spectrum analyzer is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSoundSpectrum.isOnline()</c> to test if the sound spectrum analyzer is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a sound spectrum analyzer by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the sound spectrum analyzer, for instance
  '''   <c>MyDevice.soundSpectrum</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YSoundSpectrum</c> object allowing you to drive the sound spectrum analyzer.
  ''' </returns>
  '''/
  Public Function yFindSoundSpectrum(ByVal func As String) As YSoundSpectrum
    Return YSoundSpectrum.FindSoundSpectrum(func)
  End Function

  '''*
  ''' <summary>
  '''   A
  ''' <para>
  '''   lias for Y{$classname}.First{$classname}()
  ''' </para>
  ''' </summary>
  '''/
  Public Function yFirstSoundSpectrum() As YSoundSpectrum
    Return YSoundSpectrum.FirstSoundSpectrum()
  End Function


  REM --- (end of YSoundSpectrum functions)

End Module
