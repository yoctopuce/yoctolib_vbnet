'*********************************************************************
'*
'* $Id: yocto_oscontrol.vb 28740 2017-10-03 08:09:13Z seb $
'*
'* Implements yFindOsControl(), the high-level API for OsControl functions
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

Module yocto_oscontrol

    REM --- (YOsControl return codes)
    REM --- (end of YOsControl return codes)
    REM --- (YOsControl dlldef)
    REM --- (end of YOsControl dlldef)
  REM --- (YOsControl globals)

  Public Const Y_SHUTDOWNCOUNTDOWN_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YOsControlValueCallback(ByVal func As YOsControl, ByVal value As String)
  Public Delegate Sub YOsControlTimedReportCallback(ByVal func As YOsControl, ByVal measure As YMeasure)
  REM --- (end of YOsControl globals)

  REM --- (YOsControl class start)

  '''*
  ''' <summary>
  '''   The OScontrol object allows some control over the operating system running a VirtualHub.
  ''' <para>
  '''   OsControl is available on the VirtualHub software only. This feature must be activated at the VirtualHub
  '''   start up with -o option.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YOsControl
    Inherits YFunction
    REM --- (end of YOsControl class start)

    REM --- (YOsControl definitions)
    Public Const SHUTDOWNCOUNTDOWN_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YOsControl definitions)

    REM --- (YOsControl attributes declaration)
    Protected _shutdownCountdown As Integer
    Protected _valueCallbackOsControl As YOsControlValueCallback
    REM --- (end of YOsControl attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "OsControl"
      REM --- (YOsControl attributes initialization)
      _shutdownCountdown = SHUTDOWNCOUNTDOWN_INVALID
      _valueCallbackOsControl = Nothing
      REM --- (end of YOsControl attributes initialization)
    End Sub

    REM --- (YOsControl private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("shutdownCountdown") Then
        _shutdownCountdown = CInt(json_val.getLong("shutdownCountdown"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YOsControl private methods declaration)

    REM --- (YOsControl public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the remaining number of seconds before the OS shutdown, or zero when no
    '''   shutdown has been scheduled.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the remaining number of seconds before the OS shutdown, or zero when no
    '''   shutdown has been scheduled
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SHUTDOWNCOUNTDOWN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_shutdownCountdown() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SHUTDOWNCOUNTDOWN_INVALID
        End If
      End If
      res = Me._shutdownCountdown
      Return res
    End Function


    Public Function set_shutdownCountdown(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("shutdownCountdown", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves OS control for a given identifier.
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
    '''   This function does not require that the OS control is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YOsControl.isOnline()</c> to test if the OS control is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   OS control by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the OS control
    ''' </param>
    ''' <returns>
    '''   a <c>YOsControl</c> object allowing you to drive the OS control.
    ''' </returns>
    '''/
    Public Shared Function FindOsControl(func As String) As YOsControl
      Dim obj As YOsControl
      obj = CType(YFunction._FindFromCache("OsControl", func), YOsControl)
      If ((obj Is Nothing)) Then
        obj = New YOsControl(func)
        YFunction._AddToCache("OsControl", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YOsControlValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackOsControl = callback
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
      If (Not (Me._valueCallbackOsControl Is Nothing)) Then
        Me._valueCallbackOsControl(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Schedules an OS shutdown after a given number of seconds.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="secBeforeShutDown">
    '''   number of seconds before shutdown
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function shutdown(secBeforeShutDown As Integer) As Integer
      Return Me.set_shutdownCountdown(secBeforeShutDown)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of OS control started using <c>yFirstOsControl()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YOsControl</c> object, corresponding to
    '''   OS control currently online, or a <c>Nothing</c> pointer
    '''   if there are no more OS control to enumerate.
    ''' </returns>
    '''/
    Public Function nextOsControl() As YOsControl
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YOsControl.FindOsControl(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of OS control currently accessible.
    ''' <para>
    '''   Use the method <c>YOsControl.nextOsControl()</c> to iterate on
    '''   next OS control.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YOsControl</c> object, corresponding to
    '''   the first OS control currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstOsControl() As YOsControl
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("OsControl", 0, p, size, neededsize, errmsg)
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
      Return YOsControl.FindOsControl(serial + "." + funcId)
    End Function

    REM --- (end of YOsControl public methods declaration)

  End Class

  REM --- (YOsControl functions)

  '''*
  ''' <summary>
  '''   Retrieves OS control for a given identifier.
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
  '''   This function does not require that the OS control is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YOsControl.isOnline()</c> to test if the OS control is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   OS control by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the OS control
  ''' </param>
  ''' <returns>
  '''   a <c>YOsControl</c> object allowing you to drive the OS control.
  ''' </returns>
  '''/
  Public Function yFindOsControl(ByVal func As String) As YOsControl
    Return YOsControl.FindOsControl(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of OS control currently accessible.
  ''' <para>
  '''   Use the method <c>YOsControl.nextOsControl()</c> to iterate on
  '''   next OS control.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YOsControl</c> object, corresponding to
  '''   the first OS control currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstOsControl() As YOsControl
    Return YOsControl.FirstOsControl()
  End Function


  REM --- (end of YOsControl functions)

End Module
