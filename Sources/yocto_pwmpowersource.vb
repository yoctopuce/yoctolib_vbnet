'*********************************************************************
'*
'* $Id: yocto_pwmpowersource.vb 15529 2014-03-20 17:54:15Z seb $
'*
'* Implements yFindPwmPowerSource(), the high-level API for PwmPowerSource functions
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

Module yocto_pwmpowersource

    REM --- (YPwmPowerSource return codes)
    REM --- (end of YPwmPowerSource return codes)
  REM --- (YPwmPowerSource globals)

  Public Const Y_POWERMODE_USB_5V As Integer = 0
  Public Const Y_POWERMODE_USB_3V As Integer = 1
  Public Const Y_POWERMODE_EXT_V As Integer = 2
  Public Const Y_POWERMODE_OPNDRN As Integer = 3
  Public Const Y_POWERMODE_INVALID As Integer = -1

  Public Delegate Sub YPwmPowerSourceValueCallback(ByVal func As YPwmPowerSource, ByVal value As String)
  Public Delegate Sub YPwmPowerSourceTimedReportCallback(ByVal func As YPwmPowerSource, ByVal measure As YMeasure)
  REM --- (end of YPwmPowerSource globals)

  REM --- (YPwmPowerSource class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to configure
  '''   the voltage source used by all PWM on the same device.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YPwmPowerSource
    Inherits YFunction
    REM --- (end of YPwmPowerSource class start)

    REM --- (YPwmPowerSource definitions)
    Public Const POWERMODE_USB_5V As Integer = 0
    Public Const POWERMODE_USB_3V As Integer = 1
    Public Const POWERMODE_EXT_V As Integer = 2
    Public Const POWERMODE_OPNDRN As Integer = 3
    Public Const POWERMODE_INVALID As Integer = -1

    REM --- (end of YPwmPowerSource definitions)

    REM --- (YPwmPowerSource attributes declaration)
    Protected _powerMode As Integer
    Protected _valueCallbackPwmPowerSource As YPwmPowerSourceValueCallback
    REM --- (end of YPwmPowerSource attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "PwmPowerSource"
      REM --- (YPwmPowerSource attributes initialization)
      _powerMode = POWERMODE_INVALID
      _valueCallbackPwmPowerSource = Nothing
      REM --- (end of YPwmPowerSource attributes initialization)
    End Sub

    REM --- (YPwmPowerSource private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "powerMode") Then
        _powerMode = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YPwmPowerSource private methods declaration)

    REM --- (YPwmPowerSource public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the selected power source for the PWM on the same device
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_POWERMODE_USB_5V</c>, <c>Y_POWERMODE_USB_3V</c>, <c>Y_POWERMODE_EXT_V</c> and
    '''   <c>Y_POWERMODE_OPNDRN</c> corresponding to the selected power source for the PWM on the same device
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_POWERMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_powerMode() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return POWERMODE_INVALID
        End If
      End If
      Return Me._powerMode
    End Function


    '''*
    ''' <summary>
    '''   Changes  the PWM power source.
    ''' <para>
    '''   PWM can use isolated 5V from USB, isolated 3V from USB or
    '''   voltage from an external power source. The PWM can also work in open drain  mode. In that
    '''   mode, the PWM actively pulls the line down.
    '''   Warning: this setting is common to all PWM on the same device. If you change that parameter,
    '''   all PWM located on the same device are  affected.
    '''   If you want the change to be kept after a device reboot, make sure  to call the matching
    '''   module <c>saveToFlash()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_POWERMODE_USB_5V</c>, <c>Y_POWERMODE_USB_3V</c>, <c>Y_POWERMODE_EXT_V</c> and
    '''   <c>Y_POWERMODE_OPNDRN</c> corresponding to  the PWM power source
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
    Public Function set_powerMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("powerMode", rest_val)
    End Function
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
    '''   Use the method <c>YPwmPowerSource.isOnline()</c> to test if the voltage source is
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
    '''   a <c>YPwmPowerSource</c> object allowing you to drive the voltage source.
    ''' </returns>
    '''/
    Public Shared Function FindPwmPowerSource(func As String) As YPwmPowerSource
      Dim obj As YPwmPowerSource
      obj = CType(YFunction._FindFromCache("PwmPowerSource", func), YPwmPowerSource)
      If ((obj Is Nothing)) Then
        obj = New YPwmPowerSource(func)
        YFunction._AddToCache("PwmPowerSource", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YPwmPowerSourceValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackPwmPowerSource = callback
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
      If (Not (Me._valueCallbackPwmPowerSource Is Nothing)) Then
        Me._valueCallbackPwmPowerSource(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of Voltage sources started using <c>yFirstPwmPowerSource()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPwmPowerSource</c> object, corresponding to
    '''   a voltage source currently online, or a <c>null</c> pointer
    '''   if there are no more Voltage sources to enumerate.
    ''' </returns>
    '''/
    Public Function nextPwmPowerSource() As YPwmPowerSource
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YPwmPowerSource.FindPwmPowerSource(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of Voltage sources currently accessible.
    ''' <para>
    '''   Use the method <c>YPwmPowerSource.nextPwmPowerSource()</c> to iterate on
    '''   next Voltage sources.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPwmPowerSource</c> object, corresponding to
    '''   the first source currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstPwmPowerSource() As YPwmPowerSource
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("PwmPowerSource", 0, p, size, neededsize, errmsg)
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
      Return YPwmPowerSource.FindPwmPowerSource(serial + "." + funcId)
    End Function

    REM --- (end of YPwmPowerSource public methods declaration)

  End Class

  REM --- (PwmPowerSource functions)

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
  '''   Use the method <c>YPwmPowerSource.isOnline()</c> to test if the voltage source is
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
  '''   a <c>YPwmPowerSource</c> object allowing you to drive the voltage source.
  ''' </returns>
  '''/
  Public Function yFindPwmPowerSource(ByVal func As String) As YPwmPowerSource
    Return YPwmPowerSource.FindPwmPowerSource(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of Voltage sources currently accessible.
  ''' <para>
  '''   Use the method <c>YPwmPowerSource.nextPwmPowerSource()</c> to iterate on
  '''   next Voltage sources.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YPwmPowerSource</c> object, corresponding to
  '''   the first source currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstPwmPowerSource() As YPwmPowerSource
    Return YPwmPowerSource.FirstPwmPowerSource()
  End Function


  REM --- (end of PwmPowerSource functions)

End Module
