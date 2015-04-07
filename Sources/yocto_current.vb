'*********************************************************************
'*
'* $Id: yocto_current.vb 19575 2015-03-04 10:42:56Z seb $
'*
'* Implements yFindCurrent(), the high-level API for Current functions
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

Module yocto_current

    REM --- (YCurrent return codes)
    REM --- (end of YCurrent return codes)
    REM --- (YCurrent dlldef)
    REM --- (end of YCurrent dlldef)
  REM --- (YCurrent globals)

  Public Delegate Sub YCurrentValueCallback(ByVal func As YCurrent, ByVal value As String)
  Public Delegate Sub YCurrentTimedReportCallback(ByVal func As YCurrent, ByVal measure As YMeasure)
  REM --- (end of YCurrent globals)

  REM --- (YCurrent class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce class YCurrent allows you to read and configure Yoctopuce current
  '''   sensors.
  ''' <para>
  '''   It inherits from YSensor class the core functions to read measurements,
  '''   register callback functions, access to the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YCurrent
    Inherits YSensor
    REM --- (end of YCurrent class start)

    REM --- (YCurrent definitions)
    REM --- (end of YCurrent definitions)

    REM --- (YCurrent attributes declaration)
    Protected _valueCallbackCurrent As YCurrentValueCallback
    Protected _timedReportCallbackCurrent As YCurrentTimedReportCallback
    REM --- (end of YCurrent attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Current"
      REM --- (YCurrent attributes initialization)
      _valueCallbackCurrent = Nothing
      _timedReportCallbackCurrent = Nothing
      REM --- (end of YCurrent attributes initialization)
    End Sub

    REM --- (YCurrent private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YCurrent private methods declaration)

    REM --- (YCurrent public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a current sensor for a given identifier.
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
    '''   This function does not require that the current sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YCurrent.isOnline()</c> to test if the current sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a current sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the current sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YCurrent</c> object allowing you to drive the current sensor.
    ''' </returns>
    '''/
    Public Shared Function FindCurrent(func As String) As YCurrent
      Dim obj As YCurrent
      obj = CType(YFunction._FindFromCache("Current", func), YCurrent)
      If ((obj Is Nothing)) Then
        obj = New YCurrent(func)
        YFunction._AddToCache("Current", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YCurrentValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackCurrent = callback
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
      If (Not (Me._valueCallbackCurrent Is Nothing)) Then
        Me._valueCallbackCurrent(Me, value)
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
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YCurrentTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me, True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me, False)
      End If
      Me._timedReportCallbackCurrent = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackCurrent Is Nothing)) Then
        Me._timedReportCallbackCurrent(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of current sensors started using <c>yFirstCurrent()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCurrent</c> object, corresponding to
    '''   a current sensor currently online, or a <c>null</c> pointer
    '''   if there are no more current sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextCurrent() As YCurrent
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YCurrent.FindCurrent(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of current sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YCurrent.nextCurrent()</c> to iterate on
    '''   next current sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCurrent</c> object, corresponding to
    '''   the first current sensor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstCurrent() As YCurrent
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Current", 0, p, size, neededsize, errmsg)
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
      Return YCurrent.FindCurrent(serial + "." + funcId)
    End Function

    REM --- (end of YCurrent public methods declaration)

  End Class

  REM --- (Current functions)

  '''*
  ''' <summary>
  '''   Retrieves a current sensor for a given identifier.
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
  '''   This function does not require that the current sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YCurrent.isOnline()</c> to test if the current sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a current sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the current sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YCurrent</c> object allowing you to drive the current sensor.
  ''' </returns>
  '''/
  Public Function yFindCurrent(ByVal func As String) As YCurrent
    Return YCurrent.FindCurrent(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of current sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YCurrent.nextCurrent()</c> to iterate on
  '''   next current sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YCurrent</c> object, corresponding to
  '''   the first current sensor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstCurrent() As YCurrent
    Return YCurrent.FirstCurrent()
  End Function


  REM --- (end of Current functions)

End Module
