'*********************************************************************
'*
'* $Id: yocto_pressure.vb 14798 2014-01-31 14:58:42Z seb $
'*
'* Implements yFindPressure(), the high-level API for Pressure functions
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

Module yocto_pressure

    REM --- (YPressure return codes)
    REM --- (end of YPressure return codes)
  REM --- (YPressure globals)

  Public Delegate Sub YPressureValueCallback(ByVal func As YPressure, ByVal value As String)
  Public Delegate Sub YPressureTimedReportCallback(ByVal func As YPressure, ByVal measure As YMeasure)
  REM --- (end of YPressure globals)

  REM --- (YPressure class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YPressure
    Inherits YSensor
    REM --- (end of YPressure class start)

    REM --- (YPressure definitions)
    REM --- (end of YPressure definitions)

    REM --- (YPressure attributes declaration)
    Protected _valueCallbackPressure As YPressureValueCallback
    Protected _timedReportCallbackPressure As YPressureTimedReportCallback
    REM --- (end of YPressure attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Pressure"
      REM --- (YPressure attributes initialization)
      _valueCallbackPressure = Nothing
      _timedReportCallbackPressure = Nothing
      REM --- (end of YPressure attributes initialization)
    End Sub

  REM --- (YPressure private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YPressure private methods declaration)

    REM --- (YPressure public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a pressure sensor for a given identifier.
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
    '''   This function does not require that the pressure sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YPressure.isOnline()</c> to test if the pressure sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a pressure sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the pressure sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YPressure</c> object allowing you to drive the pressure sensor.
    ''' </returns>
    '''/
    Public Shared Function FindPressure(func As String) As YPressure
      Dim obj As YPressure
      obj = CType(YFunction._FindFromCache("Pressure", func), YPressure)
      If ((obj Is Nothing)) Then
        obj = New YPressure(func)
        YFunction._AddToCache("Pressure", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YPressureValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me , True)
      Else
        YFunction._UpdateValueCallbackList(Me , False)
      End If
      Me._valueCallbackPressure = callback
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
      If (Not (Me._valueCallbackPressure Is Nothing)) Then
        Me._valueCallbackPressure(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YPressureTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me , True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me , False)
      End If
      Me._timedReportCallbackPressure = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackPressure Is Nothing)) Then
        Me._timedReportCallbackPressure(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of pressure sensors started using <c>yFirstPressure()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPressure</c> object, corresponding to
    '''   a pressure sensor currently online, or a <c>null</c> pointer
    '''   if there are no more pressure sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextPressure() As YPressure
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YPressure.FindPressure(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of pressure sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YPressure.nextPressure()</c> to iterate on
    '''   next pressure sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPressure</c> object, corresponding to
    '''   the first pressure sensor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstPressure() As YPressure
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Pressure", 0, p, size, neededsize, errmsg)
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
      Return YPressure.FindPressure(serial + "." + funcId)
    End Function

    REM --- (end of YPressure public methods declaration)

  End Class

  REM --- (Pressure functions)

  '''*
  ''' <summary>
  '''   Retrieves a pressure sensor for a given identifier.
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
  '''   This function does not require that the pressure sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YPressure.isOnline()</c> to test if the pressure sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a pressure sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the pressure sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YPressure</c> object allowing you to drive the pressure sensor.
  ''' </returns>
  '''/
  Public Function yFindPressure(ByVal func As String) As YPressure
    Return YPressure.FindPressure(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of pressure sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YPressure.nextPressure()</c> to iterate on
  '''   next pressure sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YPressure</c> object, corresponding to
  '''   the first pressure sensor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstPressure() As YPressure
    Return YPressure.FirstPressure()
  End Function


  REM --- (end of Pressure functions)

End Module
