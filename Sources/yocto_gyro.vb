'*********************************************************************
'*
'* $Id: yocto_gyro.vb 14977 2014-02-14 06:19:29Z mvuilleu $
'*
'* Implements yFindGyro(), the high-level API for Gyro functions
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

Module yocto_gyro

    REM --- (YGyro return codes)
    REM --- (end of YGyro return codes)
  REM --- (YGyro globals)

  Public Const Y_AXIS_X = 0
  Public Const Y_AXIS_Y = 1
  Public Const Y_AXIS_Z = 2
  Public Const Y_AXIS_INVALID = -1

  Public Delegate Sub YGyroValueCallback(ByVal func As YGyro, ByVal value As String)
  Public Delegate Sub YGyroTimedReportCallback(ByVal func As YGyro, ByVal measure As YMeasure)
  REM --- (end of YGyro globals)

  REM --- (YGyro class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YGyro
    Inherits YSensor
    REM --- (end of YGyro class start)

    REM --- (YGyro definitions)
    Public Const AXIS_X = 0
    Public Const AXIS_Y = 1
    Public Const AXIS_Z = 2
    Public Const AXIS_INVALID = -1

    REM --- (end of YGyro definitions)

    REM --- (YGyro attributes declaration)
    Protected _axis As Integer
    Protected _valueCallbackGyro As YGyroValueCallback
    Protected _timedReportCallbackGyro As YGyroTimedReportCallback
    REM --- (end of YGyro attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Gyro"
      REM --- (YGyro attributes initialization)
      _axis = AXIS_INVALID
      _valueCallbackGyro = Nothing
      _timedReportCallbackGyro = Nothing
      REM --- (end of YGyro attributes initialization)
    End Sub

  REM --- (YGyro private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "axis") Then
        _axis = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YGyro private methods declaration)

    REM --- (YGyro public methods declaration)
    Public Function get_axis() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DEFAULTCACHEVALIDITY) <> YAPI.SUCCESS) Then
          Return AXIS_INVALID
        End If
      End If
      Return Me._axis
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a gyroscope for a given identifier.
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
    '''   This function does not require that the gyroscope is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YGyro.isOnline()</c> to test if the gyroscope is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a gyroscope by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the gyroscope
    ''' </param>
    ''' <returns>
    '''   a <c>YGyro</c> object allowing you to drive the gyroscope.
    ''' </returns>
    '''/
    Public Shared Function FindGyro(func As String) As YGyro
      Dim obj As YGyro
      obj = CType(YFunction._FindFromCache("Gyro", func), YGyro)
      If ((obj Is Nothing)) Then
        obj = New YGyro(func)
        YFunction._AddToCache("Gyro", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YGyroValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me , True)
      Else
        YFunction._UpdateValueCallbackList(Me , False)
      End If
      Me._valueCallbackGyro = callback
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
      If (Not (Me._valueCallbackGyro Is Nothing)) Then
        Me._valueCallbackGyro(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YGyroTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me , True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me , False)
      End If
      Me._timedReportCallbackGyro = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackGyro Is Nothing)) Then
        Me._timedReportCallbackGyro(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of gyroscopes started using <c>yFirstGyro()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGyro</c> object, corresponding to
    '''   a gyroscope currently online, or a <c>null</c> pointer
    '''   if there are no more gyroscopes to enumerate.
    ''' </returns>
    '''/
    Public Function nextGyro() As YGyro
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YGyro.FindGyro(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of gyroscopes currently accessible.
    ''' <para>
    '''   Use the method <c>YGyro.nextGyro()</c> to iterate on
    '''   next gyroscopes.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGyro</c> object, corresponding to
    '''   the first gyro currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstGyro() As YGyro
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Gyro", 0, p, size, neededsize, errmsg)
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
      Return YGyro.FindGyro(serial + "." + funcId)
    End Function

    REM --- (end of YGyro public methods declaration)

  End Class

  REM --- (Gyro functions)

  '''*
  ''' <summary>
  '''   Retrieves a gyroscope for a given identifier.
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
  '''   This function does not require that the gyroscope is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YGyro.isOnline()</c> to test if the gyroscope is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a gyroscope by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the gyroscope
  ''' </param>
  ''' <returns>
  '''   a <c>YGyro</c> object allowing you to drive the gyroscope.
  ''' </returns>
  '''/
  Public Function yFindGyro(ByVal func As String) As YGyro
    Return YGyro.FindGyro(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of gyroscopes currently accessible.
  ''' <para>
  '''   Use the method <c>YGyro.nextGyro()</c> to iterate on
  '''   next gyroscopes.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YGyro</c> object, corresponding to
  '''   the first gyro currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstGyro() As YGyro
    Return YGyro.FirstGyro()
  End Function


  REM --- (end of Gyro functions)

End Module
