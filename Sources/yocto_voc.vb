'*********************************************************************
'*
'* $Id: yocto_voc.vb 17356 2014-08-29 14:38:39Z seb $
'*
'* Implements yFindVoc(), the high-level API for Voc functions
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

Module yocto_voc

    REM --- (YVoc return codes)
    REM --- (end of YVoc return codes)
    REM --- (YVoc dlldef)
    REM --- (end of YVoc dlldef)
  REM --- (YVoc globals)

  Public Delegate Sub YVocValueCallback(ByVal func As YVoc, ByVal value As String)
  Public Delegate Sub YVocTimedReportCallback(ByVal func As YVoc, ByVal measure As YMeasure)
  REM --- (end of YVoc globals)

  REM --- (YVoc class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YVoc
    Inherits YSensor
    REM --- (end of YVoc class start)

    REM --- (YVoc definitions)
    REM --- (end of YVoc definitions)

    REM --- (YVoc attributes declaration)
    Protected _valueCallbackVoc As YVocValueCallback
    Protected _timedReportCallbackVoc As YVocTimedReportCallback
    REM --- (end of YVoc attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Voc"
      REM --- (YVoc attributes initialization)
      _valueCallbackVoc = Nothing
      _timedReportCallbackVoc = Nothing
      REM --- (end of YVoc attributes initialization)
    End Sub

    REM --- (YVoc private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YVoc private methods declaration)

    REM --- (YVoc public methods declaration)
    '''*
    ''' <summary>
    '''   Retrieves a Volatile Organic Compound sensor for a given identifier.
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
    '''   This function does not require that the Volatile Organic Compound sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YVoc.isOnline()</c> to test if the Volatile Organic Compound sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a Volatile Organic Compound sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the Volatile Organic Compound sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YVoc</c> object allowing you to drive the Volatile Organic Compound sensor.
    ''' </returns>
    '''/
    Public Shared Function FindVoc(func As String) As YVoc
      Dim obj As YVoc
      obj = CType(YFunction._FindFromCache("Voc", func), YVoc)
      If ((obj Is Nothing)) Then
        obj = New YVoc(func)
        YFunction._AddToCache("Voc", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YVocValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackVoc = callback
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
      If (Not (Me._valueCallbackVoc Is Nothing)) Then
        Me._valueCallbackVoc(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YVocTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me, True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me, False)
      End If
      Me._timedReportCallbackVoc = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackVoc Is Nothing)) Then
        Me._timedReportCallbackVoc(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of Volatile Organic Compound sensors started using <c>yFirstVoc()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVoc</c> object, corresponding to
    '''   a Volatile Organic Compound sensor currently online, or a <c>null</c> pointer
    '''   if there are no more Volatile Organic Compound sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextVoc() As YVoc
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YVoc.FindVoc(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of Volatile Organic Compound sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YVoc.nextVoc()</c> to iterate on
    '''   next Volatile Organic Compound sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YVoc</c> object, corresponding to
    '''   the first Volatile Organic Compound sensor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstVoc() As YVoc
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Voc", 0, p, size, neededsize, errmsg)
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
      Return YVoc.FindVoc(serial + "." + funcId)
    End Function

    REM --- (end of YVoc public methods declaration)

  End Class

  REM --- (Voc functions)

  '''*
  ''' <summary>
  '''   Retrieves a Volatile Organic Compound sensor for a given identifier.
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
  '''   This function does not require that the Volatile Organic Compound sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YVoc.isOnline()</c> to test if the Volatile Organic Compound sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a Volatile Organic Compound sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the Volatile Organic Compound sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YVoc</c> object allowing you to drive the Volatile Organic Compound sensor.
  ''' </returns>
  '''/
  Public Function yFindVoc(ByVal func As String) As YVoc
    Return YVoc.FindVoc(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of Volatile Organic Compound sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YVoc.nextVoc()</c> to iterate on
  '''   next Volatile Organic Compound sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YVoc</c> object, corresponding to
  '''   the first Volatile Organic Compound sensor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstVoc() As YVoc
    Return YVoc.FirstVoc()
  End Function


  REM --- (end of Voc functions)

End Module
