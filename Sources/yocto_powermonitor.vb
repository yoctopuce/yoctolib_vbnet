'*********************************************************************
'*
'* $Id: yocto_powermonitor.vb 5952 2012-04-04 05:40:52Z mvuilleu $
'*
'* Implements yFindPowerMonitor(), the high-level API for PowerMonitor functions
'*
'* - - - - - - - - - License information: - - - - - - - - - 
'*
'* Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'*
'* 1) If you have obtained this file from www.yoctopuce.com,
'*    Yoctopuce Sarl licenses to you (hereafter Licensee) the
'*    right to use, modify, copy, and integrate this source file
'*    into your own solution for the sole purpose of interfacing
'*    a Yoctopuce product with Licensee's solution.
'*
'*    The use of this file and all relationship between Yoctopuce 
'*    and Licensee are governed by Yoctopuce General Terms and 
'*    Conditions.
'*
'*    THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'*    WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING 
'*    WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS 
'*    FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'*    EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'*    INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA, 
'*    COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR 
'*    SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT 
'*    LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'*    CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'*    BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'*    WARRANTY, OR OTHERWISE.
'*
'* 2) If your intent is not to interface with Yoctopuce products,
'*    you are not entitled to use, read or create any derived
'*    material from this source file.
'*
'*********************************************************************/


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_powermonitor

  REM --- (definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YPowerMonitor, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI_INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI_INVALID_STRING
  Public Const Y_MONITORSTATE_POWEROK = 0
  Public Const Y_MONITORSTATE_UNDERVOLTAGE = 1
  Public Const Y_MONITORSTATE_OVERCURRENT = 2
  Public Const Y_MONITORSTATE_INVALID = -1

  Public Const Y_VOLTAGE_INVALID As Double = YAPI_INVALID_DOUBLE
  Public Const Y_CURRENT_INVALID As Double = YAPI_INVALID_DOUBLE
  Public Const Y_MINVOLTAGE_INVALID As Double = YAPI_INVALID_DOUBLE
  Public Const Y_MAXCURRENT_INVALID As Double = YAPI_INVALID_DOUBLE


  REM --- (end of definitions)

  REM --- (YPowerMonitor implementation)

  Private _PowerMonitorCache As New Hashtable()
  Private _callback As UpdateCallback

  Public Class YPowerMonitor
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI_INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI_INVALID_STRING
    Public Const MONITORSTATE_POWEROK = 0
    Public Const MONITORSTATE_UNDERVOLTAGE = 1
    Public Const MONITORSTATE_OVERCURRENT = 2
    Public Const MONITORSTATE_INVALID = -1

    Public Const VOLTAGE_INVALID As Double = YAPI_INVALID_DOUBLE
    Public Const CURRENT_INVALID As Double = YAPI_INVALID_DOUBLE
    Public Const MINVOLTAGE_INVALID As Double = YAPI_INVALID_DOUBLE
    Public Const MAXCURRENT_INVALID As Double = YAPI_INVALID_DOUBLE

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _monitorState As Integer
    Protected _voltage As Double
    Protected _current As Double
    Protected _minVoltage As Double
    Protected _maxCurrent As Double

    Public Sub New(ByVal func As String)
      MyBase.new("PowerMonitor", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _monitorState = Y_MONITORSTATE_INVALID
      _voltage = Y_VOLTAGE_INVALID
      _current = Y_CURRENT_INVALID
      _minVoltage = Y_MINVOLTAGE_INVALID
      _maxCurrent = Y_MAXCURRENT_INVALID
    End Sub

    Protected Overrides Function _parse(ByRef j As TJSONRECORD) As Integer
      Dim member As TJSONRECORD
      Dim i As Integer
      If (j.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then
        Return -1
      End If
      For i = 0 To j.membercount - 1
        member = j.members(i)
        If (member.name = "logicalName") Then
          _logicalName = member.svalue
        ElseIf (member.name = "advertisedValue") Then
          _advertisedValue = member.svalue
        ElseIf (member.name = "monitorState") Then
          _monitorState = member.ivalue
        ElseIf (member.name = "voltage") Then
          _voltage = Math.Round(member.ivalue/6553.6)/10
        ElseIf (member.name = "current") Then
          _current = Math.Round(member.ivalue/6553.6)/10
        ElseIf (member.name = "minVoltage") Then
          _minVoltage = Math.Round(member.ivalue/6553.6)/10
        ElseIf (member.name = "maxCurrent") Then
          _maxCurrent = Math.Round(member.ivalue/6553.6)/10
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the external power control.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the external power control
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_LOGICALNAME_INVALID
        End If
      End If
      Return _logicalName
    End Function

    '''*
    ''' <summary>
    '''   Changes the logical name of the external power control.
    ''' <para>
    '''   You can use <c>yCheckLogicalName()</c>
    '''   prior to this call to make sure that your parameter is valid.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the logical name of the external power control
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
    Public Function set_logicalName(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("logicalName", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of the external power control (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the external power control (no more than 6 characters)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_ADVERTISEDVALUE_INVALID
        End If
      End If
      Return _advertisedValue
    End Function

    Public Function get_monitorState() As Integer
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_MONITORSTATE_INVALID
        End If
      End If
      Return _monitorState
    End Function

    Public Function set_monitorState(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("monitorState", rest_val)
    End Function

    Public Function get_voltage() As Double
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_VOLTAGE_INVALID
        End If
      End If
      Return _voltage
    End Function

    Public Function get_current() As Double
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_CURRENT_INVALID
        End If
      End If
      Return _current
    End Function

    Public Function get_minVoltage() As Double
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_MINVOLTAGE_INVALID
        End If
      End If
      Return _minVoltage
    End Function

    Public Function set_minVoltage(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("minVoltage", rest_val)
    End Function

    Public Function get_maxCurrent() As Double
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_MAXCURRENT_INVALID
        End If
      End If
      Return _maxCurrent
    End Function

    Public Function set_maxCurrent(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("maxCurrent", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of external power controls started using <c>yFirstPowerMonitor()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPowerMonitor</c> object, corresponding to
    '''   an external power control currently online, or a <c>null</c> pointer
    '''   if there are no more external power controls to enumerate.
    ''' </returns>
    '''/
    Public Function nextPowerMonitor() as YPowerMonitor
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindPowerMonitor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   comment from .
    ''' <para>
    '''   yc definition
    ''' </para>
    ''' </summary>
    '''/
  Public Overloads Sub registerValueCallback(ByVal callback As UpdateCallback)
   If (callback IsNot Nothing) Then
     registerFuncCallback(Me)
   Else
     unregisterFuncCallback(Me)
   End If
   _callback = callback
  End Sub

  Public Sub set_callback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Sub setCallback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Overrides Sub advertiseValue(ByVal value As String)
    If (_callback IsNot Nothing) Then _callback(Me, value)
  End Sub


    '''*
    ''' <summary>
    '''   Retrieves an external power control for a given identifier.
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
    '''   This function does not require that the external power control is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YPowerMonitor.isOnline()</c> to test if the external power control is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an external power control by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the external power control
    ''' </param>
    ''' <returns>
    '''   a <c>YPowerMonitor</c> object allowing you to drive the external power control.
    ''' </returns>
    '''/
    Public Shared Function FindPowerMonitor(ByVal func As String) As YPowerMonitor
      Dim res As YPowerMonitor
      If (_PowerMonitorCache.ContainsKey(func)) Then
        Return CType(_PowerMonitorCache(func), YPowerMonitor)
      End If
      res = New YPowerMonitor(func)
      _PowerMonitorCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of external power controls currently accessible.
    ''' <para>
    '''   Use the method <c>YPowerMonitor.nextPowerMonitor()</c> to iterate on
    '''   next external power controls.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPowerMonitor</c> object, corresponding to
    '''   the first external power control currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstPowerMonitor() As YPowerMonitor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("PowerMonitor", 0, p, size, neededsize, errmsg)
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
      Return YPowerMonitor.FindPowerMonitor(serial + "." + funcId)
    End Function

    REM --- (end of YPowerMonitor implementation)

  End Class

  REM --- (PowerMonitor functions)

  '''*
  ''' <summary>
  '''   Retrieves an external power control for a given identifier.
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
  '''   This function does not require that the external power control is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YPowerMonitor.isOnline()</c> to test if the external power control is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an external power control by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the external power control
  ''' </param>
  ''' <returns>
  '''   a <c>YPowerMonitor</c> object allowing you to drive the external power control.
  ''' </returns>
  '''/
  Public Function yFindPowerMonitor(ByVal func As String) As YPowerMonitor
    Return YPowerMonitor.FindPowerMonitor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of external power controls currently accessible.
  ''' <para>
  '''   Use the method <c>YPowerMonitor.nextPowerMonitor()</c> to iterate on
  '''   next external power controls.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YPowerMonitor</c> object, corresponding to
  '''   the first external power control currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstPowerMonitor() As YPowerMonitor
    Return YPowerMonitor.FirstPowerMonitor()
  End Function

  Private Sub _PowerMonitorCleanup()
  End Sub


  REM --- (end of PowerMonitor functions)

End Module
