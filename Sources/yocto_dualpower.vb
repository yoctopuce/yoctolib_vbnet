'*********************************************************************
'*
'* $Id: yocto_dualpower.vb 9921 2013-02-20 09:39:16Z seb $
'*
'* Implements yFindDualPower(), the high-level API for DualPower functions
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

Module yocto_dualpower

  REM --- (YDualPower definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YDualPower, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_POWERSTATE_OFF = 0
  Public Const Y_POWERSTATE_FROM_USB = 1
  Public Const Y_POWERSTATE_FROM_EXT = 2
  Public Const Y_POWERSTATE_INVALID = -1

  Public Const Y_POWERCONTROL_AUTO = 0
  Public Const Y_POWERCONTROL_FROM_USB = 1
  Public Const Y_POWERCONTROL_FROM_EXT = 2
  Public Const Y_POWERCONTROL_OFF = 3
  Public Const Y_POWERCONTROL_INVALID = -1

  Public Const Y_EXTVOLTAGE_INVALID As Integer = YAPI.INVALID_UNSIGNED


  REM --- (end of YDualPower definitions)

  REM --- (YDualPower implementation)

  Private _DualPowerCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   Yoctopuce application programming interface allows you to control
  '''   the power source to use for module functions that require high current.
  ''' <para>
  '''   The module can also automatically disconnect the external power
  '''   when a voltage drop is observed on the external power source
  '''   (external battery running out of power).
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDualPower
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const POWERSTATE_OFF = 0
    Public Const POWERSTATE_FROM_USB = 1
    Public Const POWERSTATE_FROM_EXT = 2
    Public Const POWERSTATE_INVALID = -1

    Public Const POWERCONTROL_AUTO = 0
    Public Const POWERCONTROL_FROM_USB = 1
    Public Const POWERCONTROL_FROM_EXT = 2
    Public Const POWERCONTROL_OFF = 3
    Public Const POWERCONTROL_INVALID = -1

    Public Const EXTVOLTAGE_INVALID As Integer = YAPI.INVALID_UNSIGNED

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _powerState As Long
    Protected _powerControl As Long
    Protected _extVoltage As Long

    Public Sub New(ByVal func As String)
      MyBase.new("DualPower", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _powerState = Y_POWERSTATE_INVALID
      _powerControl = Y_POWERCONTROL_INVALID
      _extVoltage = Y_EXTVOLTAGE_INVALID
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
        ElseIf (member.name = "powerState") Then
          _powerState = CLng(member.ivalue)
        ElseIf (member.name = "powerControl") Then
          _powerControl = CLng(member.ivalue)
        ElseIf (member.name = "extVoltage") Then
          _extVoltage = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the power control.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the power control
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LOGICALNAME_INVALID
        End If
      End If
      Return _logicalName
    End Function

    '''*
    ''' <summary>
    '''   Changes the logical name of the power control.
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
    '''   a string corresponding to the logical name of the power control
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
    '''   Returns the current value of the power control (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the power control (no more than 6 characters)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ADVERTISEDVALUE_INVALID
        End If
      End If
      Return _advertisedValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the current power source for module functions that require lots of current.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_POWERSTATE_OFF</c>, <c>Y_POWERSTATE_FROM_USB</c> and
    '''   <c>Y_POWERSTATE_FROM_EXT</c> corresponding to the current power source for module functions that
    '''   require lots of current
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_POWERSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_powerState() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_POWERSTATE_INVALID
        End If
      End If
      Return CType(_powerState,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns the selected power source for module functions that require lots of current.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_POWERCONTROL_AUTO</c>, <c>Y_POWERCONTROL_FROM_USB</c>,
    '''   <c>Y_POWERCONTROL_FROM_EXT</c> and <c>Y_POWERCONTROL_OFF</c> corresponding to the selected power
    '''   source for module functions that require lots of current
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_POWERCONTROL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_powerControl() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_POWERCONTROL_INVALID
        End If
      End If
      Return CType(_powerControl,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the selected power source for module functions that require lots of current.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_POWERCONTROL_AUTO</c>, <c>Y_POWERCONTROL_FROM_USB</c>,
    '''   <c>Y_POWERCONTROL_FROM_EXT</c> and <c>Y_POWERCONTROL_OFF</c> corresponding to the selected power
    '''   source for module functions that require lots of current
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
    Public Function set_powerControl(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("powerControl", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the measured voltage on the external power source, in millivolts.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the measured voltage on the external power source, in millivolts
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_EXTVOLTAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_extVoltage() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_EXTVOLTAGE_INVALID
        End If
      End If
      Return CType(_extVoltage,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of dual power controls started using <c>yFirstDualPower()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDualPower</c> object, corresponding to
    '''   a dual power control currently online, or a <c>null</c> pointer
    '''   if there are no more dual power controls to enumerate.
    ''' </returns>
    '''/
    Public Function nextDualPower() as YDualPower
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindDualPower(hwid)
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
    '''   Retrieves a dual power control for a given identifier.
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
    '''   This function does not require that the power control is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YDualPower.isOnline()</c> to test if the power control is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a dual power control by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the power control
    ''' </param>
    ''' <returns>
    '''   a <c>YDualPower</c> object allowing you to drive the power control.
    ''' </returns>
    '''/
    Public Shared Function FindDualPower(ByVal func As String) As YDualPower
      Dim res As YDualPower
      If (_DualPowerCache.ContainsKey(func)) Then
        Return CType(_DualPowerCache(func), YDualPower)
      End If
      res = New YDualPower(func)
      _DualPowerCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of dual power controls currently accessible.
    ''' <para>
    '''   Use the method <c>YDualPower.nextDualPower()</c> to iterate on
    '''   next dual power controls.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDualPower</c> object, corresponding to
    '''   the first dual power control currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstDualPower() As YDualPower
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("DualPower", 0, p, size, neededsize, errmsg)
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
      Return YDualPower.FindDualPower(serial + "." + funcId)
    End Function

    REM --- (end of YDualPower implementation)

  End Class

  REM --- (DualPower functions)

  '''*
  ''' <summary>
  '''   Retrieves a dual power control for a given identifier.
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
  '''   This function does not require that the power control is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YDualPower.isOnline()</c> to test if the power control is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a dual power control by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the power control
  ''' </param>
  ''' <returns>
  '''   a <c>YDualPower</c> object allowing you to drive the power control.
  ''' </returns>
  '''/
  Public Function yFindDualPower(ByVal func As String) As YDualPower
    Return YDualPower.FindDualPower(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of dual power controls currently accessible.
  ''' <para>
  '''   Use the method <c>YDualPower.nextDualPower()</c> to iterate on
  '''   next dual power controls.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YDualPower</c> object, corresponding to
  '''   the first dual power control currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstDualPower() As YDualPower
    Return YDualPower.FirstDualPower()
  End Function

  Private Sub _DualPowerCleanup()
  End Sub


  REM --- (end of DualPower functions)

End Module
