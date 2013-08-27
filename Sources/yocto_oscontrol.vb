'*********************************************************************
'*
'* $Id: yocto_oscontrol.vb 12337 2013-08-14 15:22:22Z mvuilleu $
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

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YOsControl definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YOsControl, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SHUTDOWNCOUNTDOWN_INVALID As Integer = YAPI.INVALID_UNSIGNED


  REM --- (end of YOsControl definitions)

  REM --- (YOsControl implementation)

  Private _OsControlCache As New Hashtable()
  Private _callback As UpdateCallback

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
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const SHUTDOWNCOUNTDOWN_INVALID As Integer = YAPI.INVALID_UNSIGNED

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _shutdownCountdown As Long

    Public Sub New(ByVal func As String)
      MyBase.new("OsControl", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _shutdownCountdown = Y_SHUTDOWNCOUNTDOWN_INVALID
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
        ElseIf (member.name = "shutdownCountdown") Then
          _shutdownCountdown = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the OS control, corresponding to the network name of the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the OS control, corresponding to the network name of the module
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
    '''   Changes the logical name of the OS control.
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
    '''   a string corresponding to the logical name of the OS control
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
    '''   Returns the current value of the OS control (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the OS control (no more than 6 characters)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_SHUTDOWNCOUNTDOWN_INVALID
        End If
      End If
      Return CType(_shutdownCountdown,Integer)
    End Function

    Public Function set_shutdownCountdown(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("shutdownCountdown", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Schedules an OS shutdown after a given number of seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="secBeforeShutDown">
    '''   number of seconds before shutdown
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
    Public Function shutdown(ByVal secBeforeShutDown As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(secBeforeShutDown))
      Return _setAttr("shutdownCountdown", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of OS control started using <c>yFirstOsControl()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YOsControl</c> object, corresponding to
    '''   OS control currently online, or a <c>null</c> pointer
    '''   if there are no more OS control to enumerate.
    ''' </returns>
    '''/
    Public Function nextOsControl() as YOsControl
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindOsControl(hwid)
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
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the OS control
    ''' </param>
    ''' <returns>
    '''   a <c>YOsControl</c> object allowing you to drive the OS control.
    ''' </returns>
    '''/
    Public Shared Function FindOsControl(ByVal func As String) As YOsControl
      Dim res As YOsControl
      If (_OsControlCache.ContainsKey(func)) Then
        Return CType(_OsControlCache(func), YOsControl)
      End If
      res = New YOsControl(func)
      _OsControlCache.Add(func, res)
      Return res
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
    '''   the first OS control currently online, or a <c>null</c> pointer
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

    REM --- (end of YOsControl implementation)

  End Class

  REM --- (OsControl functions)

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
  '''   the first OS control currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstOsControl() As YOsControl
    Return YOsControl.FirstOsControl()
  End Function

  Private Sub _OsControlCleanup()
  End Sub


  REM --- (end of OsControl functions)

End Module
