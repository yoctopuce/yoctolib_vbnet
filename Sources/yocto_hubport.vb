'*********************************************************************
'*
'* $Id: yocto_hubport.vb 12337 2013-08-14 15:22:22Z mvuilleu $
'*
'* Implements yFindHubPort(), the high-level API for HubPort functions
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

Module yocto_hubport

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YHubPort definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YHubPort, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ENABLED_FALSE = 0
  Public Const Y_ENABLED_TRUE = 1
  Public Const Y_ENABLED_INVALID = -1

  Public Const Y_PORTSTATE_OFF = 0
  Public Const Y_PORTSTATE_OVRLD = 1
  Public Const Y_PORTSTATE_ON = 2
  Public Const Y_PORTSTATE_RUN = 3
  Public Const Y_PORTSTATE_PROG = 4
  Public Const Y_PORTSTATE_INVALID = -1

  Public Const Y_BAUDRATE_INVALID As Integer = YAPI.INVALID_UNSIGNED


  REM --- (end of YHubPort definitions)

  REM --- (YHubPort implementation)

  Private _HubPortCache As New Hashtable()
  Private _callback As UpdateCallback

  Public Class YHubPort
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const ENABLED_FALSE = 0
    Public Const ENABLED_TRUE = 1
    Public Const ENABLED_INVALID = -1

    Public Const PORTSTATE_OFF = 0
    Public Const PORTSTATE_OVRLD = 1
    Public Const PORTSTATE_ON = 2
    Public Const PORTSTATE_RUN = 3
    Public Const PORTSTATE_PROG = 4
    Public Const PORTSTATE_INVALID = -1

    Public Const BAUDRATE_INVALID As Integer = YAPI.INVALID_UNSIGNED

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _enabled As Long
    Protected _portState As Long
    Protected _baudRate As Long

    Public Sub New(ByVal func As String)
      MyBase.new("HubPort", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _enabled = Y_ENABLED_INVALID
      _portState = Y_PORTSTATE_INVALID
      _baudRate = Y_BAUDRATE_INVALID
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
        ElseIf (member.name = "enabled") Then
          If (member.ivalue > 0) Then _enabled = 1 Else _enabled = 0
        ElseIf (member.name = "portState") Then
          _portState = CLng(member.ivalue)
        ElseIf (member.name = "baudRate") Then
          _baudRate = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the Yocto-hub port, which is always the serial number of the
    '''   connected module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the Yocto-hub port, which is always the serial number of the
    '''   connected module
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
    '''   It is not possible to configure the logical name of a Yocto-hub port.
    ''' <para>
    '''   The logical
    '''   name is automatically set to the serial number of the connected module.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string
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
    '''   Returns the current value of the Yocto-hub port (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the Yocto-hub port (no more than 6 characters)
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
    '''   Returns true if the Yocto-hub port is powered, false otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to true if the Yocto-hub port is
    '''   powered, false otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ENABLED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_enabled() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ENABLED_INVALID
        End If
      End If
      Return CType(_enabled,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the activation of the Yocto-hub port.
    ''' <para>
    '''   If the port is enabled, the
    '''   *      connected module is powered. Otherwise, port power is shut down.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to the activation of the Yocto-hub port
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
    Public Function set_enabled(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("enabled", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current state of the Yocto-hub port.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_PORTSTATE_OFF</c>, <c>Y_PORTSTATE_OVRLD</c>, <c>Y_PORTSTATE_ON</c>,
    '''   <c>Y_PORTSTATE_RUN</c> and <c>Y_PORTSTATE_PROG</c> corresponding to the current state of the Yocto-hub port
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PORTSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portState() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PORTSTATE_INVALID
        End If
      End If
      Return CType(_portState,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current baud rate used by this Yocto-hub port, in kbps.
    ''' <para>
    '''   The default value is 1000 kbps, but a slower rate may be used if communication
    '''   problems are encountered.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current baud rate used by this Yocto-hub port, in kbps
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BAUDRATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_baudRate() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_BAUDRATE_INVALID
        End If
      End If
      Return CType(_baudRate,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of Yocto-hub ports started using <c>yFirstHubPort()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHubPort</c> object, corresponding to
    '''   a Yocto-hub port currently online, or a <c>null</c> pointer
    '''   if there are no more Yocto-hub ports to enumerate.
    ''' </returns>
    '''/
    Public Function nextHubPort() as YHubPort
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindHubPort(hwid)
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
    '''   Retrieves a Yocto-hub port for a given identifier.
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
    '''   This function does not require that the Yocto-hub port is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YHubPort.isOnline()</c> to test if the Yocto-hub port is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a Yocto-hub port by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the Yocto-hub port
    ''' </param>
    ''' <returns>
    '''   a <c>YHubPort</c> object allowing you to drive the Yocto-hub port.
    ''' </returns>
    '''/
    Public Shared Function FindHubPort(ByVal func As String) As YHubPort
      Dim res As YHubPort
      If (_HubPortCache.ContainsKey(func)) Then
        Return CType(_HubPortCache(func), YHubPort)
      End If
      res = New YHubPort(func)
      _HubPortCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of Yocto-hub ports currently accessible.
    ''' <para>
    '''   Use the method <c>YHubPort.nextHubPort()</c> to iterate on
    '''   next Yocto-hub ports.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHubPort</c> object, corresponding to
    '''   the first Yocto-hub port currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstHubPort() As YHubPort
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("HubPort", 0, p, size, neededsize, errmsg)
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
      Return YHubPort.FindHubPort(serial + "." + funcId)
    End Function

    REM --- (end of YHubPort implementation)

  End Class

  REM --- (HubPort functions)

  '''*
  ''' <summary>
  '''   Retrieves a Yocto-hub port for a given identifier.
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
  '''   This function does not require that the Yocto-hub port is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YHubPort.isOnline()</c> to test if the Yocto-hub port is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a Yocto-hub port by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the Yocto-hub port
  ''' </param>
  ''' <returns>
  '''   a <c>YHubPort</c> object allowing you to drive the Yocto-hub port.
  ''' </returns>
  '''/
  Public Function yFindHubPort(ByVal func As String) As YHubPort
    Return YHubPort.FindHubPort(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of Yocto-hub ports currently accessible.
  ''' <para>
  '''   Use the method <c>YHubPort.nextHubPort()</c> to iterate on
  '''   next Yocto-hub ports.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YHubPort</c> object, corresponding to
  '''   the first Yocto-hub port currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstHubPort() As YHubPort
    Return YHubPort.FirstHubPort()
  End Function

  Private Sub _HubPortCleanup()
  End Sub


  REM --- (end of HubPort functions)

End Module
