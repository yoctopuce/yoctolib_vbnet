'*********************************************************************
'*
'* $Id: yocto_motor.vb 5471 2012-02-29 18:45:35Z mvuilleu $
'*
'* Implements yFindMotor(), the high-level API for Motor functions
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

Module yocto_motor

  REM --- (definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YMotor, ByVal value As String)

Public Class YMotorMove
  Public moving As System.Int64 = YAPI_INVALID_LONG
  Public target As System.Int64 = YAPI_INVALID_LONG
  Public ms As System.Int64 = YAPI_INVALID_LONG
End Class


  Public Const Y_LOGICALNAME_INVALID As String = YAPI_INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI_INVALID_STRING
  Public Const Y_POWER_INVALID As Integer = YAPI_INVALID_INT
  Public Const Y_BREAKPOWER_INVALID As Integer = YAPI_INVALID_INT

  Public Y_POWERMOVE_INVALID As YMotorMove
  Public Y_BREAKMOVE_INVALID As YMotorMove

  REM --- (end of definitions)

  REM --- (YMotor implementation)

  Private _MotorCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   Yoctopuce application programming interface allows you not only to drive
  '''   power sent to motor to make it turn both ways, but also drive accelerations
  '''   and decelerations.
  ''' <para>
  '''   The motor will then accelerate automatically: you won't
  '''   have to monitor it. The API also allows to slow dow the motor by shortening
  '''   its terminals: the motor will then act as an electromagnetic break.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMotor
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI_INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI_INVALID_STRING
    Public Const POWER_INVALID As Integer = YAPI_INVALID_INT
    Public Const BREAKPOWER_INVALID As Integer = YAPI_INVALID_INT

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _power As Integer
    Protected _breakpower As Integer
    Protected _powermove As YMotorMove
    Protected _breakmove As YMotorMove

    Public Sub New(ByVal func As String)
      MyBase.new("Motor", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _power = Y_POWER_INVALID
      _breakpower = Y_BREAKPOWER_INVALID
      _powermove = New YMotorMove()
      _breakmove = New YMotorMove()
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
        ElseIf (member.name = "power") Then
          _power = member.ivalue
        ElseIf (member.name = "breakpower") Then
          _breakpower = member.ivalue
        ElseIf (member.name = "powermove") Then
          If (member.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then 
             _parse = -1
             Exit Function
          End If
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l=0 To member.membercount-1
             submemb = member.members(l)
             If (submemb.name = "moving") Then
                _powermove.moving = submemb.ivalue
             ElseIf (submemb.name = "target") Then
                _powermove.target = submemb.ivalue
             ElseIf (submemb.name = "ms") Then
                _powermove.ms = submemb.ivalue
             End If
          Next l
          
        ElseIf (member.name = "breakmove") Then
          If (member.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then 
             _parse = -1
             Exit Function
          End If
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l=0 To member.membercount-1
             submemb = member.members(l)
             If (submemb.name = "moving") Then
                _breakmove.moving = submemb.ivalue
             ElseIf (submemb.name = "target") Then
                _breakmove.target = submemb.ivalue
             ElseIf (submemb.name = "ms") Then
                _breakmove.ms = submemb.ivalue
             End If
          Next l
          
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the motor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the motor
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
    '''   Changes the logical name of the motor.
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
    '''   a string corresponding to the logical name of the motor
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
    '''   Returns the current value of the motor (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the motor (no more than 6 characters)
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

    '''*
    ''' <summary>
    '''   Return the power sent to the motor
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_POWER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_power() As Integer
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_POWER_INVALID
        End If
      End If
      Return _power
    End Function

    '''*
    ''' <summary>
    '''   Changes immediately the power sent to the motor.
    ''' <para>
    '''   If you want go easy on your mechanics
    '''   and avoid excessive current consumption which might exceed the controler capabilities
    '''   try to avoid brutal power changes. For example, immediate transition from forward full power
    '''   to reverse full power is a very bad idea. Each time the power is modified, the
    '''   breaking power is set to zero.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to immediately the power sent to the motor
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
    Public Function set_power(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("power", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Return the breaking power applied to the motor
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BREAKPOWER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_breakpower() As Integer
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_BREAKPOWER_INVALID
        End If
      End If
      Return _breakpower
    End Function

    '''*
    ''' <summary>
    '''   Changes immediately the breaking force applied to the motor.
    ''' <para>
    '''   Each time the
    '''   breaking value is changed, the power is set to zero
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to immediately the breaking force applied to the motor
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
    Public Function set_breakpower(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("breakpower", rest_val)
    End Function

    Public Function get_powermove() As YMotorMove
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_POWERMOVE_INVALID
        End If
      End If
      Return _powermove
    End Function

    Public Function set_powermove(ByVal newval As YMotorMove) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target))+":"+Ltrim(Str(newval.ms))
      Return _setAttr("powermove", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth linear accelaration/deceleratation to a given power.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   new power value at the end of the transistion
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the transition, in milliseconds
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
    Public Function powerMove(ByVal target As Integer,ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(target))+":"+Ltrim(Str(ms_duration))
      Return _setAttr("powermove", rest_val)
    End Function

    Public Function get_breakmove() As YMotorMove
      If (_cacheExpiration <= yGetTickCount()) Then
        If (YISERR(load(YAPI_DefaultCacheValidity))) Then
          Return Y_BREAKMOVE_INVALID
        End If
      End If
      Return _breakmove
    End Function

    Public Function set_breakmove(ByVal newval As YMotorMove) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target))+":"+Ltrim(Str(newval.ms))
      Return _setAttr("breakmove", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth breaking variation.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   new breaking value at the end of the transistion
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the transition, in milliseconds
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
    Public Function breakMove(ByVal target As Integer,ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(target))+":"+Ltrim(Str(ms_duration))
      Return _setAttr("breakmove", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of motors started using <c>yFirstMotor()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMotor</c> object, corresponding to
    '''   a motor currently online, or a <c>null</c> pointer
    '''   if there are no more motors to enumerate.
    ''' </returns>
    '''/
    Public Function nextMotor() as YMotor
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindMotor(hwid)
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
    '''   Retrieves a motor for a given identifier.
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
    '''   This function does not require that the motor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YMotor.isOnline()</c> to test if the motor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a motor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the motor
    ''' </param>
    ''' <returns>
    '''   a <c>YMotor</c> object allowing you to drive the motor.
    ''' </returns>
    '''/
    Public Shared Function FindMotor(ByVal func As String) As YMotor
      Dim res As YMotor
      If (_MotorCache.ContainsKey(func)) Then
        Return CType(_MotorCache(func), YMotor)
      End If
      res = New YMotor(func)
      _MotorCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of motors currently accessible.
    ''' <para>
    '''   Use the method <c>YMotor.nextMotor()</c> to iterate on
    '''   next motors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMotor</c> object, corresponding to
    '''   the first motor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstMotor() As YMotor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Motor", 0, p, size, neededsize, errmsg)
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
      Return YMotor.FindMotor(serial + "." + funcId)
    End Function

    REM --- (end of YMotor implementation)

  End Class

  REM --- (Motor functions)

  '''*
  ''' <summary>
  '''   Retrieves a motor for a given identifier.
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
  '''   This function does not require that the motor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YMotor.isOnline()</c> to test if the motor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a motor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the motor
  ''' </param>
  ''' <returns>
  '''   a <c>YMotor</c> object allowing you to drive the motor.
  ''' </returns>
  '''/
  Public Function yFindMotor(ByVal func As String) As YMotor
    Return YMotor.FindMotor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of motors currently accessible.
  ''' <para>
  '''   Use the method <c>YMotor.nextMotor()</c> to iterate on
  '''   next motors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YMotor</c> object, corresponding to
  '''   the first motor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMotor() As YMotor
    Return YMotor.FirstMotor()
  End Function

  Private Sub _MotorCleanup()
  End Sub


  REM --- (end of Motor functions)

End Module
