'*********************************************************************
'*
'* $Id: yocto_currentloopoutput.vb 27237 2017-04-21 16:36:03Z seb $
'*
'* Implements yFindCurrentLoopOutput(), the high-level API for CurrentLoopOutput functions
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

Module yocto_currentloopoutput

    REM --- (YCurrentLoopOutput return codes)
    REM --- (end of YCurrentLoopOutput return codes)
    REM --- (YCurrentLoopOutput dlldef)
    REM --- (end of YCurrentLoopOutput dlldef)
  REM --- (YCurrentLoopOutput globals)

  Public Const Y_CURRENT_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_CURRENTTRANSITION_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CURRENTATSTARTUP_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_LOOPPOWER_NOPWR As Integer = 0
  Public Const Y_LOOPPOWER_LOWPWR As Integer = 1
  Public Const Y_LOOPPOWER_POWEROK As Integer = 2
  Public Const Y_LOOPPOWER_INVALID As Integer = -1
  Public Delegate Sub YCurrentLoopOutputValueCallback(ByVal func As YCurrentLoopOutput, ByVal value As String)
  Public Delegate Sub YCurrentLoopOutputTimedReportCallback(ByVal func As YCurrentLoopOutput, ByVal measure As YMeasure)
  REM --- (end of YCurrentLoopOutput globals)

  REM --- (YCurrentLoopOutput class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to change the value of the 4-20mA
  '''   output as well as to know the current loop state.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YCurrentLoopOutput
    Inherits YFunction
    REM --- (end of YCurrentLoopOutput class start)

    REM --- (YCurrentLoopOutput definitions)
    Public Const CURRENT_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const CURRENTTRANSITION_INVALID As String = YAPI.INVALID_STRING
    Public Const CURRENTATSTARTUP_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const LOOPPOWER_NOPWR As Integer = 0
    Public Const LOOPPOWER_LOWPWR As Integer = 1
    Public Const LOOPPOWER_POWEROK As Integer = 2
    Public Const LOOPPOWER_INVALID As Integer = -1
    REM --- (end of YCurrentLoopOutput definitions)

    REM --- (YCurrentLoopOutput attributes declaration)
    Protected _current As Double
    Protected _currentTransition As String
    Protected _currentAtStartUp As Double
    Protected _loopPower As Integer
    Protected _valueCallbackCurrentLoopOutput As YCurrentLoopOutputValueCallback
    REM --- (end of YCurrentLoopOutput attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "CurrentLoopOutput"
      REM --- (YCurrentLoopOutput attributes initialization)
      _current = CURRENT_INVALID
      _currentTransition = CURRENTTRANSITION_INVALID
      _currentAtStartUp = CURRENTATSTARTUP_INVALID
      _loopPower = LOOPPOWER_INVALID
      _valueCallbackCurrentLoopOutput = Nothing
      REM --- (end of YCurrentLoopOutput attributes initialization)
    End Sub

    REM --- (YCurrentLoopOutput private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("current") Then
        _current = Math.Round(json_val.getDouble("current") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("currentTransition") Then
        _currentTransition = json_val.getString("currentTransition")
      End If
      If json_val.has("currentAtStartUp") Then
        _currentAtStartUp = Math.Round(json_val.getDouble("currentAtStartUp") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("loopPower") Then
        _loopPower = CInt(json_val.getLong("loopPower"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YCurrentLoopOutput private methods declaration)

    REM --- (YCurrentLoopOutput public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the current loop, the valid range is from 3 to 21mA.
    ''' <para>
    '''   If the loop is
    '''   not propely powered, the  target current is not reached and
    '''   loopPower is set to LOWPWR.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the current loop, the valid range is from 3 to 21mA
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
    Public Function set_current(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("current", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the loop current set point in mA.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the loop current set point in mA
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_current() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CURRENT_INVALID
        End If
      End If
      res = Me._current
      Return res
    End Function

    Public Function get_currentTransition() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CURRENTTRANSITION_INVALID
        End If
      End If
      res = Me._currentTransition
      Return res
    End Function


    Public Function set_currentTransition(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("currentTransition", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the loop current at device start up.
    ''' <para>
    '''   Remember to call the matching
    '''   module <c>saveToFlash()</c> method, otherwise this call has no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the loop current at device start up
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
    Public Function set_currentAtStartUp(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("currentAtStartUp", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current in the loop at device startup, in mA.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current in the loop at device startup, in mA
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENTATSTARTUP_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentAtStartUp() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CURRENTATSTARTUP_INVALID
        End If
      End If
      res = Me._currentAtStartUp
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the loop powerstate.
    ''' <para>
    '''   POWEROK: the loop
    '''   is powered. NOPWR: the loop in not powered. LOWPWR: the loop is not
    '''   powered enough to maintain the current required (insufficient voltage).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_LOOPPOWER_NOPWR</c>, <c>Y_LOOPPOWER_LOWPWR</c> and <c>Y_LOOPPOWER_POWEROK</c>
    '''   corresponding to the loop powerstate
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOOPPOWER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_loopPower() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LOOPPOWER_INVALID
        End If
      End If
      res = Me._loopPower
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a 4-20mA output for a given identifier.
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
    '''   This function does not require that the 4-20mA output is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YCurrentLoopOutput.isOnline()</c> to test if the 4-20mA output is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a 4-20mA output by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the 4-20mA output
    ''' </param>
    ''' <returns>
    '''   a <c>YCurrentLoopOutput</c> object allowing you to drive the 4-20mA output.
    ''' </returns>
    '''/
    Public Shared Function FindCurrentLoopOutput(func As String) As YCurrentLoopOutput
      Dim obj As YCurrentLoopOutput
      obj = CType(YFunction._FindFromCache("CurrentLoopOutput", func), YCurrentLoopOutput)
      If ((obj Is Nothing)) Then
        obj = New YCurrentLoopOutput(func)
        YFunction._AddToCache("CurrentLoopOutput", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerValueCallback(callback As YCurrentLoopOutputValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackCurrentLoopOutput = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) AndAlso Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overrides Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackCurrentLoopOutput Is Nothing)) Then
        Me._valueCallbackCurrentLoopOutput(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth transistion of current flowing in the loop.
    ''' <para>
    '''   Any current explicit
    '''   change cancels any ongoing transition process.
    ''' </para>
    ''' </summary>
    ''' <param name="mA_target">
    '''   new current value at the end of the transition
    '''   (floating-point number, representing the transition duration in mA)
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the transition, in milliseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    '''/
    Public Overridable Function currentMove(mA_target As Double, ms_duration As Integer) As Integer
      Dim newval As String
      If (mA_target < 3.0) Then
        mA_target  = 3.0
      End If
      If (mA_target > 21.0) Then
        mA_target = 21.0
      End If
      newval = "" + Convert.ToString( CType(Math.Round(mA_target*1000), Integer)) + ":" + Convert.ToString(ms_duration)
      
      Return Me.set_currentTransition(newval)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of 4-20mA outputs started using <c>yFirstCurrentLoopOutput()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCurrentLoopOutput</c> object, corresponding to
    '''   a 4-20mA output currently online, or a <c>Nothing</c> pointer
    '''   if there are no more 4-20mA outputs to enumerate.
    ''' </returns>
    '''/
    Public Function nextCurrentLoopOutput() As YCurrentLoopOutput
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YCurrentLoopOutput.FindCurrentLoopOutput(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of 4-20mA outputs currently accessible.
    ''' <para>
    '''   Use the method <c>YCurrentLoopOutput.nextCurrentLoopOutput()</c> to iterate on
    '''   next 4-20mA outputs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCurrentLoopOutput</c> object, corresponding to
    '''   the first 4-20mA output currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstCurrentLoopOutput() As YCurrentLoopOutput
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("CurrentLoopOutput", 0, p, size, neededsize, errmsg)
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
      Return YCurrentLoopOutput.FindCurrentLoopOutput(serial + "." + funcId)
    End Function

    REM --- (end of YCurrentLoopOutput public methods declaration)

  End Class

  REM --- (CurrentLoopOutput functions)

  '''*
  ''' <summary>
  '''   Retrieves a 4-20mA output for a given identifier.
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
  '''   This function does not require that the 4-20mA output is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YCurrentLoopOutput.isOnline()</c> to test if the 4-20mA output is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a 4-20mA output by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the 4-20mA output
  ''' </param>
  ''' <returns>
  '''   a <c>YCurrentLoopOutput</c> object allowing you to drive the 4-20mA output.
  ''' </returns>
  '''/
  Public Function yFindCurrentLoopOutput(ByVal func As String) As YCurrentLoopOutput
    Return YCurrentLoopOutput.FindCurrentLoopOutput(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of 4-20mA outputs currently accessible.
  ''' <para>
  '''   Use the method <c>YCurrentLoopOutput.nextCurrentLoopOutput()</c> to iterate on
  '''   next 4-20mA outputs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YCurrentLoopOutput</c> object, corresponding to
  '''   the first 4-20mA output currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstCurrentLoopOutput() As YCurrentLoopOutput
    Return YCurrentLoopOutput.FirstCurrentLoopOutput()
  End Function


  REM --- (end of CurrentLoopOutput functions)

End Module
