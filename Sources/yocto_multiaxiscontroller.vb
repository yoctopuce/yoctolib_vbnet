'*********************************************************************
'*
'* $Id: yocto_multiaxiscontroller.vb 27237 2017-04-21 16:36:03Z seb $
'*
'* Implements yFindMultiAxisController(), the high-level API for MultiAxisController functions
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

Module yocto_multiaxiscontroller

    REM --- (YMultiAxisController return codes)
    REM --- (end of YMultiAxisController return codes)
    REM --- (YMultiAxisController dlldef)
    REM --- (end of YMultiAxisController dlldef)
  REM --- (YMultiAxisController globals)

  Public Const Y_NAXIS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_GLOBALSTATE_ABSENT As Integer = 0
  Public Const Y_GLOBALSTATE_ALERT As Integer = 1
  Public Const Y_GLOBALSTATE_HI_Z As Integer = 2
  Public Const Y_GLOBALSTATE_STOP As Integer = 3
  Public Const Y_GLOBALSTATE_RUN As Integer = 4
  Public Const Y_GLOBALSTATE_BATCH As Integer = 5
  Public Const Y_GLOBALSTATE_INVALID As Integer = -1
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YMultiAxisControllerValueCallback(ByVal func As YMultiAxisController, ByVal value As String)
  Public Delegate Sub YMultiAxisControllerTimedReportCallback(ByVal func As YMultiAxisController, ByVal measure As YMeasure)
  REM --- (end of YMultiAxisController globals)

  REM --- (YMultiAxisController class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to drive a stepper motor.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMultiAxisController
    Inherits YFunction
    REM --- (end of YMultiAxisController class start)

    REM --- (YMultiAxisController definitions)
    Public Const NAXIS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const GLOBALSTATE_ABSENT As Integer = 0
    Public Const GLOBALSTATE_ALERT As Integer = 1
    Public Const GLOBALSTATE_HI_Z As Integer = 2
    Public Const GLOBALSTATE_STOP As Integer = 3
    Public Const GLOBALSTATE_RUN As Integer = 4
    Public Const GLOBALSTATE_BATCH As Integer = 5
    Public Const GLOBALSTATE_INVALID As Integer = -1
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YMultiAxisController definitions)

    REM --- (YMultiAxisController attributes declaration)
    Protected _nAxis As Integer
    Protected _globalState As Integer
    Protected _command As String
    Protected _valueCallbackMultiAxisController As YMultiAxisControllerValueCallback
    REM --- (end of YMultiAxisController attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "MultiAxisController"
      REM --- (YMultiAxisController attributes initialization)
      _nAxis = NAXIS_INVALID
      _globalState = GLOBALSTATE_INVALID
      _command = COMMAND_INVALID
      _valueCallbackMultiAxisController = Nothing
      REM --- (end of YMultiAxisController attributes initialization)
    End Sub

    REM --- (YMultiAxisController private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("nAxis") Then
        _nAxis = CInt(json_val.getLong("nAxis"))
      End If
      If json_val.has("globalState") Then
        _globalState = CInt(json_val.getLong("globalState"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YMultiAxisController private methods declaration)

    REM --- (YMultiAxisController public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the number of synchronized controllers.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of synchronized controllers
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_NAXIS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nAxis() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return NAXIS_INVALID
        End If
      End If
      res = Me._nAxis
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of synchronized controllers.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the number of synchronized controllers
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
    Public Function set_nAxis(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("nAxis", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the stepper motor set overall state.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_GLOBALSTATE_ABSENT</c>, <c>Y_GLOBALSTATE_ALERT</c>, <c>Y_GLOBALSTATE_HI_Z</c>,
    '''   <c>Y_GLOBALSTATE_STOP</c>, <c>Y_GLOBALSTATE_RUN</c> and <c>Y_GLOBALSTATE_BATCH</c> corresponding to
    '''   the stepper motor set overall state
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_GLOBALSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_globalState() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return GLOBALSTATE_INVALID
        End If
      End If
      res = Me._globalState
      Return res
    End Function

    Public Function get_command() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      res = Me._command
      Return res
    End Function


    Public Function set_command(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("command", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a multi-axis controller for a given identifier.
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
    '''   This function does not require that the multi-axis controller is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YMultiAxisController.isOnline()</c> to test if the multi-axis controller is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a multi-axis controller by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the multi-axis controller
    ''' </param>
    ''' <returns>
    '''   a <c>YMultiAxisController</c> object allowing you to drive the multi-axis controller.
    ''' </returns>
    '''/
    Public Shared Function FindMultiAxisController(func As String) As YMultiAxisController
      Dim obj As YMultiAxisController
      obj = CType(YFunction._FindFromCache("MultiAxisController", func), YMultiAxisController)
      If ((obj Is Nothing)) Then
        obj = New YMultiAxisController(func)
        YFunction._AddToCache("MultiAxisController", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YMultiAxisControllerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackMultiAxisController = callback
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
      If (Not (Me._valueCallbackMultiAxisController Is Nothing)) Then
        Me._valueCallbackMultiAxisController(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function sendCommand(command As String) As Integer
      Return Me.set_command(command)
    End Function

    '''*
    ''' <summary>
    '''   Reinitialize all controllers and clear all alert flags.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function reset() As Integer
      Return Me.sendCommand("Z")
    End Function

    '''*
    ''' <summary>
    '''   Starts all motors backward at the specified speeds, to search for the motor home position.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="speed">
    '''   desired speed for all axis, in steps per second.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function findHomePosition(speed As List(Of Double)) As Integer
      Dim cmd As String
      Dim i As Integer = 0
      Dim ndim As Integer = 0
      ndim = speed.Count
      cmd = "H" + Convert.ToString(CType(Math.Round(1000*speed(0)), Integer))
      i = 1
      While (i < ndim)
        cmd = "" +  cmd + "," + Convert.ToString(CType(Math.Round(1000*speed(i)), Integer))
        i = i + 1
      End While
      Return Me.sendCommand(cmd)
    End Function

    '''*
    ''' <summary>
    '''   Starts all motors synchronously to reach a given absolute position.
    ''' <para>
    '''   The time needed to reach the requested position will depend on the lowest
    '''   acceleration and max speed parameters configured for all motors.
    '''   The final position will be reached on all axis at the same time.
    ''' </para>
    ''' </summary>
    ''' <param name="absPos">
    '''   absolute position, measured in steps from each origin.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function moveTo(absPos As List(Of Double)) As Integer
      Dim cmd As String
      Dim i As Integer = 0
      Dim ndim As Integer = 0
      ndim = absPos.Count
      cmd = "M" + Convert.ToString(CType(Math.Round(16*absPos(0)), Integer))
      i = 1
      While (i < ndim)
        cmd = "" +  cmd + "," + Convert.ToString(CType(Math.Round(16*absPos(i)), Integer))
        i = i + 1
      End While
      Return Me.sendCommand(cmd)
    End Function

    '''*
    ''' <summary>
    '''   Starts all motors synchronously to reach a given relative position.
    ''' <para>
    '''   The time needed to reach the requested position will depend on the lowest
    '''   acceleration and max speed parameters configured for all motors.
    '''   The final position will be reached on all axis at the same time.
    ''' </para>
    ''' </summary>
    ''' <param name="relPos">
    '''   relative position, measured in steps from the current position.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function moveRel(relPos As List(Of Double)) As Integer
      Dim cmd As String
      Dim i As Integer = 0
      Dim ndim As Integer = 0
      ndim = relPos.Count
      cmd = "m" + Convert.ToString(CType(Math.Round(16*relPos(0)), Integer))
      i = 1
      While (i < ndim)
        cmd = "" +  cmd + "," + Convert.ToString(CType(Math.Round(16*relPos(i)), Integer))
        i = i + 1
      End While
      Return Me.sendCommand(cmd)
    End Function

    '''*
    ''' <summary>
    '''   Keep the motor in the same state for the specified amount of time, before processing next command.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="waitMs">
    '''   wait time, specified in milliseconds.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function pause(waitMs As Integer) As Integer
      Return Me.sendCommand("_" + Convert.ToString(waitMs))
    End Function

    '''*
    ''' <summary>
    '''   Stops the motor with an emergency alert, without taking any additional precaution.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function emergencyStop() As Integer
      Return Me.sendCommand("!")
    End Function

    '''*
    ''' <summary>
    '''   Stops the motor smoothly as soon as possible, without waiting for ongoing move completion.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function abortAndBrake() As Integer
      Return Me.sendCommand("B")
    End Function

    '''*
    ''' <summary>
    '''   Turn the controller into Hi-Z mode immediately, without waiting for ongoing move completion.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function abortAndHiZ() As Integer
      Return Me.sendCommand("z")
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of multi-axis controllers started using <c>yFirstMultiAxisController()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMultiAxisController</c> object, corresponding to
    '''   a multi-axis controller currently online, or a <c>Nothing</c> pointer
    '''   if there are no more multi-axis controllers to enumerate.
    ''' </returns>
    '''/
    Public Function nextMultiAxisController() As YMultiAxisController
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YMultiAxisController.FindMultiAxisController(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of multi-axis controllers currently accessible.
    ''' <para>
    '''   Use the method <c>YMultiAxisController.nextMultiAxisController()</c> to iterate on
    '''   next multi-axis controllers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMultiAxisController</c> object, corresponding to
    '''   the first multi-axis controller currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstMultiAxisController() As YMultiAxisController
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("MultiAxisController", 0, p, size, neededsize, errmsg)
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
      Return YMultiAxisController.FindMultiAxisController(serial + "." + funcId)
    End Function

    REM --- (end of YMultiAxisController public methods declaration)

  End Class

  REM --- (MultiAxisController functions)

  '''*
  ''' <summary>
  '''   Retrieves a multi-axis controller for a given identifier.
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
  '''   This function does not require that the multi-axis controller is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YMultiAxisController.isOnline()</c> to test if the multi-axis controller is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a multi-axis controller by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the multi-axis controller
  ''' </param>
  ''' <returns>
  '''   a <c>YMultiAxisController</c> object allowing you to drive the multi-axis controller.
  ''' </returns>
  '''/
  Public Function yFindMultiAxisController(ByVal func As String) As YMultiAxisController
    Return YMultiAxisController.FindMultiAxisController(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of multi-axis controllers currently accessible.
  ''' <para>
  '''   Use the method <c>YMultiAxisController.nextMultiAxisController()</c> to iterate on
  '''   next multi-axis controllers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YMultiAxisController</c> object, corresponding to
  '''   the first multi-axis controller currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMultiAxisController() As YMultiAxisController
    Return YMultiAxisController.FirstMultiAxisController()
  End Function


  REM --- (end of MultiAxisController functions)

End Module
