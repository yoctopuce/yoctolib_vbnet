' ********************************************************************
'
'  $Id: yocto_micropython.vb 63328 2024-11-13 09:35:22Z seb $
'
'  Implements yFindMicroPython(), the high-level API for MicroPython functions
'
'  - - - - - - - - - License information: - - - - - - - - -
'
'  Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'
'  Yoctopuce Sarl (hereafter Licensor) grants to you a perpetual
'  non-exclusive license to use, modify, copy and integrate this
'  file into your software for the sole purpose of interfacing
'  with Yoctopuce products.
'
'  You may reproduce and distribute copies of this file in
'  source or object form, as long as the sole purpose of this
'  code is to interface with Yoctopuce products. You must retain
'  this notice in the distributed source file.
'
'  You should refer to Yoctopuce General Terms and Conditions
'  for additional information regarding your rights and
'  obligations.
'
'  THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'  WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING
'  WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS
'  FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'  EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'  INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA,
'  COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR
'  SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT
'  LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'  CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'  BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'  WARRANTY, OR OTHERWISE.
'
' *********************************************************************


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_micropython

    REM --- (YMicroPython return codes)
    REM --- (end of YMicroPython return codes)
    REM --- (YMicroPython dlldef)
    REM --- (end of YMicroPython dlldef)
   REM --- (YMicroPython yapiwrapper)
   REM --- (end of YMicroPython yapiwrapper)
  REM --- (YMicroPython globals)

  Public Const Y_LASTMSG_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_HEAPUSAGE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_XHEAPUSAGE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CURRENTSCRIPT_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_STARTUPSCRIPT_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_DEBUGMODE_OFF As Integer = 0
  Public Const Y_DEBUGMODE_ON As Integer = 1
  Public Const Y_DEBUGMODE_INVALID As Integer = -1
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YMicroPythonValueCallback(ByVal func As YMicroPython, ByVal value As String)
  Public Delegate Sub YMicroPythonTimedReportCallback(ByVal func As YMicroPython, ByVal measure As YMeasure)
  Public Delegate Sub YMicroPythonLogCallback(ByVal obj As YMicroPython, ByVal logline As String)
  Public Delegate Sub YEventCallback(ByVal func As YMicroPython, ByVal logline As String)

  Sub yInternalEventCallback(ByVal func As YMicroPython, ByVal value As String)
    func._internalEventHandler(value)
  End Sub
  REM --- (end of YMicroPython globals)

  REM --- (YMicroPython class start)

  '''*
  ''' <summary>
  '''   The <c>YMicroPython</c> class provides control of the MicroPython interpreter
  '''   that can be found on some Yoctopuce devices.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMicroPython
    Inherits YFunction
    REM --- (end of YMicroPython class start)

    REM --- (YMicroPython definitions)
    Public Const LASTMSG_INVALID As String = YAPI.INVALID_STRING
    Public Const HEAPUSAGE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const XHEAPUSAGE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CURRENTSCRIPT_INVALID As String = YAPI.INVALID_STRING
    Public Const STARTUPSCRIPT_INVALID As String = YAPI.INVALID_STRING
    Public Const DEBUGMODE_OFF As Integer = 0
    Public Const DEBUGMODE_ON As Integer = 1
    Public Const DEBUGMODE_INVALID As Integer = -1
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YMicroPython definitions)

    REM --- (YMicroPython attributes declaration)
    Protected _lastMsg As String
    Protected _heapUsage As Integer
    Protected _xheapUsage As Integer
    Protected _currentScript As String
    Protected _startupScript As String
    Protected _debugMode As Integer
    Protected _command As String
    Protected _valueCallbackMicroPython As YMicroPythonValueCallback
    Protected _logCallback As YMicroPythonLogCallback
    Protected _isFirstCb As Boolean
    Protected _prevCbPos As Integer
    Protected _logPos As Integer
    Protected _prevPartialLog As String
    REM --- (end of YMicroPython attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "MicroPython"
      REM --- (YMicroPython attributes initialization)
      _lastMsg = LASTMSG_INVALID
      _heapUsage = HEAPUSAGE_INVALID
      _xheapUsage = XHEAPUSAGE_INVALID
      _currentScript = CURRENTSCRIPT_INVALID
      _startupScript = STARTUPSCRIPT_INVALID
      _debugMode = DEBUGMODE_INVALID
      _command = COMMAND_INVALID
      _valueCallbackMicroPython = Nothing
      _prevCbPos = 0
      _logPos = 0
      REM --- (end of YMicroPython attributes initialization)
    End Sub

    REM --- (YMicroPython private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("lastMsg") Then
        _lastMsg = json_val.getString("lastMsg")
      End If
      If json_val.has("heapUsage") Then
        _heapUsage = CInt(json_val.getLong("heapUsage"))
      End If
      If json_val.has("xheapUsage") Then
        _xheapUsage = CInt(json_val.getLong("xheapUsage"))
      End If
      If json_val.has("currentScript") Then
        _currentScript = json_val.getString("currentScript")
      End If
      If json_val.has("startupScript") Then
        _startupScript = json_val.getString("startupScript")
      End If
      If json_val.has("debugMode") Then
        If (json_val.getInt("debugMode") > 0) Then _debugMode = 1 Else _debugMode = 0
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YMicroPython private methods declaration)

    REM --- (YMicroPython public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the last message produced by a python script.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the last message produced by a python script
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMicroPython.LASTMSG_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastMsg() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LASTMSG_INVALID
        End If
      End If
      res = Me._lastMsg
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the percentage of micropython main memory in use,
    '''   as observed at the end of the last garbage collection.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the percentage of micropython main memory in use,
    '''   as observed at the end of the last garbage collection
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMicroPython.HEAPUSAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_heapUsage() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return HEAPUSAGE_INVALID
        End If
      End If
      res = Me._heapUsage
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the percentage of micropython external memory in use,
    '''   as observed at the end of the last garbage collection.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the percentage of micropython external memory in use,
    '''   as observed at the end of the last garbage collection
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMicroPython.XHEAPUSAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_xheapUsage() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return XHEAPUSAGE_INVALID
        End If
      End If
      res = Me._xheapUsage
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the name of currently active script, if any.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of currently active script, if any
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMicroPython.CURRENTSCRIPT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentScript() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CURRENTSCRIPT_INVALID
        End If
      End If
      res = Me._currentScript
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Stops current running script, and/or selects a script to run immediately in a
    '''   fresh new environment.
    ''' <para>
    '''   If the MicroPython interpreter is busy running a script,
    '''   this function will abort it immediately and reset the execution environment.
    '''   If a non-empty string is given as argument, the new script will be started.
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
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_currentScript(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("currentScript", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the name of the script to run when the device is powered on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of the script to run when the device is powered on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMicroPython.STARTUPSCRIPT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_startupScript() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return STARTUPSCRIPT_INVALID
        End If
      End If
      res = Me._startupScript
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the script to run when the device is powered on.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the script to run when the device is powered on
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_startupScript(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("startupScript", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the activation state of micropython debugging interface.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YMicroPython.DEBUGMODE_OFF</c> or <c>YMicroPython.DEBUGMODE_ON</c>, according to the
    '''   activation state of micropython debugging interface
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMicroPython.DEBUGMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_debugMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DEBUGMODE_INVALID
        End If
      End If
      res = Me._debugMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the activation state of micropython debugging interface.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YMicroPython.DEBUGMODE_OFF</c> or <c>YMicroPython.DEBUGMODE_ON</c>, according to the
    '''   activation state of micropython debugging interface
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_debugMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("debugMode", rest_val)
    End Function
    Public Function get_command() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
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
    '''   Retrieves a MicroPython interpreter for a given identifier.
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
    '''   This function does not require that the MicroPython interpreter is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YMicroPython.isOnline()</c> to test if the MicroPython interpreter is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a MicroPython interpreter by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' <para>
    '''   If a call to this object's is_online() method returns FALSE although
    '''   you are certain that the matching device is plugged, make sure that you did
    '''   call registerHub() at application initialization time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the MicroPython interpreter, for instance
    '''   <c>MyDevice.microPython</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YMicroPython</c> object allowing you to drive the MicroPython interpreter.
    ''' </returns>
    '''/
    Public Shared Function FindMicroPython(func As String) As YMicroPython
      Dim obj As YMicroPython
      obj = CType(YFunction._FindFromCache("MicroPython", func), YMicroPython)
      If ((obj Is Nothing)) Then
        obj = New YMicroPython(func)
        YFunction._AddToCache("MicroPython", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YMicroPythonValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackMicroPython = callback
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
      If (Not (Me._valueCallbackMicroPython Is Nothing)) Then
        Me._valueCallbackMicroPython(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Submit MicroPython code for execution in the interpreter.
    ''' <para>
    '''   If the MicroPython interpreter is busy, this function will
    '''   block until it becomes available. The code is then uploaded,
    '''   compiled and executed on the fly, without beeing stored on the device filesystem.
    ''' </para>
    ''' <para>
    '''   There is no implicit reset of the MicroPython interpreter with
    '''   this function. Use method <c>reset()</c> if you need to start
    '''   from a fresh environment to run your code.
    ''' </para>
    ''' <para>
    '''   Note that although MicroPython is mostly compatible with recent Python 3.x
    '''   interpreters, the limited ressources on the device impose some restrictions,
    '''   in particular regarding the libraries that can be used. Please refer to
    '''   the documentation for more details.
    ''' </para>
    ''' </summary>
    ''' <param name="codeName">
    '''   name of the code file (used for error reporting only)
    ''' </param>
    ''' <param name="mpyCode">
    '''   MicroPython code to compile and execute
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function eval(codeName As String, mpyCode As String) As Integer
      Dim fullname As String
      Dim res As Integer = 0
      fullname = "mpy:" + codeName
      res = Me._upload(fullname, YAPI.DefaultEncoding.GetBytes(mpyCode))
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Stops current execution, and reset the MicroPython interpreter to initial state.
    ''' <para>
    '''   All global variables are cleared, and all imports are forgotten.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function reset() As Integer
      Dim res As Integer = 0
      Dim state As String

      res = Me.set_command("Z")
      If Not(res = YAPI.SUCCESS) Then
        me._throw(YAPI.IO_ERROR, "unable to trigger MicroPython reset")
        return YAPI.IO_ERROR
      end if
      REM // Wait until the reset is effective
      state = (Me.get_advertisedValue()).Substring(0, 1)
      While (Not (state = "z"))
        YAPI.Sleep(50, Nothing)
        state = (Me.get_advertisedValue()).Substring(0, 1)
      End While
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns a string with last logs of the MicroPython interpreter.
    ''' <para>
    '''   This method return only logs that are still in the module.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with last MicroPython logs.
    '''   On failure, throws an exception or returns  <c>YAPI.INVALID_STRING</c>.
    ''' </returns>
    '''/
    Public Overridable Function get_lastLogs() As String
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim res As String

      buff = Me._download("mpy.txt")
      bufflen = (buff).Length - 1
      While ((bufflen > 0) AndAlso (buff(bufflen) <> 64))
        bufflen = bufflen - 1
      End While
      res = (YAPI.DefaultEncoding.GetString(buff)).Substring(0, bufflen)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Registers a device log callback function.
    ''' <para>
    '''   This callback will be called each time
    '''   microPython sends a new log message.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to invoke, or a Nothing pointer.
    '''   The callback function should take two arguments:
    '''   the module object that emitted the log message,
    '''   and the character string containing the log.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </param>
    '''/
    Public Overridable Function registerLogCallback(callback As YMicroPythonLogCallback) As Integer
      Dim serial As String

      serial = Me.get_serialNumber()
      If (serial = YAPI.INVALID_STRING) Then
        Return YAPI.DEVICE_NOT_FOUND
      End If
      Me._logCallback = callback
      Me._isFirstCb = True
      If (Not (callback Is Nothing)) Then
        Me.registerValueCallback(AddressOf yInternalEventCallback)
      Else
        Me.registerValueCallback(CType(Nothing, YMicroPythonValueCallback))
      End If
      Return 0
    End Function

    Public Overridable Function get_logCallback() As YMicroPythonLogCallback
      Return Me._logCallback
    End Function

    Public Overridable Function _internalEventHandler(cbVal As String) As Integer
      Dim cbPos As Integer = 0
      Dim cbDPos As Integer = 0
      Dim url As String
      Dim content As Byte() = New Byte(){}
      Dim endPos As Integer = 0
      Dim contentStr As String
      Dim msgArr As List(Of String) = New List(Of String)()
      Dim arrLen As Integer = 0
      Dim lenStr As String
      Dim arrPos As Integer = 0
      Dim logMsg As String
      REM // detect possible power cycle of the reader to clear event pointer
      cbPos = Convert.ToInt32((cbVal).Substring(1, (cbVal).Length-1), 16)
      cbDPos = ((cbPos - Me._prevCbPos) And (&Hfffff))
      Me._prevCbPos = cbPos
      If (cbDPos > 65536) Then
        Me._logPos = 0
      End If
      If (Not (Not (Me._logCallback Is Nothing))) Then
        Return YAPI.SUCCESS
      End If
      If (Me._isFirstCb) Then
        REM // use first emulated value callback caused by registerValueCallback:
        REM // to retrieve current logs position
        Me._logPos = 0
        Me._prevPartialLog = ""
        url = "mpy.txt"
      Else
        REM // load all messages since previous call
        url = "mpy.txt?pos=" + Convert.ToString(Me._logPos)
      End If

      content = Me._download(url)
      contentStr = YAPI.DefaultEncoding.GetString(content)
      REM // look for new position indicator at end of logs
      endPos = (content).Length - 1
      While ((endPos >= 0) AndAlso (content(endPos) <> 64))
        endPos = endPos - 1
      End While
      If Not(endPos > 0) Then
        me._throw(YAPI.IO_ERROR, "fail to download micropython logs")
        return YAPI.IO_ERROR
      end if
      lenStr = (contentStr).Substring(endPos+1, (contentStr).Length-(endPos+1))
      REM // update processed event position pointer
      Me._logPos = YAPI._atoi(lenStr)
      If (Me._isFirstCb) Then
        REM // don't generate callbacks log messages before call to registerLogCallback
        Me._isFirstCb = False
        Return YAPI.SUCCESS
      End If
      REM // now generate callbacks for each complete log line
      endPos = endPos - 1
      If Not(content(endPos) = 10) Then
        me._throw(YAPI.IO_ERROR, "fail to download micropython logs")
        return YAPI.IO_ERROR
      end if
      contentStr = (contentStr).Substring(0, endPos)
      msgArr = New List(Of String)(contentStr.Split(vbLf.ToCharArray()))
      arrLen = msgArr.Count - 1
      If (arrLen > 0) Then
        logMsg = "" + Me._prevPartialLog + "" + msgArr(0)
        If (Not (Me._logCallback Is Nothing)) Then
          Me._logCallback(Me, logMsg)
        End If
        Me._prevPartialLog = ""
        arrPos = 1
        While (arrPos < arrLen)
          logMsg = msgArr(arrPos)
          If (Not (Me._logCallback Is Nothing)) Then
            Me._logCallback(Me, logMsg)
          End If
          arrPos = arrPos + 1
        End While
      End If
      Me._prevPartialLog = "" + Me._prevPartialLog + "" + msgArr(arrLen)
      Return YAPI.SUCCESS
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of MicroPython interpreters started using <c>yFirstMicroPython()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned MicroPython interpreters order.
    '''   If you want to find a specific a MicroPython interpreter, use <c>MicroPython.findMicroPython()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMicroPython</c> object, corresponding to
    '''   a MicroPython interpreter currently online, or a <c>Nothing</c> pointer
    '''   if there are no more MicroPython interpreters to enumerate.
    ''' </returns>
    '''/
    Public Function nextMicroPython() As YMicroPython
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YMicroPython.FindMicroPython(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of MicroPython interpreters currently accessible.
    ''' <para>
    '''   Use the method <c>YMicroPython.nextMicroPython()</c> to iterate on
    '''   next MicroPython interpreters.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMicroPython</c> object, corresponding to
    '''   the first MicroPython interpreter currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstMicroPython() As YMicroPython
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("MicroPython", 0, p, size, neededsize, errmsg)
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
      Return YMicroPython.FindMicroPython(serial + "." + funcId)
    End Function

    REM --- (end of YMicroPython public methods declaration)

  End Class

  REM --- (YMicroPython functions)

  '''*
  ''' <summary>
  '''   Retrieves a MicroPython interpreter for a given identifier.
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
  '''   This function does not require that the MicroPython interpreter is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YMicroPython.isOnline()</c> to test if the MicroPython interpreter is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a MicroPython interpreter by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' <para>
  '''   If a call to this object's is_online() method returns FALSE although
  '''   you are certain that the matching device is plugged, make sure that you did
  '''   call registerHub() at application initialization time.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the MicroPython interpreter, for instance
  '''   <c>MyDevice.microPython</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YMicroPython</c> object allowing you to drive the MicroPython interpreter.
  ''' </returns>
  '''/
  Public Function yFindMicroPython(ByVal func As String) As YMicroPython
    Return YMicroPython.FindMicroPython(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of MicroPython interpreters currently accessible.
  ''' <para>
  '''   Use the method <c>YMicroPython.nextMicroPython()</c> to iterate on
  '''   next MicroPython interpreters.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YMicroPython</c> object, corresponding to
  '''   the first MicroPython interpreter currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMicroPython() As YMicroPython
    Return YMicroPython.FirstMicroPython()
  End Function


  REM --- (end of YMicroPython functions)

End Module
