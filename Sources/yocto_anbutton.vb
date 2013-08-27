'*********************************************************************
'*
'* $Id: yocto_anbutton.vb 12324 2013-08-13 15:10:31Z mvuilleu $
'*
'* Implements yFindAnButton(), the high-level API for AnButton functions
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

Module yocto_anbutton

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YAnButton definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YAnButton, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CALIBRATEDVALUE_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_RAWVALUE_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_ANALOGCALIBRATION_OFF = 0
  Public Const Y_ANALOGCALIBRATION_ON = 1
  Public Const Y_ANALOGCALIBRATION_INVALID = -1

  Public Const Y_CALIBRATIONMAX_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_CALIBRATIONMIN_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_SENSITIVITY_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_ISPRESSED_FALSE = 0
  Public Const Y_ISPRESSED_TRUE = 1
  Public Const Y_ISPRESSED_INVALID = -1

  Public Const Y_LASTTIMEPRESSED_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_LASTTIMERELEASED_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PULSECOUNTER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PULSETIMER_INVALID As Long = YAPI.INVALID_LONG


  REM --- (end of YAnButton definitions)

  REM --- (YAnButton implementation)

  Private _AnButtonCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   Yoctopuce application programming interface allows you to measure the state
  '''   of a simple button as well as to read an analog potentiometer (variable resistance).
  ''' <para>
  '''   This can be use for instance with a continuous rotating knob, a throttle grip
  '''   or a joystick. The module is capable to calibrate itself on min and max values,
  '''   in order to compute a calibrated value that varies proportionally with the
  '''   potentiometer position, regardless of its total resistance.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YAnButton
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const CALIBRATEDVALUE_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const RAWVALUE_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const ANALOGCALIBRATION_OFF = 0
    Public Const ANALOGCALIBRATION_ON = 1
    Public Const ANALOGCALIBRATION_INVALID = -1

    Public Const CALIBRATIONMAX_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const CALIBRATIONMIN_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const SENSITIVITY_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const ISPRESSED_FALSE = 0
    Public Const ISPRESSED_TRUE = 1
    Public Const ISPRESSED_INVALID = -1

    Public Const LASTTIMEPRESSED_INVALID As Long = YAPI.INVALID_LONG
    Public Const LASTTIMERELEASED_INVALID As Long = YAPI.INVALID_LONG
    Public Const PULSECOUNTER_INVALID As Long = YAPI.INVALID_LONG
    Public Const PULSETIMER_INVALID As Long = YAPI.INVALID_LONG

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _calibratedValue As Long
    Protected _rawValue As Long
    Protected _analogCalibration As Long
    Protected _calibrationMax As Long
    Protected _calibrationMin As Long
    Protected _sensitivity As Long
    Protected _isPressed As Long
    Protected _lastTimePressed As Long
    Protected _lastTimeReleased As Long
    Protected _pulseCounter As Long
    Protected _pulseTimer As Long

    Public Sub New(ByVal func As String)
      MyBase.new("AnButton", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _calibratedValue = Y_CALIBRATEDVALUE_INVALID
      _rawValue = Y_RAWVALUE_INVALID
      _analogCalibration = Y_ANALOGCALIBRATION_INVALID
      _calibrationMax = Y_CALIBRATIONMAX_INVALID
      _calibrationMin = Y_CALIBRATIONMIN_INVALID
      _sensitivity = Y_SENSITIVITY_INVALID
      _isPressed = Y_ISPRESSED_INVALID
      _lastTimePressed = Y_LASTTIMEPRESSED_INVALID
      _lastTimeReleased = Y_LASTTIMERELEASED_INVALID
      _pulseCounter = Y_PULSECOUNTER_INVALID
      _pulseTimer = Y_PULSETIMER_INVALID
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
        ElseIf (member.name = "calibratedValue") Then
          _calibratedValue = CLng(member.ivalue)
        ElseIf (member.name = "rawValue") Then
          _rawValue = CLng(member.ivalue)
        ElseIf (member.name = "analogCalibration") Then
          If (member.ivalue > 0) Then _analogCalibration = 1 Else _analogCalibration = 0
        ElseIf (member.name = "calibrationMax") Then
          _calibrationMax = CLng(member.ivalue)
        ElseIf (member.name = "calibrationMin") Then
          _calibrationMin = CLng(member.ivalue)
        ElseIf (member.name = "sensitivity") Then
          _sensitivity = CLng(member.ivalue)
        ElseIf (member.name = "isPressed") Then
          If (member.ivalue > 0) Then _isPressed = 1 Else _isPressed = 0
        ElseIf (member.name = "lastTimePressed") Then
          _lastTimePressed = CLng(member.ivalue)
        ElseIf (member.name = "lastTimeReleased") Then
          _lastTimeReleased = CLng(member.ivalue)
        ElseIf (member.name = "pulseCounter") Then
          _pulseCounter = CLng(member.ivalue)
        ElseIf (member.name = "pulseTimer") Then
          _pulseTimer = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the analog input.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the analog input
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
    '''   Changes the logical name of the analog input.
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
    '''   a string corresponding to the logical name of the analog input
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
    '''   Returns the current value of the analog input (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the analog input (no more than 6 characters)
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
    '''   Returns the current calibrated input value (between 0 and 1000, included).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current calibrated input value (between 0 and 1000, included)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALIBRATEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_calibratedValue() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CALIBRATEDVALUE_INVALID
        End If
      End If
      Return CType(_calibratedValue,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current measured input value as-is (between 0 and 4095, included).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current measured input value as-is (between 0 and 4095, included)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RAWVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rawValue() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_RAWVALUE_INVALID
        End If
      End If
      Return CType(_rawValue,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Tells if a calibration process is currently ongoing.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ANALOGCALIBRATION_OFF</c> or <c>Y_ANALOGCALIBRATION_ON</c>
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ANALOGCALIBRATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_analogCalibration() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ANALOGCALIBRATION_INVALID
        End If
      End If
      Return CType(_analogCalibration,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Starts or stops the calibration process.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module at the end of the calibration if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_ANALOGCALIBRATION_OFF</c> or <c>Y_ANALOGCALIBRATION_ON</c>
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
    Public Function set_analogCalibration(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("analogCalibration", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the maximal value measured during the calibration (between 0 and 4095, included).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximal value measured during the calibration (between 0 and 4095, included)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALIBRATIONMAX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_calibrationMax() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CALIBRATIONMAX_INVALID
        End If
      End If
      Return CType(_calibrationMax,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the maximal calibration value for the input (between 0 and 4095, included), without actually
    '''   starting the automated calibration.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the maximal calibration value for the input (between 0 and 4095,
    '''   included), without actually
    '''   starting the automated calibration
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
    Public Function set_calibrationMax(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("calibrationMax", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the minimal value measured during the calibration (between 0 and 4095, included).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minimal value measured during the calibration (between 0 and 4095, included)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALIBRATIONMIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_calibrationMin() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_CALIBRATIONMIN_INVALID
        End If
      End If
      Return CType(_calibrationMin,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the minimal calibration value for the input (between 0 and 4095, included), without actually
    '''   starting the automated calibration.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minimal calibration value for the input (between 0 and 4095,
    '''   included), without actually
    '''   starting the automated calibration
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
    Public Function set_calibrationMin(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("calibrationMin", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensibility for the input (between 1 and 1000) for triggering user callbacks.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the sensibility for the input (between 1 and 1000) for triggering user callbacks
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SENSITIVITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_sensitivity() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_SENSITIVITY_INVALID
        End If
      End If
      Return CType(_sensitivity,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the sensibility for the input (between 1 and 1000) for triggering user callbacks.
    ''' <para>
    '''   The sensibility is used to filter variations around a fixed value, but does not preclude the
    '''   transmission of events when the input value evolves constantly in the same direction.
    '''   Special case: when the value 1000 is used, the callback will only be thrown when the logical state
    '''   of the input switches from pressed to released and back.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the sensibility for the input (between 1 and 1000) for triggering user callbacks
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
    Public Function set_sensitivity(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("sensitivity", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns true if the input (considered as binary) is active (closed contact), and false otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ISPRESSED_FALSE</c> or <c>Y_ISPRESSED_TRUE</c>, according to true if the input
    '''   (considered as binary) is active (closed contact), and false otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ISPRESSED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_isPressed() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ISPRESSED_INVALID
        End If
      End If
      Return CType(_isPressed,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of elapsed milliseconds between the module power on and the last time
    '''   the input button was pressed (the input contact transitionned from open to closed).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of elapsed milliseconds between the module power on and the last time
    '''   the input button was pressed (the input contact transitionned from open to closed)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LASTTIMEPRESSED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastTimePressed() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LASTTIMEPRESSED_INVALID
        End If
      End If
      Return _lastTimePressed
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of elapsed milliseconds between the module power on and the last time
    '''   the input button was released (the input contact transitionned from closed to open).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of elapsed milliseconds between the module power on and the last time
    '''   the input button was released (the input contact transitionned from closed to open)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LASTTIMERELEASED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastTimeReleased() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LASTTIMERELEASED_INVALID
        End If
      End If
      Return _lastTimeReleased
    End Function

    '''*
    ''' <summary>
    '''   Returns the pulse counter value
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the pulse counter value
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PULSECOUNTER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pulseCounter() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PULSECOUNTER_INVALID
        End If
      End If
      Return _pulseCounter
    End Function

    Public Function set_pulseCounter(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pulseCounter", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the pulse counter value as well as his timer
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function resetCounter() As Integer
      Dim rest_val As String
      rest_val = "0"
      Return _setAttr("pulseCounter", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the timer of the pulses counter (ms)
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the timer of the pulses counter (ms)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PULSETIMER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pulseTimer() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PULSETIMER_INVALID
        End If
      End If
      Return _pulseTimer
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of analog inputs started using <c>yFirstAnButton()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAnButton</c> object, corresponding to
    '''   an analog input currently online, or a <c>null</c> pointer
    '''   if there are no more analog inputs to enumerate.
    ''' </returns>
    '''/
    Public Function nextAnButton() as YAnButton
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindAnButton(hwid)
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
    '''   Retrieves an analog input for a given identifier.
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
    '''   This function does not require that the analog input is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YAnButton.isOnline()</c> to test if the analog input is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an analog input by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the analog input
    ''' </param>
    ''' <returns>
    '''   a <c>YAnButton</c> object allowing you to drive the analog input.
    ''' </returns>
    '''/
    Public Shared Function FindAnButton(ByVal func As String) As YAnButton
      Dim res As YAnButton
      If (_AnButtonCache.ContainsKey(func)) Then
        Return CType(_AnButtonCache(func), YAnButton)
      End If
      res = New YAnButton(func)
      _AnButtonCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of analog inputs currently accessible.
    ''' <para>
    '''   Use the method <c>YAnButton.nextAnButton()</c> to iterate on
    '''   next analog inputs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAnButton</c> object, corresponding to
    '''   the first analog input currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstAnButton() As YAnButton
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("AnButton", 0, p, size, neededsize, errmsg)
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
      Return YAnButton.FindAnButton(serial + "." + funcId)
    End Function

    REM --- (end of YAnButton implementation)

  End Class

  REM --- (AnButton functions)

  '''*
  ''' <summary>
  '''   Retrieves an analog input for a given identifier.
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
  '''   This function does not require that the analog input is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YAnButton.isOnline()</c> to test if the analog input is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an analog input by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the analog input
  ''' </param>
  ''' <returns>
  '''   a <c>YAnButton</c> object allowing you to drive the analog input.
  ''' </returns>
  '''/
  Public Function yFindAnButton(ByVal func As String) As YAnButton
    Return YAnButton.FindAnButton(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of analog inputs currently accessible.
  ''' <para>
  '''   Use the method <c>YAnButton.nextAnButton()</c> to iterate on
  '''   next analog inputs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YAnButton</c> object, corresponding to
  '''   the first analog input currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstAnButton() As YAnButton
    Return YAnButton.FirstAnButton()
  End Function

  Private Sub _AnButtonCleanup()
  End Sub


  REM --- (end of AnButton functions)

End Module
