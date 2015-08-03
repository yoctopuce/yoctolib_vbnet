'*********************************************************************
'*
'* $Id: yocto_refframe.vb 20508 2015-06-01 16:32:48Z seb $
'*
'* Implements yFindRefFrame(), the high-level API for RefFrame functions
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

Module yocto_refframe

    REM --- (YRefFrame return codes)
    REM --- (end of YRefFrame return codes)
    REM --- (YRefFrame dlldef)
    REM --- (end of YRefFrame dlldef)
  REM --- (YRefFrame globals)

 Public Enum  Y_MOUNTPOSITION
  BOTTOM = 0
  TOP = 1
  FRONT = 2
  REAR = 3
  RIGHT = 4
  LEFT = 5
end enum
 Public Enum  Y_MOUNTORIENTATION
  TWELVE = 0
  THREE = 1
  SIX = 2
  NINE = 3
end enum
  Public Const Y_MOUNTPOS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BEARING_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YRefFrameValueCallback(ByVal func As YRefFrame, ByVal value As String)
  Public Delegate Sub YRefFrameTimedReportCallback(ByVal func As YRefFrame, ByVal measure As YMeasure)
  REM --- (end of YRefFrame globals)

  REM --- (YRefFrame class start)

  '''*
  ''' <summary>
  '''   This class is used to setup the base orientation of the Yocto-3D, so that
  '''   the orientation functions, relative to the earth surface plane, use
  '''   the proper reference frame.
  ''' <para>
  '''   The class also implements a tridimensional
  '''   sensor calibration process, which can compensate for local variations
  '''   of standard gravity and improve the precision of the tilt sensors.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YRefFrame
    Inherits YFunction
    REM --- (end of YRefFrame class start)

    REM --- (YRefFrame definitions)
    Public Const MOUNTPOS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BEARING_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YRefFrame definitions)

    REM --- (YRefFrame attributes declaration)
    Protected _mountPos As Integer
    Protected _bearing As Double
    Protected _calibrationParam As String
    Protected _valueCallbackRefFrame As YRefFrameValueCallback
    Protected _calibStage As Integer
    Protected _calibStageHint As String
    Protected _calibStageProgress As Integer
    Protected _calibProgress As Integer
    Protected _calibLogMsg As String
    Protected _calibSavedParams As String
    Protected _calibCount As Integer
    Protected _calibInternalPos As Integer
    Protected _calibPrevTick As Integer
    Protected _calibOrient As List(Of Integer)
    Protected _calibDataAccX As List(Of Double)
    Protected _calibDataAccY As List(Of Double)
    Protected _calibDataAccZ As List(Of Double)
    Protected _calibDataAcc As List(Of Double)
    Protected _calibAccXOfs As Double
    Protected _calibAccYOfs As Double
    Protected _calibAccZOfs As Double
    Protected _calibAccXScale As Double
    Protected _calibAccYScale As Double
    Protected _calibAccZScale As Double
    REM --- (end of YRefFrame attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "RefFrame"
      REM --- (YRefFrame attributes initialization)
      _mountPos = MOUNTPOS_INVALID
      _bearing = BEARING_INVALID
      _calibrationParam = CALIBRATIONPARAM_INVALID
      _valueCallbackRefFrame = Nothing
      _calibStage = 0
      _calibStageProgress = 0
      _calibProgress = 0
      _calibCount = 0
      _calibInternalPos = 0
      _calibPrevTick = 0
      _calibOrient = New List(Of Integer)()
      _calibDataAccX = New List(Of Double)()
      _calibDataAccY = New List(Of Double)()
      _calibDataAccZ = New List(Of Double)()
      _calibDataAcc = New List(Of Double)()
      _calibAccXOfs = 0
      _calibAccYOfs = 0
      _calibAccZOfs = 0
      _calibAccXScale = 0
      _calibAccYScale = 0
      _calibAccZScale = 0
      REM --- (end of YRefFrame attributes initialization)
    End Sub

    REM --- (YRefFrame private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "mountPos") Then
        _mountPos = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "bearing") Then
        _bearing = Math.Round(member.ivalue * 1000.0 / 65536.0) / 1000.0
        Return 1
      End If
      If (member.name = "calibrationParam") Then
        _calibrationParam = member.svalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YRefFrame private methods declaration)

    REM --- (YRefFrame public methods declaration)
    Public Function get_mountPos() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MOUNTPOS_INVALID
        End If
      End If
      Return Me._mountPos
    End Function


    Public Function set_mountPos(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("mountPos", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the reference bearing used by the compass.
    ''' <para>
    '''   The relative bearing
    '''   indicated by the compass is the difference between the measured magnetic
    '''   heading and the reference bearing indicated here.
    ''' </para>
    ''' <para>
    '''   For instance, if you setup as reference bearing the value of the earth
    '''   magnetic declination, the compass will provide the orientation relative
    '''   to the geographic North.
    ''' </para>
    ''' <para>
    '''   Similarly, when the sensor is not mounted along the standard directions
    '''   because it has an additional yaw angle, you can set this angle in the reference
    '''   bearing so that the compass provides the expected natural direction.
    ''' </para>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the reference bearing used by the compass
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
    Public Function set_bearing(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("bearing", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the reference bearing used by the compass.
    ''' <para>
    '''   The relative bearing
    '''   indicated by the compass is the difference between the measured magnetic
    '''   heading and the reference bearing indicated here.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the reference bearing used by the compass
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BEARING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bearing() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BEARING_INVALID
        End If
      End If
      Return Me._bearing
    End Function

    Public Function get_calibrationParam() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALIBRATIONPARAM_INVALID
        End If
      End If
      Return Me._calibrationParam
    End Function


    Public Function set_calibrationParam(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("calibrationParam", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a reference frame for a given identifier.
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
    '''   This function does not require that the reference frame is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YRefFrame.isOnline()</c> to test if the reference frame is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a reference frame by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the reference frame
    ''' </param>
    ''' <returns>
    '''   a <c>YRefFrame</c> object allowing you to drive the reference frame.
    ''' </returns>
    '''/
    Public Shared Function FindRefFrame(func As String) As YRefFrame
      Dim obj As YRefFrame
      obj = CType(YFunction._FindFromCache("RefFrame", func), YRefFrame)
      If ((obj Is Nothing)) Then
        obj = New YRefFrame(func)
        YFunction._AddToCache("RefFrame", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YRefFrameValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackRefFrame = callback
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
      If (Not (Me._valueCallbackRefFrame Is Nothing)) Then
        Me._valueCallbackRefFrame(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the installation position of the device, as configured
    '''   in order to define the reference frame for the compass and the
    '''   pitch/roll tilt sensors.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among the <c>Y_MOUNTPOSITION</c> enumeration
    '''   (<c>Y_MOUNTPOSITION_BOTTOM</c>,   <c>Y_MOUNTPOSITION_TOP</c>,
    '''   <c>Y_MOUNTPOSITION_FRONT</c>,    <c>Y_MOUNTPOSITION_RIGHT</c>,
    '''   <c>Y_MOUNTPOSITION_REAR</c>,     <c>Y_MOUNTPOSITION_LEFT</c>),
    '''   corresponding to the installation in a box, on one of the six faces.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_mountPosition() As Y_MOUNTPOSITION
      Dim position As Integer = 0
      position = Me.get_mountPos()
      return CType(((position) >> (2)), Y_MOUNTPOSITION)
    End Function

    '''*
    ''' <summary>
    '''   Returns the installation orientation of the device, as configured
    '''   in order to define the reference frame for the compass and the
    '''   pitch/roll tilt sensors.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among the enumeration <c>Y_MOUNTORIENTATION</c>
    '''   (<c>Y_MOUNTORIENTATION_TWELVE</c>, <c>Y_MOUNTORIENTATION_THREE</c>,
    '''   <c>Y_MOUNTORIENTATION_SIX</c>,     <c>Y_MOUNTORIENTATION_NINE</c>)
    '''   corresponding to the orientation of the "X" arrow on the device,
    '''   as on a clock dial seen from an observer in the center of the box.
    '''   On the bottom face, the 12H orientation points to the front, while
    '''   on the top face, the 12H orientation points to the rear.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_mountOrientation() As Y_MOUNTORIENTATION
      Dim position As Integer = 0
      position = Me.get_mountPos()
      return CType(((position) And (3)), Y_MOUNTORIENTATION)
    End Function

    '''*
    ''' <summary>
    '''   Changes the compass and tilt sensor frame of reference.
    ''' <para>
    '''   The magnetic compass
    '''   and the tilt sensors (pitch and roll) naturally work in the plane
    '''   parallel to the earth surface. In case the device is not installed upright
    '''   and horizontally, you must select its reference orientation (parallel to
    '''   the earth surface) so that the measures are made relative to this position.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="position">
    '''   a value among the <c>Y_MOUNTPOSITION</c> enumeration
    '''   (<c>Y_MOUNTPOSITION_BOTTOM</c>,   <c>Y_MOUNTPOSITION_TOP</c>,
    '''   <c>Y_MOUNTPOSITION_FRONT</c>,    <c>Y_MOUNTPOSITION_RIGHT</c>,
    '''   <c>Y_MOUNTPOSITION_REAR</c>,     <c>Y_MOUNTPOSITION_LEFT</c>),
    '''   corresponding to the installation in a box, on one of the six faces.
    ''' </param>
    ''' <param name="orientation">
    '''   a value among the enumeration <c>Y_MOUNTORIENTATION</c>
    '''   (<c>Y_MOUNTORIENTATION_TWELVE</c>, <c>Y_MOUNTORIENTATION_THREE</c>,
    '''   <c>Y_MOUNTORIENTATION_SIX</c>,     <c>Y_MOUNTORIENTATION_NINE</c>)
    '''   corresponding to the orientation of the "X" arrow on the device,
    '''   as on a clock dial seen from an observer in the center of the box.
    '''   On the bottom face, the 12H orientation points to the front, while
    '''   on the top face, the 12H orientation points to the rear.
    ''' </param>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_mountPosition(position As Y_MOUNTPOSITION, orientation As Y_MOUNTORIENTATION) As Integer
      Dim mixedPos As Integer = 0
      mixedPos = ((position) << (2)) + orientation
      return Me.set_mountPos(mixedPos)
    End Function

    Public Overridable Function _calibSort(start As Integer, stopidx As Integer) As Integer
      Dim idx As Integer = 0
      Dim changed As Integer = 0
      Dim a As Double = 0
      Dim b As Double = 0
      Dim xa As Double = 0
      Dim xb As Double = 0
      
      REM // bubble sort is good since we will re-sort again after offset adjustment
      changed = 1
      While (changed > 0)
        changed = 0
        a = Me._calibDataAcc(start)
        idx = start + 1
        While (idx < stopidx)
          b = Me._calibDataAcc(idx)
          If (a > b) Then
            Me._calibDataAcc( idx-1) = b
            Me._calibDataAcc( idx) = a
            xa = Me._calibDataAccX(idx-1)
            xb = Me._calibDataAccX(idx)
            Me._calibDataAccX( idx-1) = xb
            Me._calibDataAccX( idx) = xa
            xa = Me._calibDataAccY(idx-1)
            xb = Me._calibDataAccY(idx)
            Me._calibDataAccY( idx-1) = xb
            Me._calibDataAccY( idx) = xa
            xa = Me._calibDataAccZ(idx-1)
            xb = Me._calibDataAccZ(idx)
            Me._calibDataAccZ( idx-1) = xb
            Me._calibDataAccZ( idx) = xa
            changed = changed + 1
          Else
            a = b
          End If
          idx = idx + 1
        End While
      End While
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Initiates the sensors tridimensional calibration process.
    ''' <para>
    '''   This calibration is used at low level for inertial position estimation
    '''   and to enhance the precision of the tilt sensors.
    ''' </para>
    ''' <para>
    '''   After calling this method, the device should be moved according to the
    '''   instructions provided by method <c>get_3DCalibrationHint</c>,
    '''   and <c>more3DCalibration</c> should be invoked about 5 times per second.
    '''   The calibration procedure is completed when the method
    '''   <c>get_3DCalibrationProgress</c> returns 100. At this point,
    '''   the computed calibration parameters can be applied using method
    '''   <c>save3DCalibration</c>. The calibration process can be canceled
    '''   at any time using method <c>cancel3DCalibration</c>.
    ''' </para>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function start3DCalibration() As Integer
      REM // may throw an exception
      If (Not (Me.isOnline())) Then
        Return YAPI.DEVICE_NOT_FOUND
      End If
      If (Me._calibStage <> 0) Then
        Me.cancel3DCalibration()
      End If
      Me._calibSavedParams = Me.get_calibrationParam()
      Me.set_calibrationParam("0")
      Me._calibCount = 50
      Me._calibStage = 1
      Me._calibStageHint = "Set down the device on a steady horizontal surface"
      Me._calibStageProgress = 0
      Me._calibProgress = 1
      Me._calibInternalPos = 0
      Me._calibPrevTick = CType(((YAPI.GetTickCount()) And (&H7FFFFFFF)), Integer)
      Me._calibOrient.Clear()
      Me._calibDataAccX.Clear()
      Me._calibDataAccY.Clear()
      Me._calibDataAccZ.Clear()
      Me._calibDataAcc.Clear()
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Continues the sensors tridimensional calibration process previously
    '''   initiated using method <c>start3DCalibration</c>.
    ''' <para>
    '''   This method should be called approximately 5 times per second, while
    '''   positioning the device according to the instructions provided by method
    '''   <c>get_3DCalibrationHint</c>. Note that the instructions change during
    '''   the calibration process.
    ''' </para>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function more3DCalibration() As Integer
      REM // may throw an exception
      Dim currTick As Integer = 0
      Dim jsonData As Byte()
      Dim xVal As Double = 0
      Dim yVal As Double = 0
      Dim zVal As Double = 0
      Dim xSq As Double = 0
      Dim ySq As Double = 0
      Dim zSq As Double = 0
      Dim norm As Double = 0
      Dim orient As Integer = 0
      Dim idx As Integer = 0
      Dim intpos As Integer = 0
      Dim err As Integer = 0
      REM // make sure calibration has been started
      If (Me._calibStage = 0) Then
        Return YAPI.INVALID_ARGUMENT
      End If
      If (Me._calibProgress = 100) Then
        Return YAPI.SUCCESS
      End If
      
      REM // make sure we leave at least 160ms between samples
      currTick =  CType(((YAPI.GetTickCount()) And (&H7FFFFFFF)), Integer)
      If (((currTick - Me._calibPrevTick) And (&H7FFFFFFF)) < 160) Then
        Return YAPI.SUCCESS
      End If
      REM // load current accelerometer values, make sure we are on a straight angle
      REM // (default timeout to 0,5 sec without reading measure when out of range)
      Me._calibStageHint = "Set down the device on a steady horizontal surface"
      Me._calibPrevTick = ((currTick + 500) And (&H7FFFFFFF))
      jsonData = Me._download("api/accelerometer.json")
      xVal = YAPI._atoi(Me._json_get_key(jsonData, "xValue")) / 65536.0
      yVal = YAPI._atoi(Me._json_get_key(jsonData, "yValue")) / 65536.0
      zVal = YAPI._atoi(Me._json_get_key(jsonData, "zValue")) / 65536.0
      xSq = xVal * xVal
      If (xSq >= 0.04 And xSq < 0.64) Then
        Return YAPI.SUCCESS
      End If
      If (xSq >= 1.44) Then
        Return YAPI.SUCCESS
      End If
      ySq = yVal * yVal
      If (ySq >= 0.04 And ySq < 0.64) Then
        Return YAPI.SUCCESS
      End If
      If (ySq >= 1.44) Then
        Return YAPI.SUCCESS
      End If
      zSq = zVal * zVal
      If (zSq >= 0.04 And zSq < 0.64) Then
        Return YAPI.SUCCESS
      End If
      If (zSq >= 1.44) Then
        Return YAPI.SUCCESS
      End If
      norm = Math.sqrt(xSq + ySq + zSq)
      If (norm < 0.8 Or norm > 1.2) Then
        Return YAPI.SUCCESS
      End If
      Me._calibPrevTick = currTick
      
      REM // Determine the device orientation index
      orient = 0
      If (zSq > 0.5) Then
        If (zVal > 0) Then
          orient = 0
        Else
          orient = 1
        End If
      End If
      If (xSq > 0.5) Then
        If (xVal > 0) Then
          orient = 2
        Else
          orient = 3
        End If
      End If
      If (ySq > 0.5) Then
        If (yVal > 0) Then
          orient = 4
        Else
          orient = 5
        End If
      End If
      
      REM // Discard measures that are not in the proper orientation
      If (Me._calibStageProgress = 0) Then
        REM
        idx = 0
        err = 0
        While (idx + 1 < Me._calibStage)
          If (Me._calibOrient(idx) = orient) Then
            err = 1
          End If
          idx = idx + 1
        End While
        If (err <> 0) Then
          Me._calibStageHint = "Turn the device on another face"
          Return YAPI.SUCCESS
        End If
        Me._calibOrient.Add(orient)
      Else
        REM
        If (orient <> Me._calibOrient(Me._calibStage-1)) Then
          Me._calibStageHint = "Not yet done, please move back to the previous face"
          Return YAPI.SUCCESS
        End If
      End If
      
      REM // Save measure
      Me._calibStageHint = "calibrating.."
      Me._calibDataAccX.Add(xVal)
      Me._calibDataAccY.Add(yVal)
      Me._calibDataAccZ.Add(zVal)
      Me._calibDataAcc.Add(norm)
      Me._calibInternalPos = Me._calibInternalPos + 1
      Me._calibProgress = 1 + 16 * (Me._calibStage - 1) + (16 * Me._calibInternalPos \ Me._calibCount)
      If (Me._calibInternalPos < Me._calibCount) Then
        Me._calibStageProgress = 1 + (99 * Me._calibInternalPos \ Me._calibCount)
        Return YAPI.SUCCESS
      End If
      
      REM // Stage done, compute preliminary result
      intpos = (Me._calibStage - 1) * Me._calibCount
      Me._calibSort(intpos, intpos + Me._calibCount)
      intpos = intpos + (Me._calibCount \ 2)
      Me._calibLogMsg = "Stage " + Convert.ToString( Me._calibStage) + ": median is " + Convert.ToString(
      CType(Math.Round(1000*Me._calibDataAccX(intpos)), Integer)) + "," + Convert.ToString(
      CType(Math.Round(1000*Me._calibDataAccY(intpos)), Integer)) + "," + Convert.ToString(CType(Math.Round(1000*Me._calibDataAccZ(intpos)), Integer))
      
      REM // move to next stage
      Me._calibStage = Me._calibStage + 1
      If (Me._calibStage < 7) Then
        Me._calibStageHint = "Turn the device on another face"
        Me._calibPrevTick = ((currTick + 500) And (&H7FFFFFFF))
        Me._calibStageProgress = 0
        Me._calibInternalPos = 0
        Return YAPI.SUCCESS
      End If
      REM // Data collection completed, compute accelerometer shift
      xVal = 0
      yVal = 0
      zVal = 0
      idx = 0
      While (idx < 6)
        intpos = idx * Me._calibCount + (Me._calibCount \ 2)
        orient = Me._calibOrient(idx)
        If (orient = 0 Or orient = 1) Then
          zVal = zVal + Me._calibDataAccZ(intpos)
        End If
        If (orient = 2 Or orient = 3) Then
          xVal = xVal + Me._calibDataAccX(intpos)
        End If
        If (orient = 4 Or orient = 5) Then
          yVal = yVal + Me._calibDataAccY(intpos)
        End If
        idx = idx + 1
      End While
      Me._calibAccXOfs = xVal / 2.0
      Me._calibAccYOfs = yVal / 2.0
      Me._calibAccZOfs = zVal / 2.0
      
      REM // Recompute all norms, taking into account the computed shift, and re-sort
      intpos = 0
      While (intpos < Me._calibDataAcc.Count)
        xVal = Me._calibDataAccX(intpos) - Me._calibAccXOfs
        yVal = Me._calibDataAccY(intpos) - Me._calibAccYOfs
        zVal = Me._calibDataAccZ(intpos) - Me._calibAccZOfs
        norm = Math.sqrt(xVal * xVal + yVal * yVal + zVal * zVal)
        Me._calibDataAcc( intpos) = norm
        intpos = intpos + 1
      End While
      idx = 0
      While (idx < 6)
        intpos = idx * Me._calibCount
        Me._calibSort(intpos, intpos + Me._calibCount)
        idx = idx + 1
      End While
      
      REM // Compute the scaling factor for each axis
      xVal = 0
      yVal = 0
      zVal = 0
      idx = 0
      While (idx < 6)
        intpos = idx * Me._calibCount + (Me._calibCount \ 2)
        orient = Me._calibOrient(idx)
        If (orient = 0 Or orient = 1) Then
          zVal = zVal + Me._calibDataAcc(intpos)
        End If
        If (orient = 2 Or orient = 3) Then
          xVal = xVal + Me._calibDataAcc(intpos)
        End If
        If (orient = 4 Or orient = 5) Then
          yVal = yVal + Me._calibDataAcc(intpos)
        End If
        idx = idx + 1
      End While
      Me._calibAccXScale = xVal / 2.0
      Me._calibAccYScale = yVal / 2.0
      Me._calibAccZScale = zVal / 2.0
      
      REM // Report completion
      Me._calibProgress = 100
      Me._calibStageHint = "Calibration data ready for saving"
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns instructions to proceed to the tridimensional calibration initiated with
    '''   method <c>start3DCalibration</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a character string.
    ''' </returns>
    '''/
    Public Overridable Function get_3DCalibrationHint() As String
      Return Me._calibStageHint
    End Function

    '''*
    ''' <summary>
    '''   Returns the global process indicator for the tridimensional calibration
    '''   initiated with method <c>start3DCalibration</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer between 0 (not started) and 100 (stage completed).
    ''' </returns>
    '''/
    Public Overridable Function get_3DCalibrationProgress() As Integer
      Return Me._calibProgress
    End Function

    '''*
    ''' <summary>
    '''   Returns index of the current stage of the calibration
    '''   initiated with method <c>start3DCalibration</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer, growing each time a calibration stage is completed.
    ''' </returns>
    '''/
    Public Overridable Function get_3DCalibrationStage() As Integer
      Return Me._calibStage
    End Function

    '''*
    ''' <summary>
    '''   Returns the process indicator for the current stage of the calibration
    '''   initiated with method <c>start3DCalibration</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer between 0 (not started) and 100 (stage completed).
    ''' </returns>
    '''/
    Public Overridable Function get_3DCalibrationStageProgress() As Integer
      Return Me._calibStageProgress
    End Function

    '''*
    ''' <summary>
    '''   Returns the latest log message from the calibration process.
    ''' <para>
    '''   When no new message is available, returns an empty string.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a character string.
    ''' </returns>
    '''/
    Public Overridable Function get_3DCalibrationLogMsg() As String
      Dim msg As String
      msg = Me._calibLogMsg
      Me._calibLogMsg = ""
      Return msg
    End Function

    '''*
    ''' <summary>
    '''   Applies the sensors tridimensional calibration parameters that have just been computed.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>  method of the module if the changes
    '''   must be kept when the device is restarted.
    ''' </para>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function save3DCalibration() As Integer
      REM // may throw an exception
      Dim shiftX As Integer = 0
      Dim shiftY As Integer = 0
      Dim shiftZ As Integer = 0
      Dim scaleExp As Integer = 0
      Dim scaleX As Integer = 0
      Dim scaleY As Integer = 0
      Dim scaleZ As Integer = 0
      Dim scaleLo As Integer = 0
      Dim scaleHi As Integer = 0
      Dim newcalib As String
      If (Me._calibProgress <> 100) Then
        Return YAPI.INVALID_ARGUMENT
      End If
      
      REM // Compute integer values (correction unit is 732ug/count)
      shiftX = -CType(Math.Round(Me._calibAccXOfs / 0.000732), Integer)
      If (shiftX < 0) Then
        shiftX = shiftX + 65536
      End If
      shiftY = -CType(Math.Round(Me._calibAccYOfs / 0.000732), Integer)
      If (shiftY < 0) Then
        shiftY = shiftY + 65536
      End If
      shiftZ = -CType(Math.Round(Me._calibAccZOfs / 0.000732), Integer)
      If (shiftZ < 0) Then
        shiftZ = shiftZ + 65536
      End If
      scaleX = CType(Math.Round(2048.0 / Me._calibAccXScale), Integer) - 2048
      scaleY = CType(Math.Round(2048.0 / Me._calibAccYScale), Integer) - 2048
      scaleZ = CType(Math.Round(2048.0 / Me._calibAccZScale), Integer) - 2048
      If (scaleX < -2048 Or scaleX >= 2048 Or scaleY < -2048 Or scaleY >= 2048 Or scaleZ < -2048 Or scaleZ >= 2048) Then
        scaleExp = 3
      Else
        If (scaleX < -1024 Or scaleX >= 1024 Or scaleY < -1024 Or scaleY >= 1024 Or scaleZ < -1024 Or scaleZ >= 1024) Then
          scaleExp = 2
        Else
          If (scaleX < -512 Or scaleX >= 512 Or scaleY < -512 Or scaleY >= 512 Or scaleZ < -512 Or scaleZ >= 512) Then
            scaleExp = 1
          Else
            scaleExp = 0
          End If
        End If
      End If
      If (scaleExp > 0) Then
        scaleX = ((scaleX) >> (scaleExp))
        scaleY = ((scaleY) >> (scaleExp))
        scaleZ = ((scaleZ) >> (scaleExp))
      End If
      If (scaleX < 0) Then
        scaleX = scaleX + 1024
      End If
      If (scaleY < 0) Then
        scaleY = scaleY + 1024
      End If
      If (scaleZ < 0) Then
        scaleZ = scaleZ + 1024
      End If
      scaleLo = ((((scaleY) And (15))) << (12)) + ((scaleX) << (2)) + scaleExp
      scaleHi = ((scaleZ) << (6)) + ((scaleY) >> (4))
      
      REM // Save calibration parameters
      newcalib = "5," + Convert.ToString( shiftX) + "," + Convert.ToString( shiftY) + "," + Convert.ToString( shiftZ) + "," + Convert.ToString( scaleLo) + "," + Convert.ToString(scaleHi)
      Me._calibStage = 0
      Return Me.set_calibrationParam(newcalib)
    End Function

    '''*
    ''' <summary>
    '''   Aborts the sensors tridimensional calibration process et restores normal settings.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function cancel3DCalibration() As Integer
      If (Me._calibStage = 0) Then
        Return YAPI.SUCCESS
      End If
      REM // may throw an exception
      Me._calibStage = 0
      Return Me.set_calibrationParam(Me._calibSavedParams)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of reference frames started using <c>yFirstRefFrame()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRefFrame</c> object, corresponding to
    '''   a reference frame currently online, or a <c>null</c> pointer
    '''   if there are no more reference frames to enumerate.
    ''' </returns>
    '''/
    Public Function nextRefFrame() As YRefFrame
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YRefFrame.FindRefFrame(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of reference frames currently accessible.
    ''' <para>
    '''   Use the method <c>YRefFrame.nextRefFrame()</c> to iterate on
    '''   next reference frames.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRefFrame</c> object, corresponding to
    '''   the first reference frame currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstRefFrame() As YRefFrame
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("RefFrame", 0, p, size, neededsize, errmsg)
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
      Return YRefFrame.FindRefFrame(serial + "." + funcId)
    End Function

    REM --- (end of YRefFrame public methods declaration)

  End Class

  REM --- (RefFrame functions)

  '''*
  ''' <summary>
  '''   Retrieves a reference frame for a given identifier.
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
  '''   This function does not require that the reference frame is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YRefFrame.isOnline()</c> to test if the reference frame is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a reference frame by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the reference frame
  ''' </param>
  ''' <returns>
  '''   a <c>YRefFrame</c> object allowing you to drive the reference frame.
  ''' </returns>
  '''/
  Public Function yFindRefFrame(ByVal func As String) As YRefFrame
    Return YRefFrame.FindRefFrame(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of reference frames currently accessible.
  ''' <para>
  '''   Use the method <c>YRefFrame.nextRefFrame()</c> to iterate on
  '''   next reference frames.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YRefFrame</c> object, corresponding to
  '''   the first reference frame currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstRefFrame() As YRefFrame
    Return YRefFrame.FirstRefFrame()
  End Function


  REM --- (end of RefFrame functions)

End Module
