' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindColorSensor(), the high-level API for ColorSensor functions
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

Module yocto_colorsensor

    REM --- (YColorSensor return codes)
    REM --- (end of YColorSensor return codes)
    REM --- (YColorSensor dlldef)
    REM --- (end of YColorSensor dlldef)
   REM --- (YColorSensor yapiwrapper)
   REM --- (end of YColorSensor yapiwrapper)
  REM --- (YColorSensor globals)

  Public Const Y_ESTIMATIONMODEL_REFLECTION As Integer = 0
  Public Const Y_ESTIMATIONMODEL_EMISSION As Integer = 1
  Public Const Y_ESTIMATIONMODEL_INVALID As Integer = -1
  Public Const Y_WORKINGMODE_AUTO As Integer = 0
  Public Const Y_WORKINGMODE_EXPERT As Integer = 1
  Public Const Y_WORKINGMODE_INVALID As Integer = -1
  Public Const Y_LEDCURRENT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_LEDCALIBRATION_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_INTEGRATIONTIME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_GAIN_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_SATURATION_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ESTIMATEDRGB_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ESTIMATEDHSL_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ESTIMATEDXYZ_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ESTIMATEDOKLAB_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_NEARRAL1_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_NEARRAL2_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_NEARRAL3_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_NEARHTMLCOLOR_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_NEARSIMPLECOLORINDEX_BROWN As Integer = 0
  Public Const Y_NEARSIMPLECOLORINDEX_RED As Integer = 1
  Public Const Y_NEARSIMPLECOLORINDEX_ORANGE As Integer = 2
  Public Const Y_NEARSIMPLECOLORINDEX_YELLOW As Integer = 3
  Public Const Y_NEARSIMPLECOLORINDEX_WHITE As Integer = 4
  Public Const Y_NEARSIMPLECOLORINDEX_GRAY As Integer = 5
  Public Const Y_NEARSIMPLECOLORINDEX_BLACK As Integer = 6
  Public Const Y_NEARSIMPLECOLORINDEX_GREEN As Integer = 7
  Public Const Y_NEARSIMPLECOLORINDEX_BLUE As Integer = 8
  Public Const Y_NEARSIMPLECOLORINDEX_PURPLE As Integer = 9
  Public Const Y_NEARSIMPLECOLORINDEX_PINK As Integer = 10
  Public Const Y_NEARSIMPLECOLORINDEX_INVALID As Integer = -1
  Public Const Y_NEARSIMPLECOLOR_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YColorSensorValueCallback(ByVal func As YColorSensor, ByVal value As String)
  Public Delegate Sub YColorSensorTimedReportCallback(ByVal func As YColorSensor, ByVal measure As YMeasure)
  REM --- (end of YColorSensor globals)

  REM --- (YColorSensor class start)

  '''*
  ''' <summary>
  '''   The <c>YColorSensor</c> class allows you to read and configure Yoctopuce color sensors.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YColorSensor
    Inherits YFunction
    REM --- (end of YColorSensor class start)

    REM --- (YColorSensor definitions)
    Public Const ESTIMATIONMODEL_REFLECTION As Integer = 0
    Public Const ESTIMATIONMODEL_EMISSION As Integer = 1
    Public Const ESTIMATIONMODEL_INVALID As Integer = -1
    Public Const WORKINGMODE_AUTO As Integer = 0
    Public Const WORKINGMODE_EXPERT As Integer = 1
    Public Const WORKINGMODE_INVALID As Integer = -1
    Public Const LEDCURRENT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const LEDCALIBRATION_INVALID As Integer = YAPI.INVALID_UINT
    Public Const INTEGRATIONTIME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const GAIN_INVALID As Integer = YAPI.INVALID_UINT
    Public Const SATURATION_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ESTIMATEDRGB_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ESTIMATEDHSL_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ESTIMATEDXYZ_INVALID As String = YAPI.INVALID_STRING
    Public Const ESTIMATEDOKLAB_INVALID As String = YAPI.INVALID_STRING
    Public Const NEARRAL1_INVALID As String = YAPI.INVALID_STRING
    Public Const NEARRAL2_INVALID As String = YAPI.INVALID_STRING
    Public Const NEARRAL3_INVALID As String = YAPI.INVALID_STRING
    Public Const NEARHTMLCOLOR_INVALID As String = YAPI.INVALID_STRING
    Public Const NEARSIMPLECOLORINDEX_BROWN As Integer = 0
    Public Const NEARSIMPLECOLORINDEX_RED As Integer = 1
    Public Const NEARSIMPLECOLORINDEX_ORANGE As Integer = 2
    Public Const NEARSIMPLECOLORINDEX_YELLOW As Integer = 3
    Public Const NEARSIMPLECOLORINDEX_WHITE As Integer = 4
    Public Const NEARSIMPLECOLORINDEX_GRAY As Integer = 5
    Public Const NEARSIMPLECOLORINDEX_BLACK As Integer = 6
    Public Const NEARSIMPLECOLORINDEX_GREEN As Integer = 7
    Public Const NEARSIMPLECOLORINDEX_BLUE As Integer = 8
    Public Const NEARSIMPLECOLORINDEX_PURPLE As Integer = 9
    Public Const NEARSIMPLECOLORINDEX_PINK As Integer = 10
    Public Const NEARSIMPLECOLORINDEX_INVALID As Integer = -1
    Public Const NEARSIMPLECOLOR_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YColorSensor definitions)

    REM --- (YColorSensor attributes declaration)
    Protected _estimationModel As Integer
    Protected _workingMode As Integer
    Protected _ledCurrent As Integer
    Protected _ledCalibration As Integer
    Protected _integrationTime As Integer
    Protected _gain As Integer
    Protected _saturation As Integer
    Protected _estimatedRGB As Integer
    Protected _estimatedHSL As Integer
    Protected _estimatedXYZ As String
    Protected _estimatedOkLab As String
    Protected _nearRAL1 As String
    Protected _nearRAL2 As String
    Protected _nearRAL3 As String
    Protected _nearHTMLColor As String
    Protected _nearSimpleColorIndex As Integer
    Protected _nearSimpleColor As String
    Protected _valueCallbackColorSensor As YColorSensorValueCallback
    REM --- (end of YColorSensor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "ColorSensor"
      REM --- (YColorSensor attributes initialization)
      _estimationModel = ESTIMATIONMODEL_INVALID
      _workingMode = WORKINGMODE_INVALID
      _ledCurrent = LEDCURRENT_INVALID
      _ledCalibration = LEDCALIBRATION_INVALID
      _integrationTime = INTEGRATIONTIME_INVALID
      _gain = GAIN_INVALID
      _saturation = SATURATION_INVALID
      _estimatedRGB = ESTIMATEDRGB_INVALID
      _estimatedHSL = ESTIMATEDHSL_INVALID
      _estimatedXYZ = ESTIMATEDXYZ_INVALID
      _estimatedOkLab = ESTIMATEDOKLAB_INVALID
      _nearRAL1 = NEARRAL1_INVALID
      _nearRAL2 = NEARRAL2_INVALID
      _nearRAL3 = NEARRAL3_INVALID
      _nearHTMLColor = NEARHTMLCOLOR_INVALID
      _nearSimpleColorIndex = NEARSIMPLECOLORINDEX_INVALID
      _nearSimpleColor = NEARSIMPLECOLOR_INVALID
      _valueCallbackColorSensor = Nothing
      REM --- (end of YColorSensor attributes initialization)
    End Sub

    REM --- (YColorSensor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("estimationModel") Then
        _estimationModel = CInt(json_val.getLong("estimationModel"))
      End If
      If json_val.has("workingMode") Then
        _workingMode = CInt(json_val.getLong("workingMode"))
      End If
      If json_val.has("ledCurrent") Then
        _ledCurrent = CInt(json_val.getLong("ledCurrent"))
      End If
      If json_val.has("ledCalibration") Then
        _ledCalibration = CInt(json_val.getLong("ledCalibration"))
      End If
      If json_val.has("integrationTime") Then
        _integrationTime = CInt(json_val.getLong("integrationTime"))
      End If
      If json_val.has("gain") Then
        _gain = CInt(json_val.getLong("gain"))
      End If
      If json_val.has("saturation") Then
        _saturation = CInt(json_val.getLong("saturation"))
      End If
      If json_val.has("estimatedRGB") Then
        _estimatedRGB = CInt(json_val.getLong("estimatedRGB"))
      End If
      If json_val.has("estimatedHSL") Then
        _estimatedHSL = CInt(json_val.getLong("estimatedHSL"))
      End If
      If json_val.has("estimatedXYZ") Then
        _estimatedXYZ = json_val.getString("estimatedXYZ")
      End If
      If json_val.has("estimatedOkLab") Then
        _estimatedOkLab = json_val.getString("estimatedOkLab")
      End If
      If json_val.has("nearRAL1") Then
        _nearRAL1 = json_val.getString("nearRAL1")
      End If
      If json_val.has("nearRAL2") Then
        _nearRAL2 = json_val.getString("nearRAL2")
      End If
      If json_val.has("nearRAL3") Then
        _nearRAL3 = json_val.getString("nearRAL3")
      End If
      If json_val.has("nearHTMLColor") Then
        _nearHTMLColor = json_val.getString("nearHTMLColor")
      End If
      If json_val.has("nearSimpleColorIndex") Then
        _nearSimpleColorIndex = CInt(json_val.getLong("nearSimpleColorIndex"))
      End If
      If json_val.has("nearSimpleColor") Then
        _nearSimpleColor = json_val.getString("nearSimpleColor")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YColorSensor private methods declaration)

    REM --- (YColorSensor public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the predictive model used for color estimation (reflective or emissive).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YColorSensor.ESTIMATIONMODEL_REFLECTION</c> or <c>YColorSensor.ESTIMATIONMODEL_EMISSION</c>,
    '''   according to the predictive model used for color estimation (reflective or emissive)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.ESTIMATIONMODEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_estimationModel() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ESTIMATIONMODEL_INVALID
        End If
      End If
      res = Me._estimationModel
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the mpredictive model to be used for color estimation (reflective or emissive).
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YColorSensor.ESTIMATIONMODEL_REFLECTION</c> or <c>YColorSensor.ESTIMATIONMODEL_EMISSION</c>,
    '''   according to the mpredictive model to be used for color estimation (reflective or emissive)
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
    Public Function set_estimationModel(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("estimationModel", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the sensor working mode.
    ''' <para>
    '''   In Auto mode, sensor parameters are automatically set based on the selected estimation model.
    '''   In Expert mode, sensor parameters such as gain and integration time are configured manually.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YColorSensor.WORKINGMODE_AUTO</c> or <c>YColorSensor.WORKINGMODE_EXPERT</c>, according to
    '''   the sensor working mode
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.WORKINGMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_workingMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return WORKINGMODE_INVALID
        End If
      End If
      res = Me._workingMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the sensor working mode.
    ''' <para>
    '''   In Auto mode, sensor parameters are automatically set based on the selected estimation model.
    '''   In Expert mode, sensor parameters such as gain and integration time are configured manually.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YColorSensor.WORKINGMODE_AUTO</c> or <c>YColorSensor.WORKINGMODE_EXPERT</c>, according to
    '''   the sensor working mode
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
    Public Function set_workingMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("workingMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the amount of current sent to the illumination LEDs, for reflection measurements.
    ''' <para>
    '''   The value is an integer ranging from 0 (LEDs off) to 254 (LEDs at maximum intensity).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the amount of current sent to the illumination LEDs, for reflection measurements
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.LEDCURRENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ledCurrent() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LEDCURRENT_INVALID
        End If
      End If
      res = Me._ledCurrent
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the amount of current sent to the illumination LEDs, for reflection measurements.
    ''' <para>
    '''   The value is an integer ranging from 0 (LEDs off) to 254 (LEDs at maximum intensity).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the amount of current sent to the illumination LEDs, for reflection measurements
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
    Public Function set_ledCurrent(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("ledCurrent", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current sent to the illumination LEDs during the last calibration.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current sent to the illumination LEDs during the last calibration
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.LEDCALIBRATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ledCalibration() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LEDCALIBRATION_INVALID
        End If
      End If
      res = Me._ledCalibration
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Remember the LED current sent to the illumination LEDs during a calibration.
    ''' <para>
    '''   Thanks to this, the device will be able to use the same current during measurements.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
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
    Public Function set_ledCalibration(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("ledCalibration", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current integration time for spectral measurement, in milliseconds.
    ''' <para>
    '''   A longer integration time increase the sensitivity for low light conditions,
    '''   but reduces the measurement rate and may lead to saturation for lighter colors.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current integration time for spectral measurement, in milliseconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.INTEGRATIONTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_integrationTime() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return INTEGRATIONTIME_INVALID
        End If
      End If
      res = Me._integrationTime
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the integration time for spectral measurement, in milliseconds.
    ''' <para>
    '''   A longer integration time increase the sensitivity for low light conditions,
    '''   but reduces the measurement rate and may lead to saturation for lighter colors.
    '''   This method can only be used when the sensor is configured in expert mode;
    '''   when running in auto mode, the change will be ignored.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the integration time for spectral measurement, in milliseconds
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
    Public Function set_integrationTime(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("integrationTime", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current spectral channel detector gain exponent.
    ''' <para>
    '''   For a value <c>n</c> ranging from 0 to 12, the applied gain is 2^(n-1).
    '''   0 corresponds to a gain of 0.5, and 12 corresponds to a gain of 2048.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current spectral channel detector gain exponent
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.GAIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_gain() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return GAIN_INVALID
        End If
      End If
      res = Me._gain
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the spectral channel detector gain exponent.
    ''' <para>
    '''   For a value <c>n</c> ranging from 0 to 12, the applied gain is 2^(n-1).
    '''   0 corresponds to a gain of 0.5, and 12 corresponds to a gain of 2048.
    '''   This method can only be used when the sensor is configured in expert mode;
    '''   when running in auto mode, the change will be ignored.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the spectral channel detector gain exponent
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
    Public Function set_gain(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("gain", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current saturation state of the sensor, as an integer.
    ''' <para>
    '''   Bit 0 indicates saturation of the analog sensor, which can only
    '''   be corrected by reducing the gain parameters or the luminosity.
    '''   Bit 1 indicates saturation of the digital interface, which can
    '''   be corrected by reducing the integration time or the gain.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current saturation state of the sensor, as an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.SATURATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_saturation() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SATURATION_INVALID
        End If
      End If
      res = Me._saturation
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the estimated color in RGB color model (0xRRGGBB).
    ''' <para>
    '''   The RGB color model describes each color using a combination of 3 components:
    ''' </para>
    ''' <para>
    '''   - Red (R): the intensity of red, in thee range 0...255
    ''' </para>
    ''' <para>
    '''   - Green (G): the intensity of green, in thee range 0...255
    ''' </para>
    ''' <para>
    '''   - Blue (B): the intensity of blue, in thee range 0...255
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the estimated color in RGB color model (0xRRGGBB)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.ESTIMATEDRGB_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_estimatedRGB() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ESTIMATEDRGB_INVALID
        End If
      End If
      res = Me._estimatedRGB
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the estimated color in HSL color model (0xHHSSLL).
    ''' <para>
    '''   The HSL color model describes each color using a combination of 3 components:
    ''' </para>
    ''' <para>
    '''   - Hue (H): the angle on the color wheel (0-360 degrees), mapped to 0...255
    ''' </para>
    ''' <para>
    '''   - Saturation (S): the intensity of the color (0-100%), mapped to 0...255
    ''' </para>
    ''' <para>
    '''   - Lightness (L): the brightness of the color (0-100%), mapped to 0...255
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the estimated color in HSL color model (0xHHSSLL)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.ESTIMATEDHSL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_estimatedHSL() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ESTIMATEDHSL_INVALID
        End If
      End If
      res = Me._estimatedHSL
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the estimated color according to the CIE XYZ color model.
    ''' <para>
    '''   This color model is based on human vision and light perception, with three components
    '''   represented by real numbers between 0 and 1:
    ''' </para>
    ''' <para>
    '''   - X: corresponds to a component mixing sensitivity to red and green
    ''' </para>
    ''' <para>
    '''   - Y: represents luminance (perceived brightness)
    ''' </para>
    ''' <para>
    '''   - Z: corresponds to sensitivity to blue
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the estimated color according to the CIE XYZ color model
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.ESTIMATEDXYZ_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_estimatedXYZ() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ESTIMATEDXYZ_INVALID
        End If
      End If
      res = Me._estimatedXYZ
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the estimated color according to the OkLab color model.
    ''' <para>
    '''   OkLab is a perceptual color model that aims to align human color perception with numerical
    '''   values, so that visually near colors are also numerically near. Colors are represented using three components:
    ''' </para>
    ''' <para>
    '''   - L: lightness, a real number between 0 and 1
    '''   -
    ''' </para>
    ''' <para>
    '''   - a: color variations between green and red, between -0.5 and 0.5
    '''   -
    ''' </para>
    ''' <para>
    '''   - b: color variations between blue and yellow, between -0.5 and 0.5.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the estimated color according to the OkLab color model
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.ESTIMATEDOKLAB_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_estimatedOkLab() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ESTIMATEDOKLAB_INVALID
        End If
      End If
      res = Me._estimatedOkLab
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the RAL Classic color closest to the estimated color, with a similarity ratio.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the RAL Classic color closest to the estimated color, with a similarity ratio
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.NEARRAL1_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nearRAL1() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NEARRAL1_INVALID
        End If
      End If
      res = Me._nearRAL1
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the second closest RAL Classic color to the estimated color, with a similarity ratio.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the second closest RAL Classic color to the estimated color, with a similarity ratio
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.NEARRAL2_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nearRAL2() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NEARRAL2_INVALID
        End If
      End If
      res = Me._nearRAL2
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the third closest RAL Classic color to the estimated color, with a similarity ratio.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the third closest RAL Classic color to the estimated color, with a similarity ratio
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.NEARRAL3_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nearRAL3() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NEARRAL3_INVALID
        End If
      End If
      res = Me._nearRAL3
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the name of the HTML color closest to the estimated color.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of the HTML color closest to the estimated color
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.NEARHTMLCOLOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nearHTMLColor() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NEARHTMLCOLOR_INVALID
        End If
      End If
      res = Me._nearHTMLColor
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the index of the basic color typically used to refer to the estimated color (enumerated value).
    ''' <para>
    '''   The list of basic colors recognized is:
    ''' </para>
    ''' <para>
    '''   - 0 - Brown
    ''' </para>
    ''' <para>
    '''   - 1 - Red
    ''' </para>
    ''' <para>
    '''   - 2 - Orange
    ''' </para>
    ''' <para>
    '''   - 3 - Yellow
    ''' </para>
    ''' <para>
    '''   - 4 - White
    ''' </para>
    ''' <para>
    '''   - 5 - Gray
    ''' </para>
    ''' <para>
    '''   - 6 - Black
    ''' </para>
    ''' <para>
    '''   - 7 - Green
    ''' </para>
    ''' <para>
    '''   - 8 - Blue
    ''' </para>
    ''' <para>
    '''   - 9 - Purple
    ''' </para>
    ''' <para>
    '''   - 10 - Pink
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YColorSensor.NEARSIMPLECOLORINDEX_BROWN</c>,
    '''   <c>YColorSensor.NEARSIMPLECOLORINDEX_RED</c>, <c>YColorSensor.NEARSIMPLECOLORINDEX_ORANGE</c>,
    '''   <c>YColorSensor.NEARSIMPLECOLORINDEX_YELLOW</c>, <c>YColorSensor.NEARSIMPLECOLORINDEX_WHITE</c>,
    '''   <c>YColorSensor.NEARSIMPLECOLORINDEX_GRAY</c>, <c>YColorSensor.NEARSIMPLECOLORINDEX_BLACK</c>,
    '''   <c>YColorSensor.NEARSIMPLECOLORINDEX_GREEN</c>, <c>YColorSensor.NEARSIMPLECOLORINDEX_BLUE</c>,
    '''   <c>YColorSensor.NEARSIMPLECOLORINDEX_PURPLE</c> and <c>YColorSensor.NEARSIMPLECOLORINDEX_PINK</c>
    '''   corresponding to the index of the basic color typically used to refer to the estimated color (enumerated value)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.NEARSIMPLECOLORINDEX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nearSimpleColorIndex() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NEARSIMPLECOLORINDEX_INVALID
        End If
      End If
      res = Me._nearSimpleColorIndex
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the name of the basic color typically used to refer to the estimated color.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of the basic color typically used to refer to the estimated color
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YColorSensor.NEARSIMPLECOLOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nearSimpleColor() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NEARSIMPLECOLOR_INVALID
        End If
      End If
      res = Me._nearSimpleColor
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a color sensor for a given identifier.
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
    '''   This function does not require that the color sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YColorSensor.isOnline()</c> to test if the color sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a color sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the color sensor, for instance
    '''   <c>MyDevice.colorSensor</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YColorSensor</c> object allowing you to drive the color sensor.
    ''' </returns>
    '''/
    Public Shared Function FindColorSensor(func As String) As YColorSensor
      Dim obj As YColorSensor
      obj = CType(YFunction._FindFromCache("ColorSensor", func), YColorSensor)
      If ((obj Is Nothing)) Then
        obj = New YColorSensor(func)
        YFunction._AddToCache("ColorSensor", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YColorSensorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackColorSensor = callback
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
      If (Not (Me._valueCallbackColorSensor Is Nothing)) Then
        Me._valueCallbackColorSensor(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Turns on the built-in illumination LEDs using the same current as used during last calibration.
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function turnLedOn() As Integer
      Return Me.set_ledCurrent(Me.get_ledCalibration())
    End Function

    '''*
    ''' <summary>
    '''   Turns off the built-in illumination LEDs.
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function turnLedOff() As Integer
      Return Me.set_ledCurrent(0)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of color sensors started using <c>yFirstColorSensor()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned color sensors order.
    '''   If you want to find a specific a color sensor, use <c>ColorSensor.findColorSensor()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YColorSensor</c> object, corresponding to
    '''   a color sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more color sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextColorSensor() As YColorSensor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YColorSensor.FindColorSensor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of color sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YColorSensor.nextColorSensor()</c> to iterate on
    '''   next color sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YColorSensor</c> object, corresponding to
    '''   the first color sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstColorSensor() As YColorSensor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("ColorSensor", 0, p, size, neededsize, errmsg)
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
      Return YColorSensor.FindColorSensor(serial + "." + funcId)
    End Function

    REM --- (end of YColorSensor public methods declaration)

  End Class

  REM --- (YColorSensor functions)

  '''*
  ''' <summary>
  '''   Retrieves a color sensor for a given identifier.
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
  '''   This function does not require that the color sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YColorSensor.isOnline()</c> to test if the color sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a color sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the color sensor, for instance
  '''   <c>MyDevice.colorSensor</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YColorSensor</c> object allowing you to drive the color sensor.
  ''' </returns>
  '''/
  Public Function yFindColorSensor(ByVal func As String) As YColorSensor
    Return YColorSensor.FindColorSensor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of color sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YColorSensor.nextColorSensor()</c> to iterate on
  '''   next color sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YColorSensor</c> object, corresponding to
  '''   the first color sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstColorSensor() As YColorSensor
    Return YColorSensor.FirstColorSensor()
  End Function


  REM --- (end of YColorSensor functions)

End Module
