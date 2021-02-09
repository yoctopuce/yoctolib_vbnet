' ********************************************************************
'
'  $Id: yocto_multicellweighscale.vb 43580 2021-01-26 17:46:01Z mvuilleu $
'
'  Implements yFindMultiCellWeighScale(), the high-level API for MultiCellWeighScale functions
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

Module yocto_multicellweighscale

    REM --- (YMultiCellWeighScale return codes)
    REM --- (end of YMultiCellWeighScale return codes)
    REM --- (YMultiCellWeighScale dlldef)
    REM --- (end of YMultiCellWeighScale dlldef)
   REM --- (YMultiCellWeighScale yapiwrapper)
   REM --- (end of YMultiCellWeighScale yapiwrapper)
  REM --- (YMultiCellWeighScale globals)

  Public Const Y_CELLCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_EXTERNALSENSE_FALSE As Integer = 0
  Public Const Y_EXTERNALSENSE_TRUE As Integer = 1
  Public Const Y_EXTERNALSENSE_INVALID As Integer = -1
  Public Const Y_EXCITATION_OFF As Integer = 0
  Public Const Y_EXCITATION_DC As Integer = 1
  Public Const Y_EXCITATION_AC As Integer = 2
  Public Const Y_EXCITATION_INVALID As Integer = -1
  Public Const Y_TEMPAVGADAPTRATIO_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_TEMPCHGADAPTRATIO_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_COMPTEMPAVG_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_COMPTEMPCHG_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_COMPENSATION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_ZEROTRACKING_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YMultiCellWeighScaleValueCallback(ByVal func As YMultiCellWeighScale, ByVal value As String)
  Public Delegate Sub YMultiCellWeighScaleTimedReportCallback(ByVal func As YMultiCellWeighScale, ByVal measure As YMeasure)
  REM --- (end of YMultiCellWeighScale globals)

  REM --- (YMultiCellWeighScale class start)

  '''*
  ''' <summary>
  '''   The <c>YMultiCellWeighScale</c> class provides a weight measurement from a set of ratiometric
  '''   sensors.
  ''' <para>
  '''   It can be used to control the bridge excitation parameters, in order to avoid
  '''   measure shifts caused by temperature variation in the electronics, and can also
  '''   automatically apply an additional correction factor based on temperature to
  '''   compensate for offsets in the load cells themselves.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMultiCellWeighScale
    Inherits YSensor
    REM --- (end of YMultiCellWeighScale class start)

    REM --- (YMultiCellWeighScale definitions)
    Public Const CELLCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const EXTERNALSENSE_FALSE As Integer = 0
    Public Const EXTERNALSENSE_TRUE As Integer = 1
    Public Const EXTERNALSENSE_INVALID As Integer = -1
    Public Const EXCITATION_OFF As Integer = 0
    Public Const EXCITATION_DC As Integer = 1
    Public Const EXCITATION_AC As Integer = 2
    Public Const EXCITATION_INVALID As Integer = -1
    Public Const TEMPAVGADAPTRATIO_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const TEMPCHGADAPTRATIO_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const COMPTEMPAVG_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const COMPTEMPCHG_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const COMPENSATION_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const ZEROTRACKING_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YMultiCellWeighScale definitions)

    REM --- (YMultiCellWeighScale attributes declaration)
    Protected _cellCount As Integer
    Protected _externalSense As Integer
    Protected _excitation As Integer
    Protected _tempAvgAdaptRatio As Double
    Protected _tempChgAdaptRatio As Double
    Protected _compTempAvg As Double
    Protected _compTempChg As Double
    Protected _compensation As Double
    Protected _zeroTracking As Double
    Protected _command As String
    Protected _valueCallbackMultiCellWeighScale As YMultiCellWeighScaleValueCallback
    Protected _timedReportCallbackMultiCellWeighScale As YMultiCellWeighScaleTimedReportCallback
    REM --- (end of YMultiCellWeighScale attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "MultiCellWeighScale"
      REM --- (YMultiCellWeighScale attributes initialization)
      _cellCount = CELLCOUNT_INVALID
      _externalSense = EXTERNALSENSE_INVALID
      _excitation = EXCITATION_INVALID
      _tempAvgAdaptRatio = TEMPAVGADAPTRATIO_INVALID
      _tempChgAdaptRatio = TEMPCHGADAPTRATIO_INVALID
      _compTempAvg = COMPTEMPAVG_INVALID
      _compTempChg = COMPTEMPCHG_INVALID
      _compensation = COMPENSATION_INVALID
      _zeroTracking = ZEROTRACKING_INVALID
      _command = COMMAND_INVALID
      _valueCallbackMultiCellWeighScale = Nothing
      _timedReportCallbackMultiCellWeighScale = Nothing
      REM --- (end of YMultiCellWeighScale attributes initialization)
    End Sub

    REM --- (YMultiCellWeighScale private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("cellCount") Then
        _cellCount = CInt(json_val.getLong("cellCount"))
      End If
      If json_val.has("externalSense") Then
        If (json_val.getInt("externalSense") > 0) Then _externalSense = 1 Else _externalSense = 0
      End If
      If json_val.has("excitation") Then
        _excitation = CInt(json_val.getLong("excitation"))
      End If
      If json_val.has("tempAvgAdaptRatio") Then
        _tempAvgAdaptRatio = Math.Round(json_val.getDouble("tempAvgAdaptRatio") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("tempChgAdaptRatio") Then
        _tempChgAdaptRatio = Math.Round(json_val.getDouble("tempChgAdaptRatio") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("compTempAvg") Then
        _compTempAvg = Math.Round(json_val.getDouble("compTempAvg") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("compTempChg") Then
        _compTempChg = Math.Round(json_val.getDouble("compTempChg") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("compensation") Then
        _compensation = Math.Round(json_val.getDouble("compensation") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("zeroTracking") Then
        _zeroTracking = Math.Round(json_val.getDouble("zeroTracking") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YMultiCellWeighScale private methods declaration)

    REM --- (YMultiCellWeighScale public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the measuring unit for the weight.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the measuring unit for the weight
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
    Public Function set_unit(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("unit", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of load cells in use.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of load cells in use
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.CELLCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_cellCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CELLCOUNT_INVALID
        End If
      End If
      res = Me._cellCount
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of load cells in use.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the number of load cells in use
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
    Public Function set_cellCount(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("cellCount", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns true if entry 4 is used as external sense for 6-wires load cells.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YMultiCellWeighScale.EXTERNALSENSE_FALSE</c> or <c>YMultiCellWeighScale.EXTERNALSENSE_TRUE</c>,
    '''   according to true if entry 4 is used as external sense for 6-wires load cells
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.EXTERNALSENSE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_externalSense() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return EXTERNALSENSE_INVALID
        End If
      End If
      res = Me._externalSense
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the configuration to tell if entry 4 is used as external sense for
    '''   6-wires load cells.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the
    '''   module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YMultiCellWeighScale.EXTERNALSENSE_FALSE</c> or <c>YMultiCellWeighScale.EXTERNALSENSE_TRUE</c>,
    '''   according to the configuration to tell if entry 4 is used as external sense for
    '''   6-wires load cells
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
    Public Function set_externalSense(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("externalSense", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current load cell bridge excitation method.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YMultiCellWeighScale.EXCITATION_OFF</c>, <c>YMultiCellWeighScale.EXCITATION_DC</c>
    '''   and <c>YMultiCellWeighScale.EXCITATION_AC</c> corresponding to the current load cell bridge excitation method
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.EXCITATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_excitation() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return EXCITATION_INVALID
        End If
      End If
      res = Me._excitation
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current load cell bridge excitation method.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YMultiCellWeighScale.EXCITATION_OFF</c>, <c>YMultiCellWeighScale.EXCITATION_DC</c>
    '''   and <c>YMultiCellWeighScale.EXCITATION_AC</c> corresponding to the current load cell bridge excitation method
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
    Public Function set_excitation(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("excitation", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the averaged temperature update rate, in per mille.
    ''' <para>
    '''   The purpose of this adaptation ratio is to model the thermal inertia of the load cell.
    '''   The averaged temperature is updated every 10 seconds, by applying this adaptation rate
    '''   to the difference between the measures ambient temperature and the current compensation
    '''   temperature. The standard rate is 0.2 per mille, and the maximal rate is 65 per mille.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the averaged temperature update rate, in per mille
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
    Public Function set_tempAvgAdaptRatio(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("tempAvgAdaptRatio", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the averaged temperature update rate, in per mille.
    ''' <para>
    '''   The purpose of this adaptation ratio is to model the thermal inertia of the load cell.
    '''   The averaged temperature is updated every 10 seconds, by applying this adaptation rate
    '''   to the difference between the measures ambient temperature and the current compensation
    '''   temperature. The standard rate is 0.2 per mille, and the maximal rate is 65 per mille.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the averaged temperature update rate, in per mille
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.TEMPAVGADAPTRATIO_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_tempAvgAdaptRatio() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TEMPAVGADAPTRATIO_INVALID
        End If
      End If
      res = Me._tempAvgAdaptRatio
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the temperature change update rate, in per mille.
    ''' <para>
    '''   The temperature change is updated every 10 seconds, by applying this adaptation rate
    '''   to the difference between the measures ambient temperature and the current temperature used for
    '''   change compensation. The standard rate is 0.6 per mille, and the maximal rate is 65 per mille.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the temperature change update rate, in per mille
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
    Public Function set_tempChgAdaptRatio(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("tempChgAdaptRatio", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the temperature change update rate, in per mille.
    ''' <para>
    '''   The temperature change is updated every 10 seconds, by applying this adaptation rate
    '''   to the difference between the measures ambient temperature and the current temperature used for
    '''   change compensation. The standard rate is 0.6 per mille, and the maximal rate is 65 per mille.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the temperature change update rate, in per mille
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.TEMPCHGADAPTRATIO_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_tempChgAdaptRatio() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TEMPCHGADAPTRATIO_INVALID
        End If
      End If
      res = Me._tempChgAdaptRatio
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current averaged temperature, used for thermal compensation.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current averaged temperature, used for thermal compensation
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.COMPTEMPAVG_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_compTempAvg() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COMPTEMPAVG_INVALID
        End If
      End If
      res = Me._compTempAvg
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current temperature variation, used for thermal compensation.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current temperature variation, used for thermal compensation
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.COMPTEMPCHG_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_compTempChg() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COMPTEMPCHG_INVALID
        End If
      End If
      res = Me._compTempChg
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current current thermal compensation value.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current current thermal compensation value
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.COMPENSATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_compensation() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COMPENSATION_INVALID
        End If
      End If
      res = Me._compensation
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the zero tracking threshold value.
    ''' <para>
    '''   When this threshold is larger than
    '''   zero, any measure under the threshold will automatically be ignored and the
    '''   zero compensation will be updated.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the zero tracking threshold value
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
    Public Function set_zeroTracking(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("zeroTracking", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the zero tracking threshold value.
    ''' <para>
    '''   When this threshold is larger than
    '''   zero, any measure under the threshold will automatically be ignored and the
    '''   zero compensation will be updated.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the zero tracking threshold value
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMultiCellWeighScale.ZEROTRACKING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_zeroTracking() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ZEROTRACKING_INVALID
        End If
      End If
      res = Me._zeroTracking
      Return res
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
    '''   Retrieves a multi-cell weighing scale sensor for a given identifier.
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
    '''   This function does not require that the multi-cell weighing scale sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YMultiCellWeighScale.isOnline()</c> to test if the multi-cell weighing scale sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a multi-cell weighing scale sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the multi-cell weighing scale sensor, for instance
    '''   <c>YWMBRDG1.multiCellWeighScale</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YMultiCellWeighScale</c> object allowing you to drive the multi-cell weighing scale sensor.
    ''' </returns>
    '''/
    Public Shared Function FindMultiCellWeighScale(func As String) As YMultiCellWeighScale
      Dim obj As YMultiCellWeighScale
      obj = CType(YFunction._FindFromCache("MultiCellWeighScale", func), YMultiCellWeighScale)
      If ((obj Is Nothing)) Then
        obj = New YMultiCellWeighScale(func)
        YFunction._AddToCache("MultiCellWeighScale", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YMultiCellWeighScaleValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackMultiCellWeighScale = callback
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
      If (Not (Me._valueCallbackMultiCellWeighScale Is Nothing)) Then
        Me._valueCallbackMultiCellWeighScale(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every periodic timed notification.
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
    '''   arguments: the function object of which the value has changed, and an <c>YMeasure</c> object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YMultiCellWeighScaleTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackMultiCellWeighScale = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackMultiCellWeighScale Is Nothing)) Then
        Me._timedReportCallbackMultiCellWeighScale(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Adapts the load cell signal bias (stored in the corresponding genericSensor)
    '''   so that the current signal corresponds to a zero weight.
    ''' <para>
    '''   Remember to call the
    '''   <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function tare() As Integer
      Return Me.set_command("T")
    End Function

    '''*
    ''' <summary>
    '''   Configures the load cells span parameters (stored in the corresponding genericSensors)
    '''   so that the current signal corresponds to the specified reference weight.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="currWeight">
    '''   reference weight presently on the load cell.
    ''' </param>
    ''' <param name="maxWeight">
    '''   maximum weight to be expected on the load cell.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function setupSpan(currWeight As Double, maxWeight As Double) As Integer
      Return Me.set_command("S" + Convert.ToString( CType(Math.Round(1000*currWeight), Integer)) + ":" + Convert.ToString(CType(Math.Round(1000*maxWeight), Integer)))
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of multi-cell weighing scale sensors started using <c>yFirstMultiCellWeighScale()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned multi-cell weighing scale sensors order.
    '''   If you want to find a specific a multi-cell weighing scale sensor, use
    '''   <c>MultiCellWeighScale.findMultiCellWeighScale()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMultiCellWeighScale</c> object, corresponding to
    '''   a multi-cell weighing scale sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more multi-cell weighing scale sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextMultiCellWeighScale() As YMultiCellWeighScale
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YMultiCellWeighScale.FindMultiCellWeighScale(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of multi-cell weighing scale sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YMultiCellWeighScale.nextMultiCellWeighScale()</c> to iterate on
    '''   next multi-cell weighing scale sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMultiCellWeighScale</c> object, corresponding to
    '''   the first multi-cell weighing scale sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstMultiCellWeighScale() As YMultiCellWeighScale
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("MultiCellWeighScale", 0, p, size, neededsize, errmsg)
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
      Return YMultiCellWeighScale.FindMultiCellWeighScale(serial + "." + funcId)
    End Function

    REM --- (end of YMultiCellWeighScale public methods declaration)

  End Class

  REM --- (YMultiCellWeighScale functions)

  '''*
  ''' <summary>
  '''   Retrieves a multi-cell weighing scale sensor for a given identifier.
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
  '''   This function does not require that the multi-cell weighing scale sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YMultiCellWeighScale.isOnline()</c> to test if the multi-cell weighing scale sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a multi-cell weighing scale sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the multi-cell weighing scale sensor, for instance
  '''   <c>YWMBRDG1.multiCellWeighScale</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YMultiCellWeighScale</c> object allowing you to drive the multi-cell weighing scale sensor.
  ''' </returns>
  '''/
  Public Function yFindMultiCellWeighScale(ByVal func As String) As YMultiCellWeighScale
    Return YMultiCellWeighScale.FindMultiCellWeighScale(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of multi-cell weighing scale sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YMultiCellWeighScale.nextMultiCellWeighScale()</c> to iterate on
  '''   next multi-cell weighing scale sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YMultiCellWeighScale</c> object, corresponding to
  '''   the first multi-cell weighing scale sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMultiCellWeighScale() As YMultiCellWeighScale
    Return YMultiCellWeighScale.FirstMultiCellWeighScale()
  End Function


  REM --- (end of YMultiCellWeighScale functions)

End Module
