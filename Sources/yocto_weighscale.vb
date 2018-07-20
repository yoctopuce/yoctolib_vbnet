'*********************************************************************
'*
'* $Id: yocto_weighscale.vb 31016 2018-06-04 08:45:40Z mvuilleu $
'*
'* Implements yFindWeighScale(), the high-level API for WeighScale functions
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

Module yocto_weighscale

    REM --- (YWeighScale return codes)
    REM --- (end of YWeighScale return codes)
    REM --- (YWeighScale dlldef)
    REM --- (end of YWeighScale dlldef)
  REM --- (YWeighScale globals)

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
  Public Delegate Sub YWeighScaleValueCallback(ByVal func As YWeighScale, ByVal value As String)
  Public Delegate Sub YWeighScaleTimedReportCallback(ByVal func As YWeighScale, ByVal measure As YMeasure)
  REM --- (end of YWeighScale globals)

  REM --- (YWeighScale class start)

  '''*
  ''' <summary>
  '''   The YWeighScale class provides a weight measurement from a ratiometric load cell
  '''   sensor.
  ''' <para>
  '''   It can be used to control the bridge excitation parameters, in order to avoid
  '''   measure shifts caused by temperature variation in the electronics, and can also
  '''   automatically apply an additional correction factor based on temperature to
  '''   compensate for offsets in the load cell itself.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YWeighScale
    Inherits YSensor
    REM --- (end of YWeighScale class start)

    REM --- (YWeighScale definitions)
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
    REM --- (end of YWeighScale definitions)

    REM --- (YWeighScale attributes declaration)
    Protected _excitation As Integer
    Protected _tempAvgAdaptRatio As Double
    Protected _tempChgAdaptRatio As Double
    Protected _compTempAvg As Double
    Protected _compTempChg As Double
    Protected _compensation As Double
    Protected _zeroTracking As Double
    Protected _command As String
    Protected _valueCallbackWeighScale As YWeighScaleValueCallback
    Protected _timedReportCallbackWeighScale As YWeighScaleTimedReportCallback
    REM --- (end of YWeighScale attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "WeighScale"
      REM --- (YWeighScale attributes initialization)
      _excitation = EXCITATION_INVALID
      _tempAvgAdaptRatio = TEMPAVGADAPTRATIO_INVALID
      _tempChgAdaptRatio = TEMPCHGADAPTRATIO_INVALID
      _compTempAvg = COMPTEMPAVG_INVALID
      _compTempChg = COMPTEMPCHG_INVALID
      _compensation = COMPENSATION_INVALID
      _zeroTracking = ZEROTRACKING_INVALID
      _command = COMMAND_INVALID
      _valueCallbackWeighScale = Nothing
      _timedReportCallbackWeighScale = Nothing
      REM --- (end of YWeighScale attributes initialization)
    End Sub

    REM --- (YWeighScale private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
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

    REM --- (end of YWeighScale private methods declaration)

    REM --- (YWeighScale public methods declaration)

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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   Returns the current load cell bridge excitation method.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_EXCITATION_OFF</c>, <c>Y_EXCITATION_DC</c> and <c>Y_EXCITATION_AC</c>
    '''   corresponding to the current load cell bridge excitation method
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_EXCITATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_excitation() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_EXCITATION_OFF</c>, <c>Y_EXCITATION_DC</c> and <c>Y_EXCITATION_AC</c>
    '''   corresponding to the current load cell bridge excitation method
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
    '''   to the difference between the measures ambiant temperature and the current compensation
    '''   temperature. The standard rate is 0.2 per mille, and the maximal rate is 65 per mille.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   to the difference between the measures ambiant temperature and the current compensation
    '''   temperature. The standard rate is 0.2 per mille, and the maximal rate is 65 per mille.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the averaged temperature update rate, in per mille
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_TEMPAVGADAPTRATIO_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_tempAvgAdaptRatio() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   to the difference between the measures ambiant temperature and the current temperature used for
    '''   change compensation. The standard rate is 0.6 per mille, and the maximal rate is 65 pour mille.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   to the difference between the measures ambiant temperature and the current temperature used for
    '''   change compensation. The standard rate is 0.6 per mille, and the maximal rate is 65 pour mille.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the temperature change update rate, in per mille
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_TEMPCHGADAPTRATIO_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_tempChgAdaptRatio() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_COMPTEMPAVG_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_compTempAvg() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_COMPTEMPCHG_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_compTempChg() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_COMPENSATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_compensation() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   On failure, throws an exception or returns <c>Y_ZEROTRACKING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_zeroTracking() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ZEROTRACKING_INVALID
        End If
      End If
      res = Me._zeroTracking
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
    '''   Retrieves a weighing scale sensor for a given identifier.
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
    '''   This function does not require that the weighing scale sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YWeighScale.isOnline()</c> to test if the weighing scale sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a weighing scale sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the weighing scale sensor
    ''' </param>
    ''' <returns>
    '''   a <c>YWeighScale</c> object allowing you to drive the weighing scale sensor.
    ''' </returns>
    '''/
    Public Shared Function FindWeighScale(func As String) As YWeighScale
      Dim obj As YWeighScale
      obj = CType(YFunction._FindFromCache("WeighScale", func), YWeighScale)
      If ((obj Is Nothing)) Then
        obj = New YWeighScale(func)
        YFunction._AddToCache("WeighScale", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YWeighScaleValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackWeighScale = callback
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
      If (Not (Me._valueCallbackWeighScale Is Nothing)) Then
        Me._valueCallbackWeighScale(Me, value)
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
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YWeighScaleTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackWeighScale = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackWeighScale Is Nothing)) Then
        Me._timedReportCallbackWeighScale(Me, value)
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
    ''' </para>
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
    Public Overridable Function tare() As Integer
      Return Me.set_command("T")
    End Function

    '''*
    ''' <summary>
    '''   Configures the load cell span parameters (stored in the corresponding genericSensor)
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
    '''   maximum weight to be expectect on the load cell.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function setupSpan(currWeight As Double, maxWeight As Double) As Integer
      Return Me.set_command("S" + Convert.ToString( CType(Math.Round(1000*currWeight), Integer)) + ":" + Convert.ToString(CType(Math.Round(1000*maxWeight), Integer)))
    End Function

    Public Overridable Function setCompensationTable(tableIndex As Integer, tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Dim siz As Integer = 0
      Dim res As Integer = 0
      Dim idx As Integer = 0
      Dim found As Integer = 0
      Dim prev As Double = 0
      Dim curr As Double = 0
      Dim currComp As Double = 0
      Dim idxTemp As Double = 0
      siz = tempValues.Count
      If Not(siz <> 1) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "thermal compensation table must have at least two points")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not(siz = compValues.Count) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "table sizes mismatch")
        return YAPI.INVALID_ARGUMENT
      end if

      res = Me.set_command("" + Convert.ToString(tableIndex) + "Z")
      If Not(res=YAPI.SUCCESS) Then
        me._throw( YAPI.IO_ERROR,  "unable to reset thermal compensation table")
        return YAPI.IO_ERROR
      end if
      REM // add records in growing temperature value
      found = 1
      prev = -999999.0
      While (found > 0)
        found = 0
        curr = 99999999.0
        currComp = -999999.0
        idx = 0
        While (idx < siz)
          idxTemp = tempValues(idx)
          If ((idxTemp > prev) AndAlso (idxTemp < curr)) Then
            curr = idxTemp
            currComp = compValues(idx)
            found = 1
          End If
          idx = idx + 1
        End While
        If (found > 0) Then
          res = Me.set_command("" + Convert.ToString( tableIndex) + "m" + Convert.ToString( CType(Math.Round(1000*curr), Integer)) + ":" + Convert.ToString(CType(Math.Round(1000*currComp), Integer)))
          If Not(res=YAPI.SUCCESS) Then
            me._throw( YAPI.IO_ERROR,  "unable to set thermal compensation table")
            return YAPI.IO_ERROR
          end if
          prev = curr
        End If
      End While
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function loadCompensationTable(tableIndex As Integer, tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Dim id As String
      Dim bin_json As Byte()
      Dim paramlist As List(Of String) = New List(Of String)()
      Dim siz As Integer = 0
      Dim idx As Integer = 0
      Dim temp As Double = 0
      Dim comp As Double = 0

      id = Me.get_functionId()
      id = (id).Substring( 10, (id).Length - 10)
      bin_json = Me._download("extra.json?page=" + Convert.ToString((4*YAPI._atoi(id))+tableIndex))
      paramlist = Me._json_get_array(bin_json)
      REM // convert all values to float and append records
      siz = ((paramlist.Count) >> (1))
      tempValues.Clear()
      compValues.Clear()
      idx = 0
      While (idx < siz)
        temp = Double.Parse(paramlist(2*idx))/1000.0
        comp = Double.Parse(paramlist(2*idx+1))/1000.0
        tempValues.Add(temp)
        compValues.Add(comp)
        idx = idx + 1
      End While


      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Records a weight offset thermal compensation table, in order to automatically correct the
    '''   measured weight based on the averaged compensation temperature.
    ''' <para>
    '''   The weight correction will be applied by linear interpolation between specified points.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, corresponding to all averaged
    '''   temperatures for which an offset correction is specified.
    ''' </param>
    ''' <param name="compValues">
    '''   array of floating point numbers, corresponding to the offset correction
    '''   to apply for each of the temperature included in the first
    '''   argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_offsetAvgCompensationTable(tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Return Me.setCompensationTable(0, tempValues, compValues)
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the weight offset thermal compensation table previously configured using the
    '''   <c>set_offsetAvgCompensationTable</c> function.
    ''' <para>
    '''   The weight correction is applied by linear interpolation between specified points.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with all averaged temperatures for which an offset correction is specified.
    ''' </param>
    ''' <param name="compValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with the offset correction applied for each of the temperature
    '''   included in the first argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadOffsetAvgCompensationTable(tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Return Me.loadCompensationTable(0, tempValues, compValues)
    End Function

    '''*
    ''' <summary>
    '''   Records a weight offset thermal compensation table, in order to automatically correct the
    '''   measured weight based on the variation of temperature.
    ''' <para>
    '''   The weight correction will be applied by linear interpolation between specified points.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, corresponding to temperature
    '''   variations for which an offset correction is specified.
    ''' </param>
    ''' <param name="compValues">
    '''   array of floating point numbers, corresponding to the offset correction
    '''   to apply for each of the temperature variation included in the first
    '''   argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_offsetChgCompensationTable(tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Return Me.setCompensationTable(1, tempValues, compValues)
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the weight offset thermal compensation table previously configured using the
    '''   <c>set_offsetChgCompensationTable</c> function.
    ''' <para>
    '''   The weight correction is applied by linear interpolation between specified points.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with all temperature variations for which an offset correction is specified.
    ''' </param>
    ''' <param name="compValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with the offset correction applied for each of the temperature
    '''   variation included in the first argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadOffsetChgCompensationTable(tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Return Me.loadCompensationTable(1, tempValues, compValues)
    End Function

    '''*
    ''' <summary>
    '''   Records a weight span thermal compensation table, in order to automatically correct the
    '''   measured weight based on the compensation temperature.
    ''' <para>
    '''   The weight correction will be applied by linear interpolation between specified points.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, corresponding to all averaged
    '''   temperatures for which a span correction is specified.
    ''' </param>
    ''' <param name="compValues">
    '''   array of floating point numbers, corresponding to the span correction
    '''   (in percents) to apply for each of the temperature included in the first
    '''   argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_spanAvgCompensationTable(tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Return Me.setCompensationTable(2, tempValues, compValues)
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the weight span thermal compensation table previously configured using the
    '''   <c>set_spanAvgCompensationTable</c> function.
    ''' <para>
    '''   The weight correction is applied by linear interpolation between specified points.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with all averaged temperatures for which an span correction is specified.
    ''' </param>
    ''' <param name="compValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with the span correction applied for each of the temperature
    '''   included in the first argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadSpanAvgCompensationTable(tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Return Me.loadCompensationTable(2, tempValues, compValues)
    End Function

    '''*
    ''' <summary>
    '''   Records a weight span thermal compensation table, in order to automatically correct the
    '''   measured weight based on the variation of temperature.
    ''' <para>
    '''   The weight correction will be applied by linear interpolation between specified points.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, corresponding to all variations of
    '''   temperatures for which a span correction is specified.
    ''' </param>
    ''' <param name="compValues">
    '''   array of floating point numbers, corresponding to the span correction
    '''   (in percents) to apply for each of the temperature variation included
    '''   in the first argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_spanChgCompensationTable(tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Return Me.setCompensationTable(3, tempValues, compValues)
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the weight span thermal compensation table previously configured using the
    '''   <c>set_spanChgCompensationTable</c> function.
    ''' <para>
    '''   The weight correction is applied by linear interpolation between specified points.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tempValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with all variation of temperature for which an span correction is specified.
    ''' </param>
    ''' <param name="compValues">
    '''   array of floating point numbers, that is filled by the function
    '''   with the span correction applied for each of variation of temperature
    '''   included in the first argument, index by index.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadSpanChgCompensationTable(tempValues As List(Of Double), compValues As List(Of Double)) As Integer
      Return Me.loadCompensationTable(3, tempValues, compValues)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of weighing scale sensors started using <c>yFirstWeighScale()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWeighScale</c> object, corresponding to
    '''   a weighing scale sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more weighing scale sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextWeighScale() As YWeighScale
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YWeighScale.FindWeighScale(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of weighing scale sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YWeighScale.nextWeighScale()</c> to iterate on
    '''   next weighing scale sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWeighScale</c> object, corresponding to
    '''   the first weighing scale sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstWeighScale() As YWeighScale
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("WeighScale", 0, p, size, neededsize, errmsg)
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
      Return YWeighScale.FindWeighScale(serial + "." + funcId)
    End Function

    REM --- (end of YWeighScale public methods declaration)

  End Class

  REM --- (YWeighScale functions)

  '''*
  ''' <summary>
  '''   Retrieves a weighing scale sensor for a given identifier.
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
  '''   This function does not require that the weighing scale sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YWeighScale.isOnline()</c> to test if the weighing scale sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a weighing scale sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the weighing scale sensor
  ''' </param>
  ''' <returns>
  '''   a <c>YWeighScale</c> object allowing you to drive the weighing scale sensor.
  ''' </returns>
  '''/
  Public Function yFindWeighScale(ByVal func As String) As YWeighScale
    Return YWeighScale.FindWeighScale(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of weighing scale sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YWeighScale.nextWeighScale()</c> to iterate on
  '''   next weighing scale sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YWeighScale</c> object, corresponding to
  '''   the first weighing scale sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstWeighScale() As YWeighScale
    Return YWeighScale.FirstWeighScale()
  End Function


  REM --- (end of YWeighScale functions)

End Module
