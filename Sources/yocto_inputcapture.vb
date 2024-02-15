' ********************************************************************
'
'  $Id: yocto_i2cport.vb 52943 2023-01-26 15:46:47Z mvuilleu $
'
'  Implements yFindI2cPort(), the high-level API for I2cPort functions
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
Imports YRETCODE = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_inputcapture

  REM --- (generated code: YInputCapture return codes)
    REM --- (end of generated code: YInputCapture return codes)
  REM --- (generated code: YInputCapture dlldef)
    REM --- (end of generated code: YInputCapture dlldef)
  REM --- (generated code: YInputCapture yapiwrapper)
   REM --- (end of generated code: YInputCapture yapiwrapper)
  REM --- (generated code: YInputCapture globals)

  Public Const Y_LASTCAPTURETIME_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_NSAMPLES_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_SAMPLINGRATE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CAPTURETYPE_NONE As Integer = 0
  Public Const Y_CAPTURETYPE_TIMED As Integer = 1
  Public Const Y_CAPTURETYPE_V_MAX As Integer = 2
  Public Const Y_CAPTURETYPE_V_MIN As Integer = 3
  Public Const Y_CAPTURETYPE_I_MAX As Integer = 4
  Public Const Y_CAPTURETYPE_I_MIN As Integer = 5
  Public Const Y_CAPTURETYPE_P_MAX As Integer = 6
  Public Const Y_CAPTURETYPE_P_MIN As Integer = 7
  Public Const Y_CAPTURETYPE_V_AVG_MAX As Integer = 8
  Public Const Y_CAPTURETYPE_V_AVG_MIN As Integer = 9
  Public Const Y_CAPTURETYPE_V_RMS_MAX As Integer = 10
  Public Const Y_CAPTURETYPE_V_RMS_MIN As Integer = 11
  Public Const Y_CAPTURETYPE_I_AVG_MAX As Integer = 12
  Public Const Y_CAPTURETYPE_I_AVG_MIN As Integer = 13
  Public Const Y_CAPTURETYPE_I_RMS_MAX As Integer = 14
  Public Const Y_CAPTURETYPE_I_RMS_MIN As Integer = 15
  Public Const Y_CAPTURETYPE_P_AVG_MAX As Integer = 16
  Public Const Y_CAPTURETYPE_P_AVG_MIN As Integer = 17
  Public Const Y_CAPTURETYPE_PF_MIN As Integer = 18
  Public Const Y_CAPTURETYPE_DPF_MIN As Integer = 19
  Public Const Y_CAPTURETYPE_INVALID As Integer = -1
  Public Const Y_CONDVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_CONDALIGN_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CAPTURETYPEATSTARTUP_NONE As Integer = 0
  Public Const Y_CAPTURETYPEATSTARTUP_TIMED As Integer = 1
  Public Const Y_CAPTURETYPEATSTARTUP_V_MAX As Integer = 2
  Public Const Y_CAPTURETYPEATSTARTUP_V_MIN As Integer = 3
  Public Const Y_CAPTURETYPEATSTARTUP_I_MAX As Integer = 4
  Public Const Y_CAPTURETYPEATSTARTUP_I_MIN As Integer = 5
  Public Const Y_CAPTURETYPEATSTARTUP_P_MAX As Integer = 6
  Public Const Y_CAPTURETYPEATSTARTUP_P_MIN As Integer = 7
  Public Const Y_CAPTURETYPEATSTARTUP_V_AVG_MAX As Integer = 8
  Public Const Y_CAPTURETYPEATSTARTUP_V_AVG_MIN As Integer = 9
  Public Const Y_CAPTURETYPEATSTARTUP_V_RMS_MAX As Integer = 10
  Public Const Y_CAPTURETYPEATSTARTUP_V_RMS_MIN As Integer = 11
  Public Const Y_CAPTURETYPEATSTARTUP_I_AVG_MAX As Integer = 12
  Public Const Y_CAPTURETYPEATSTARTUP_I_AVG_MIN As Integer = 13
  Public Const Y_CAPTURETYPEATSTARTUP_I_RMS_MAX As Integer = 14
  Public Const Y_CAPTURETYPEATSTARTUP_I_RMS_MIN As Integer = 15
  Public Const Y_CAPTURETYPEATSTARTUP_P_AVG_MAX As Integer = 16
  Public Const Y_CAPTURETYPEATSTARTUP_P_AVG_MIN As Integer = 17
  Public Const Y_CAPTURETYPEATSTARTUP_PF_MIN As Integer = 18
  Public Const Y_CAPTURETYPEATSTARTUP_DPF_MIN As Integer = 19
  Public Const Y_CAPTURETYPEATSTARTUP_INVALID As Integer = -1
  Public Const Y_CONDVALUEATSTARTUP_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YInputCaptureValueCallback(ByVal func As YInputCapture, ByVal value As String)
  Public Delegate Sub YInputCaptureTimedReportCallback(ByVal func As YInputCapture, ByVal measure As YMeasure)
  REM --- (end of generated code: YInputCapture globals)

  REM --- (generated code: YInputCaptureData class start)

  '''*
  ''' <c>InputCaptureData</c> objects represent raw data
  ''' sampled by the analog/digital converter present in
  ''' a Yoctopuce electrical sensor. When several inputs
  ''' are samples simultaneously, their data are provided
  ''' as distinct series.
  '''/
  Public Class YInputCaptureData
    REM --- (end of generated code: YInputCaptureData class start)
    REM --- (generated code: YInputCaptureData definitions)
    REM --- (end of generated code: YInputCaptureData definitions)
    REM --- (generated code: YInputCaptureData attributes declaration)
    Protected _fmt As Integer
    Protected _var1size As Integer
    Protected _var2size As Integer
    Protected _var3size As Integer
    Protected _nVars As Integer
    Protected _recOfs As Integer
    Protected _nRecs As Integer
    Protected _samplesPerSec As Integer
    Protected _trigType As Integer
    Protected _trigVal As Double
    Protected _trigPos As Integer
    Protected _trigUTC As Double
    Protected _var1unit As String
    Protected _var2unit As String
    Protected _var3unit As String
    Protected _var1samples As List(Of Double)
    Protected _var2samples As List(Of Double)
    Protected _var3samples As List(Of Double)
    REM --- (end of generated code: YInputCaptureData attributes declaration)

    REM --- (generated code: YInputCaptureData private methods declaration)

    REM --- (end of generated code: YInputCaptureData private methods declaration)

    REM --- (generated code: YInputCaptureData public methods declaration)
    Public Overridable Function _decodeU16(sdata As Byte(), ofs As Integer) As Integer
      Dim v As Integer = 0
      v = sdata(ofs)
      v = v + 256 * sdata(ofs+1)
      Return v
    End Function

    Public Overridable Function _decodeU32(sdata As Byte(), ofs As Integer) As Double
      Dim v As Double = 0
      v = Me._decodeU16(sdata, ofs)
      v = v + 65536.0 * Me._decodeU16(sdata, ofs+2)
      Return v
    End Function

    Public Overridable Function _decodeVal(sdata As Byte(), ofs As Integer, len As Integer) As Double
      Dim v As Double = 0
      Dim b As Double = 0
      v = Me._decodeU16(sdata, ofs)
      b = 65536.0
      ofs = ofs + 2
      len = len - 2
      While (len > 0)
        v = v + b * sdata(ofs)
        b = b * 256
        ofs = ofs + 1
        len = len - 1
      End While
      If (v > (b/2)) Then
        REM // negative number
        v = v - b
      End If
      Return v
    End Function

    Public Overridable Function _decodeSnapBin(sdata As Byte()) As Integer
      Dim buffSize As Integer = 0
      Dim recOfs As Integer = 0
      Dim ms As Integer = 0
      Dim recSize As Integer = 0
      Dim count As Integer = 0
      Dim mult1 As Integer = 0
      Dim mult2 As Integer = 0
      Dim mult3 As Integer = 0
      Dim v As Double = 0

      buffSize = (sdata).Length
      If Not(buffSize >= 24) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "Invalid snapshot data (too short)")
        return YAPI.INVALID_ARGUMENT
      end if
      Me._fmt = sdata(0)
      Me._var1size = sdata(1) - 48
      Me._var2size = sdata(2) - 48
      Me._var3size = sdata(3) - 48
      If Not(Me._fmt = 83) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "Unsupported snapshot format")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not((Me._var1size >= 2) AndAlso (Me._var1size <= 4)) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "Invalid sample size")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not((Me._var2size >= 0) AndAlso (Me._var1size <= 4)) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "Invalid sample size")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not((Me._var3size >= 0) AndAlso (Me._var1size <= 4)) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "Invalid sample size")
        return YAPI.INVALID_ARGUMENT
      end if
      If (Me._var2size = 0) Then
        Me._nVars = 1
      Else
        If (Me._var3size = 0) Then
          Me._nVars = 2
        Else
          Me._nVars = 3
        End If
      End If
      recSize = Me._var1size + Me._var2size + Me._var3size
      Me._recOfs = Me._decodeU16(sdata, 4)
      Me._nRecs = Me._decodeU16(sdata, 6)
      Me._samplesPerSec = Me._decodeU16(sdata, 8)
      Me._trigType = Me._decodeU16(sdata, 10)
      Me._trigVal = Me._decodeVal(sdata, 12, 4) / 1000
      Me._trigPos = Me._decodeU16(sdata, 16)
      ms = Me._decodeU16(sdata, 18)
      Me._trigUTC = Me._decodeVal(sdata, 20, 4)
      Me._trigUTC = Me._trigUTC + (ms / 1000.0)
      recOfs = 24
      While (sdata(recOfs) >= 32)
        Me._var1unit = "" +  Me._var1unit + "" + Chr(sdata(recOfs))
        recOfs = recOfs + 1
      End While
      If (Me._var2size > 0) Then
        recOfs = recOfs + 1
        While (sdata(recOfs) >= 32)
          Me._var2unit = "" +  Me._var2unit + "" + Chr(sdata(recOfs))
          recOfs = recOfs + 1
        End While
      End If
      If (Me._var3size > 0) Then
        recOfs = recOfs + 1
        While (sdata(recOfs) >= 32)
          Me._var3unit = "" +  Me._var3unit + "" + Chr(sdata(recOfs))
          recOfs = recOfs + 1
        End While
      End If
      If (((recOfs) And (1)) = 1) Then
        REM // align to next word
        recOfs = recOfs + 1
      End If
      mult1 = 1
      mult2 = 1
      mult3 = 1
      If (recOfs < Me._recOfs) Then
        REM // load optional value multiplier
        mult1 = Me._decodeU16(sdata, Me._recOfs)
        recOfs = recOfs + 2
        If (Me._var2size > 0) Then
          mult2 = Me._decodeU16(sdata, Me._recOfs)
          recOfs = recOfs + 2
        End If
        If (Me._var3size > 0) Then
          mult3 = Me._decodeU16(sdata, Me._recOfs)
          recOfs = recOfs + 2
        End If
      End If

      recOfs = Me._recOfs
      count = Me._nRecs
      While ((count > 0) AndAlso (recOfs + Me._var1size <= buffSize))
        v = Me._decodeVal(sdata, recOfs, Me._var1size) / 1000.0
        Me._var1samples.Add(v*mult1)
        recOfs = recOfs + recSize
      End While

      If (Me._var2size > 0) Then
        recOfs = Me._recOfs + Me._var1size
        count = Me._nRecs
        While ((count > 0) AndAlso (recOfs + Me._var2size <= buffSize))
          v = Me._decodeVal(sdata, recOfs, Me._var2size) / 1000.0
          Me._var2samples.Add(v*mult2)
          recOfs = recOfs + recSize
        End While
      End If
      If (Me._var3size > 0) Then
        recOfs = Me._recOfs + Me._var1size + Me._var2size
        count = Me._nRecs
        While ((count > 0) AndAlso (recOfs + Me._var3size <= buffSize))
          v = Me._decodeVal(sdata, recOfs, Me._var3size) / 1000.0
          Me._var3samples.Add(v*mult3)
          recOfs = recOfs + recSize
        End While
      End If
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of series available in the capture.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of
    '''   simultaneous data series available.
    ''' </returns>
    '''/
    Public Overridable Function get_serieCount() As Integer
      Return Me._nVars
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of records captured (in a serie).
    ''' <para>
    '''   In the exceptional case where it was not possible
    '''   to transfer all data in time, the number of records
    '''   actually present in the series might be lower than
    '''   the number of records captured
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of
    '''   records expected in each serie.
    ''' </returns>
    '''/
    Public Overridable Function get_recordCount() As Integer
      Return Me._nRecs
    End Function

    '''*
    ''' <summary>
    '''   Returns the effective sampling rate of the device.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of
    '''   samples taken each second.
    ''' </returns>
    '''/
    Public Overridable Function get_samplingRate() As Integer
      Return Me._samplesPerSec
    End Function

    '''*
    ''' <summary>
    '''   Returns the type of automatic conditional capture
    '''   that triggered the capture of this data sequence.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the type of conditional capture.
    ''' </returns>
    '''/
    Public Overridable Function get_captureType() As Integer
      Return CType(Me._trigType, Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns the threshold value that triggered
    '''   this automatic conditional capture, if it was
    '''   not an instant captured triggered manually.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the conditional threshold value
    '''   at the time of capture.
    ''' </returns>
    '''/
    Public Overridable Function get_triggerValue() As Double
      Return Me._trigVal
    End Function

    '''*
    ''' <summary>
    '''   Returns the index in the series of the sample
    '''   corresponding to the exact time when the capture
    '''   was triggered.
    ''' <para>
    '''   In case of trigger based on average
    '''   or RMS value, the trigger index corresponds to
    '''   the end of the averaging period.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to a position
    '''   in the data serie.
    ''' </returns>
    '''/
    Public Overridable Function get_triggerPosition() As Integer
      Return Me._trigPos
    End Function

    '''*
    ''' <summary>
    '''   Returns the absolute time when the capture was
    '''   triggered, as a Unix timestamp.
    ''' <para>
    '''   Milliseconds are
    '''   included in this timestamp (floating-point number).
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to
    '''   the number of seconds between the Jan 1,
    '''   1970 and the moment where the capture
    '''   was triggered.
    ''' </returns>
    '''/
    Public Overridable Function get_triggerRealTimeUTC() As Double
      Return Me._trigUTC
    End Function

    '''*
    ''' <summary>
    '''   Returns the unit of measurement for data points in
    '''   the first serie.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string containing to a physical unit of
    '''   measurement.
    ''' </returns>
    '''/
    Public Overridable Function get_serie1Unit() As String
      Return Me._var1unit
    End Function

    '''*
    ''' <summary>
    '''   Returns the unit of measurement for data points in
    '''   the second serie.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string containing to a physical unit of
    '''   measurement.
    ''' </returns>
    '''/
    Public Overridable Function get_serie2Unit() As String
      If Not(Me._nVars >= 2) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "There is no serie 2 in this capture data")
        return ""
      end if
      Return Me._var2unit
    End Function

    '''*
    ''' <summary>
    '''   Returns the unit of measurement for data points in
    '''   the third serie.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string containing to a physical unit of
    '''   measurement.
    ''' </returns>
    '''/
    Public Overridable Function get_serie3Unit() As String
      If Not(Me._nVars >= 3) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "There is no serie 3 in this capture data")
        return ""
      end if
      Return Me._var3unit
    End Function

    '''*
    ''' <summary>
    '''   Returns the sampled data corresponding to the first serie.
    ''' <para>
    '''   The corresponding physical unit can be obtained
    '''   using the method <c>get_serie1Unit()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of real numbers corresponding to all
    '''   samples received for serie 1.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_serie1Values() As List(Of Double)
      Return Me._var1samples
    End Function

    '''*
    ''' <summary>
    '''   Returns the sampled data corresponding to the second serie.
    ''' <para>
    '''   The corresponding physical unit can be obtained
    '''   using the method <c>get_serie2Unit()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of real numbers corresponding to all
    '''   samples received for serie 2.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_serie2Values() As List(Of Double)
      If Not(Me._nVars >= 2) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "There is no serie 2 in this capture data")
        return Me._var2samples
      end if
      Return Me._var2samples
    End Function

    '''*
    ''' <summary>
    '''   Returns the sampled data corresponding to the third serie.
    ''' <para>
    '''   The corresponding physical unit can be obtained
    '''   using the method <c>get_serie3Unit()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of real numbers corresponding to all
    '''   samples received for serie 3.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_serie3Values() As List(Of Double)
      If Not(Me._nVars >= 3) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "There is no serie 3 in this capture data")
        return Me._var3samples
      end if
      Return Me._var3samples
    End Function



    REM --- (end of generated code: YInputCaptureData public methods declaration)

    Public Sub _throw(ByVal errType As YRETCODE, ByVal errMsg As String)
      If Not (YAPI.ExceptionsDisabled) Then
        Throw New YAPI_Exception(errType, "YoctoApi error : " + errMsg)
      End If
    End Sub

    Public Sub New(yfun As YFunction, sdata As Byte())
      Me._decodeSnapBin(sdata)
    End Sub

  End Class


  REM --- (generated code: YInputCapture class start)

  '''*
  ''' <summary>
  '''   The <c>YInputCapture</c> class allows you to access data samples
  '''   measured by a Yoctopuce electrical sensor.
  ''' <para>
  '''   The data capture can be
  '''   triggered manually, or be configured to detect specific events.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YInputCapture
    Inherits YFunction
    REM --- (end of generated code: YInputCapture class start)

    REM --- (generated code: YInputCapture definitions)
    Public Const LASTCAPTURETIME_INVALID As Long = YAPI.INVALID_LONG
    Public Const NSAMPLES_INVALID As Integer = YAPI.INVALID_UINT
    Public Const SAMPLINGRATE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CAPTURETYPE_NONE As Integer = 0
    Public Const CAPTURETYPE_TIMED As Integer = 1
    Public Const CAPTURETYPE_V_MAX As Integer = 2
    Public Const CAPTURETYPE_V_MIN As Integer = 3
    Public Const CAPTURETYPE_I_MAX As Integer = 4
    Public Const CAPTURETYPE_I_MIN As Integer = 5
    Public Const CAPTURETYPE_P_MAX As Integer = 6
    Public Const CAPTURETYPE_P_MIN As Integer = 7
    Public Const CAPTURETYPE_V_AVG_MAX As Integer = 8
    Public Const CAPTURETYPE_V_AVG_MIN As Integer = 9
    Public Const CAPTURETYPE_V_RMS_MAX As Integer = 10
    Public Const CAPTURETYPE_V_RMS_MIN As Integer = 11
    Public Const CAPTURETYPE_I_AVG_MAX As Integer = 12
    Public Const CAPTURETYPE_I_AVG_MIN As Integer = 13
    Public Const CAPTURETYPE_I_RMS_MAX As Integer = 14
    Public Const CAPTURETYPE_I_RMS_MIN As Integer = 15
    Public Const CAPTURETYPE_P_AVG_MAX As Integer = 16
    Public Const CAPTURETYPE_P_AVG_MIN As Integer = 17
    Public Const CAPTURETYPE_PF_MIN As Integer = 18
    Public Const CAPTURETYPE_DPF_MIN As Integer = 19
    Public Const CAPTURETYPE_INVALID As Integer = -1
    Public Const CONDVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const CONDALIGN_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CAPTURETYPEATSTARTUP_NONE As Integer = 0
    Public Const CAPTURETYPEATSTARTUP_TIMED As Integer = 1
    Public Const CAPTURETYPEATSTARTUP_V_MAX As Integer = 2
    Public Const CAPTURETYPEATSTARTUP_V_MIN As Integer = 3
    Public Const CAPTURETYPEATSTARTUP_I_MAX As Integer = 4
    Public Const CAPTURETYPEATSTARTUP_I_MIN As Integer = 5
    Public Const CAPTURETYPEATSTARTUP_P_MAX As Integer = 6
    Public Const CAPTURETYPEATSTARTUP_P_MIN As Integer = 7
    Public Const CAPTURETYPEATSTARTUP_V_AVG_MAX As Integer = 8
    Public Const CAPTURETYPEATSTARTUP_V_AVG_MIN As Integer = 9
    Public Const CAPTURETYPEATSTARTUP_V_RMS_MAX As Integer = 10
    Public Const CAPTURETYPEATSTARTUP_V_RMS_MIN As Integer = 11
    Public Const CAPTURETYPEATSTARTUP_I_AVG_MAX As Integer = 12
    Public Const CAPTURETYPEATSTARTUP_I_AVG_MIN As Integer = 13
    Public Const CAPTURETYPEATSTARTUP_I_RMS_MAX As Integer = 14
    Public Const CAPTURETYPEATSTARTUP_I_RMS_MIN As Integer = 15
    Public Const CAPTURETYPEATSTARTUP_P_AVG_MAX As Integer = 16
    Public Const CAPTURETYPEATSTARTUP_P_AVG_MIN As Integer = 17
    Public Const CAPTURETYPEATSTARTUP_PF_MIN As Integer = 18
    Public Const CAPTURETYPEATSTARTUP_DPF_MIN As Integer = 19
    Public Const CAPTURETYPEATSTARTUP_INVALID As Integer = -1
    Public Const CONDVALUEATSTARTUP_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of generated code: YInputCapture definitions)

    REM --- (generated code: YInputCapture attributes declaration)
    Protected _lastCaptureTime As Long
    Protected _nSamples As Integer
    Protected _samplingRate As Integer
    Protected _captureType As Integer
    Protected _condValue As Double
    Protected _condAlign As Integer
    Protected _captureTypeAtStartup As Integer
    Protected _condValueAtStartup As Double
    Protected _valueCallbackInputCapture As YInputCaptureValueCallback
    REM --- (end of generated code: YInputCapture attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _className = "I2cPort"
      REM --- (generated code: YInputCapture attributes initialization)
      _lastCaptureTime = LASTCAPTURETIME_INVALID
      _nSamples = NSAMPLES_INVALID
      _samplingRate = SAMPLINGRATE_INVALID
      _captureType = CAPTURETYPE_INVALID
      _condValue = CONDVALUE_INVALID
      _condAlign = CONDALIGN_INVALID
      _captureTypeAtStartup = CAPTURETYPEATSTARTUP_INVALID
      _condValueAtStartup = CONDVALUEATSTARTUP_INVALID
      _valueCallbackInputCapture = Nothing
      REM --- (end of generated code: YInputCapture attributes initialization)
    End Sub

    REM --- (generated code: YInputCapture private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("lastCaptureTime") Then
        _lastCaptureTime = json_val.getLong("lastCaptureTime")
      End If
      If json_val.has("nSamples") Then
        _nSamples = CInt(json_val.getLong("nSamples"))
      End If
      If json_val.has("samplingRate") Then
        _samplingRate = CInt(json_val.getLong("samplingRate"))
      End If
      If json_val.has("captureType") Then
        _captureType = CInt(json_val.getLong("captureType"))
      End If
      If json_val.has("condValue") Then
        _condValue = Math.Round(json_val.getDouble("condValue") / 65.536) / 1000.0
      End If
      If json_val.has("condAlign") Then
        _condAlign = CInt(json_val.getLong("condAlign"))
      End If
      If json_val.has("captureTypeAtStartup") Then
        _captureTypeAtStartup = CInt(json_val.getLong("captureTypeAtStartup"))
      End If
      If json_val.has("condValueAtStartup") Then
        _condValueAtStartup = Math.Round(json_val.getDouble("condValueAtStartup") / 65.536) / 1000.0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YInputCapture private methods declaration)

    REM --- (generated code: YInputCapture public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the number of elapsed milliseconds between the module power on
    '''   and the last capture (time of trigger), or zero if no capture has been done.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of elapsed milliseconds between the module power on
    '''   and the last capture (time of trigger), or zero if no capture has been done
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputCapture.LASTCAPTURETIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastCaptureTime() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LASTCAPTURETIME_INVALID
        End If
      End If
      res = Me._lastCaptureTime
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of samples that will be captured.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of samples that will be captured
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputCapture.NSAMPLES_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nSamples() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NSAMPLES_INVALID
        End If
      End If
      res = Me._nSamples
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the type of automatic conditional capture.
    ''' <para>
    '''   The maximum number of samples depends on the device memory.
    ''' </para>
    ''' <para>
    '''   If you want the change to be kept after a device reboot,
    '''   make sure  to call the matching module <c>saveToFlash()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the type of automatic conditional capture
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
    Public Function set_nSamples(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("nSamples", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the sampling frequency, in Hz.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the sampling frequency, in Hz
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputCapture.SAMPLINGRATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_samplingRate() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SAMPLINGRATE_INVALID
        End If
      End If
      res = Me._samplingRate
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the type of automatic conditional capture.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YInputCapture.CAPTURETYPE_NONE</c>, <c>YInputCapture.CAPTURETYPE_TIMED</c>,
    '''   <c>YInputCapture.CAPTURETYPE_V_MAX</c>, <c>YInputCapture.CAPTURETYPE_V_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_I_MAX</c>, <c>YInputCapture.CAPTURETYPE_I_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_P_MAX</c>, <c>YInputCapture.CAPTURETYPE_P_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_V_AVG_MAX</c>, <c>YInputCapture.CAPTURETYPE_V_AVG_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_V_RMS_MAX</c>, <c>YInputCapture.CAPTURETYPE_V_RMS_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_I_AVG_MAX</c>, <c>YInputCapture.CAPTURETYPE_I_AVG_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_I_RMS_MAX</c>, <c>YInputCapture.CAPTURETYPE_I_RMS_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_P_AVG_MAX</c>, <c>YInputCapture.CAPTURETYPE_P_AVG_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_PF_MIN</c> and <c>YInputCapture.CAPTURETYPE_DPF_MIN</c> corresponding
    '''   to the type of automatic conditional capture
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputCapture.CAPTURETYPE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_captureType() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CAPTURETYPE_INVALID
        End If
      End If
      res = Me._captureType
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the type of automatic conditional capture.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YInputCapture.CAPTURETYPE_NONE</c>, <c>YInputCapture.CAPTURETYPE_TIMED</c>,
    '''   <c>YInputCapture.CAPTURETYPE_V_MAX</c>, <c>YInputCapture.CAPTURETYPE_V_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_I_MAX</c>, <c>YInputCapture.CAPTURETYPE_I_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_P_MAX</c>, <c>YInputCapture.CAPTURETYPE_P_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_V_AVG_MAX</c>, <c>YInputCapture.CAPTURETYPE_V_AVG_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_V_RMS_MAX</c>, <c>YInputCapture.CAPTURETYPE_V_RMS_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_I_AVG_MAX</c>, <c>YInputCapture.CAPTURETYPE_I_AVG_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_I_RMS_MAX</c>, <c>YInputCapture.CAPTURETYPE_I_RMS_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_P_AVG_MAX</c>, <c>YInputCapture.CAPTURETYPE_P_AVG_MIN</c>,
    '''   <c>YInputCapture.CAPTURETYPE_PF_MIN</c> and <c>YInputCapture.CAPTURETYPE_DPF_MIN</c> corresponding
    '''   to the type of automatic conditional capture
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
    Public Function set_captureType(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("captureType", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes current threshold value for automatic conditional capture.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to current threshold value for automatic conditional capture
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
    Public Function set_condValue(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("condValue", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns current threshold value for automatic conditional capture.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to current threshold value for automatic conditional capture
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputCapture.CONDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_condValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CONDVALUE_INVALID
        End If
      End If
      res = Me._condValue
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the relative position of the trigger event within the capture window.
    ''' <para>
    '''   When the value is 50%, the capture is centered on the event.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the relative position of the trigger event within the capture window
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputCapture.CONDALIGN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_condAlign() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CONDALIGN_INVALID
        End If
      End If
      res = Me._condAlign
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the relative position of the trigger event within the capture window.
    ''' <para>
    '''   The new value must be between 10% (on the left) and 90% (on the right).
    '''   When the value is 50%, the capture is centered on the event.
    ''' </para>
    ''' <para>
    '''   If you want the change to be kept after a device reboot,
    '''   make sure  to call the matching module <c>saveToFlash()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the relative position of the trigger event within the capture window
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
    Public Function set_condAlign(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("condAlign", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the type of automatic conditional capture
    '''   applied at device power on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YInputCapture.CAPTURETYPEATSTARTUP_NONE</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_TIMED</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_V_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_V_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_I_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_I_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_P_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_P_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_V_AVG_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_V_AVG_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_V_RMS_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_V_RMS_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_I_AVG_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_I_AVG_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_I_RMS_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_I_RMS_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_P_AVG_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_P_AVG_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_PF_MIN</c>
    '''   and <c>YInputCapture.CAPTURETYPEATSTARTUP_DPF_MIN</c> corresponding to the type of automatic conditional capture
    '''   applied at device power on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputCapture.CAPTURETYPEATSTARTUP_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_captureTypeAtStartup() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CAPTURETYPEATSTARTUP_INVALID
        End If
      End If
      res = Me._captureTypeAtStartup
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the type of automatic conditional capture
    '''   applied at device power on.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   If you want the change to be kept after a device reboot,
    '''   make sure  to call the matching module <c>saveToFlash()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YInputCapture.CAPTURETYPEATSTARTUP_NONE</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_TIMED</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_V_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_V_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_I_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_I_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_P_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_P_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_V_AVG_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_V_AVG_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_V_RMS_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_V_RMS_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_I_AVG_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_I_AVG_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_I_RMS_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_I_RMS_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_P_AVG_MAX</c>,
    '''   <c>YInputCapture.CAPTURETYPEATSTARTUP_P_AVG_MIN</c>, <c>YInputCapture.CAPTURETYPEATSTARTUP_PF_MIN</c>
    '''   and <c>YInputCapture.CAPTURETYPEATSTARTUP_DPF_MIN</c> corresponding to the type of automatic conditional capture
    '''   applied at device power on
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
    Public Function set_captureTypeAtStartup(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("captureTypeAtStartup", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes current threshold value for automatic conditional
    '''   capture applied at device power on.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   If you want the change to be kept after a device reboot,
    '''   make sure  to call the matching module <c>saveToFlash()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to current threshold value for automatic conditional
    '''   capture applied at device power on
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
    Public Function set_condValueAtStartup(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("condValueAtStartup", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the threshold value for automatic conditional
    '''   capture applied at device power on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the threshold value for automatic conditional
    '''   capture applied at device power on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputCapture.CONDVALUEATSTARTUP_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_condValueAtStartup() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CONDVALUEATSTARTUP_INVALID
        End If
      End If
      res = Me._condValueAtStartup
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves an instant snapshot trigger for a given identifier.
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
    '''   This function does not require that the instant snapshot trigger is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YInputCapture.isOnline()</c> to test if the instant snapshot trigger is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an instant snapshot trigger by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the instant snapshot trigger, for instance
    '''   <c>MyDevice.inputCapture</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YInputCapture</c> object allowing you to drive the instant snapshot trigger.
    ''' </returns>
    '''/
    Public Shared Function FindInputCapture(func As String) As YInputCapture
      Dim obj As YInputCapture
      obj = CType(YFunction._FindFromCache("InputCapture", func), YInputCapture)
      If ((obj Is Nothing)) Then
        obj = New YInputCapture(func)
        YFunction._AddToCache("InputCapture", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YInputCaptureValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackInputCapture = callback
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
      If (Not (Me._valueCallbackInputCapture Is Nothing)) Then
        Me._valueCallbackInputCapture(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns all details about the last automatic input capture.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an <c>YInputCaptureData</c> object including
    '''   data series and all related meta-information.
    '''   On failure, throws an exception or returns an capture object.
    ''' </returns>
    '''/
    Public Overridable Function get_lastCapture() As YInputCaptureData
      Dim snapData As Byte() = New Byte(){}

      snapData = Me._download("snap.bin")
      Return New YInputCaptureData(Me, snapData)
    End Function

    '''*
    ''' <summary>
    '''   Returns a new immediate capture of the device inputs.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="msDuration">
    '''   duration of the capture window,
    '''   in milliseconds (eg. between 20 and 1000).
    ''' </param>
    ''' <returns>
    '''   an <c>YInputCaptureData</c> object including
    '''   data series for the specified duration.
    '''   On failure, throws an exception or returns an capture object.
    ''' </returns>
    '''/
    Public Overridable Function get_immediateCapture(msDuration As Integer) As YInputCaptureData
      Dim snapUrl As String
      Dim snapData As Byte() = New Byte(){}
      Dim snapStart As Integer = 0
      If (msDuration < 1) Then
        msDuration = 20
      End If
      If (msDuration > 1000) Then
        msDuration = 1000
      End If
      snapStart = (-msDuration \ 2)
      snapUrl = "snap.bin?t=" + Convert.ToString( snapStart) + "&d=" + Convert.ToString(msDuration)

      snapData = Me._download(snapUrl)
      Return New YInputCaptureData(Me, snapData)
    End Function


    '''*
    ''' <summary>
    '''   c
    ''' <para>
    '''   omment from .yc definition
    ''' </para>
    ''' </summary>
    '''/
    Public Function nextInputCapture() As YInputCapture
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YInputCapture.FindInputCapture(hwid)
    End Function

    '''*
    ''' <summary>
    '''   c
    ''' <para>
    '''   omment from .yc definition
    ''' </para>
    ''' </summary>
    '''/
    Public Shared Function FirstInputCapture() As YInputCapture
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("InputCapture", 0, p, size, neededsize, errmsg)
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
      Return YInputCapture.FindInputCapture(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YInputCapture public methods declaration)

  End Class

  REM --- (generated code: YInputCapture functions)

  '''*
  ''' <summary>
  '''   Retrieves an instant snapshot trigger for a given identifier.
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
  '''   This function does not require that the instant snapshot trigger is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YInputCapture.isOnline()</c> to test if the instant snapshot trigger is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an instant snapshot trigger by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the instant snapshot trigger, for instance
  '''   <c>MyDevice.inputCapture</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YInputCapture</c> object allowing you to drive the instant snapshot trigger.
  ''' </returns>
  '''/
  Public Function yFindInputCapture(ByVal func As String) As YInputCapture
    Return YInputCapture.FindInputCapture(func)
  End Function

  '''*
  ''' <summary>
  '''   A
  ''' <para>
  '''   lias for Y{$classname}.First{$classname}()
  ''' </para>
  ''' </summary>
  '''/
  Public Function yFirstInputCapture() As YInputCapture
    Return YInputCapture.FirstInputCapture()
  End Function


  REM --- (end of generated code: YInputCapture functions)

End Module
