' ********************************************************************
'
'  $Id: yocto_gps.vb 37165 2019-09-13 16:57:27Z mvuilleu $
'
'  Implements yFindGps(), the high-level API for Gps functions
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

Module yocto_gps

    REM --- (YGps return codes)
    REM --- (end of YGps return codes)
    REM --- (YGps dlldef)
    REM --- (end of YGps dlldef)
   REM --- (YGps yapiwrapper)
   REM --- (end of YGps yapiwrapper)
  REM --- (YGps globals)

  Public Const Y_ISFIXED_FALSE As Integer = 0
  Public Const Y_ISFIXED_TRUE As Integer = 1
  Public Const Y_ISFIXED_INVALID As Integer = -1
  Public Const Y_SATCOUNT_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_COORDSYSTEM_GPS_DMS As Integer = 0
  Public Const Y_COORDSYSTEM_GPS_DM As Integer = 1
  Public Const Y_COORDSYSTEM_GPS_D As Integer = 2
  Public Const Y_COORDSYSTEM_INVALID As Integer = -1
  Public Const Y_CONSTELLATION_GPS As Integer = 0
  Public Const Y_CONSTELLATION_GLONASS As Integer = 1
  Public Const Y_CONSTELLATION_GALLILEO As Integer = 2
  Public Const Y_CONSTELLATION_GNSS As Integer = 3
  Public Const Y_CONSTELLATION_GPS_GLONASS As Integer = 4
  Public Const Y_CONSTELLATION_GPS_GALLILEO As Integer = 5
  Public Const Y_CONSTELLATION_GLONASS_GALLELIO As Integer = 6
  Public Const Y_CONSTELLATION_INVALID As Integer = -1
  Public Const Y_LATITUDE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_LONGITUDE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_DILUTION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_ALTITUDE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_GROUNDSPEED_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_DIRECTION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_UNIXTIME_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_DATETIME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_UTCOFFSET_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YGpsValueCallback(ByVal func As YGps, ByVal value As String)
  Public Delegate Sub YGpsTimedReportCallback(ByVal func As YGps, ByVal measure As YMeasure)
  REM --- (end of YGps globals)

  REM --- (YGps class start)

  '''*
  ''' <summary>
  '''   The GPS function allows you to extract positioning
  '''   data from the GPS device.
  ''' <para>
  '''   This class can provides
  '''   complete positioning information: However, if you
  '''   wish to define callbacks on position changes, you
  '''   should use the YLatitude et YLongitude classes.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YGps
    Inherits YFunction
    REM --- (end of YGps class start)

    REM --- (YGps definitions)
    Public Const ISFIXED_FALSE As Integer = 0
    Public Const ISFIXED_TRUE As Integer = 1
    Public Const ISFIXED_INVALID As Integer = -1
    Public Const SATCOUNT_INVALID As Long = YAPI.INVALID_LONG
    Public Const COORDSYSTEM_GPS_DMS As Integer = 0
    Public Const COORDSYSTEM_GPS_DM As Integer = 1
    Public Const COORDSYSTEM_GPS_D As Integer = 2
    Public Const COORDSYSTEM_INVALID As Integer = -1
    Public Const CONSTELLATION_GPS As Integer = 0
    Public Const CONSTELLATION_GLONASS As Integer = 1
    Public Const CONSTELLATION_GALLILEO As Integer = 2
    Public Const CONSTELLATION_GNSS As Integer = 3
    Public Const CONSTELLATION_GPS_GLONASS As Integer = 4
    Public Const CONSTELLATION_GPS_GALLILEO As Integer = 5
    Public Const CONSTELLATION_GLONASS_GALLELIO As Integer = 6
    Public Const CONSTELLATION_INVALID As Integer = -1
    Public Const LATITUDE_INVALID As String = YAPI.INVALID_STRING
    Public Const LONGITUDE_INVALID As String = YAPI.INVALID_STRING
    Public Const DILUTION_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const ALTITUDE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const GROUNDSPEED_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const DIRECTION_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const UNIXTIME_INVALID As Long = YAPI.INVALID_LONG
    Public Const DATETIME_INVALID As String = YAPI.INVALID_STRING
    Public Const UTCOFFSET_INVALID As Integer = YAPI.INVALID_INT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YGps definitions)

    REM --- (YGps attributes declaration)
    Protected _isFixed As Integer
    Protected _satCount As Long
    Protected _coordSystem As Integer
    Protected _constellation As Integer
    Protected _latitude As String
    Protected _longitude As String
    Protected _dilution As Double
    Protected _altitude As Double
    Protected _groundSpeed As Double
    Protected _direction As Double
    Protected _unixTime As Long
    Protected _dateTime As String
    Protected _utcOffset As Integer
    Protected _command As String
    Protected _valueCallbackGps As YGpsValueCallback
    REM --- (end of YGps attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Gps"
      REM --- (YGps attributes initialization)
      _isFixed = ISFIXED_INVALID
      _satCount = SATCOUNT_INVALID
      _coordSystem = COORDSYSTEM_INVALID
      _constellation = CONSTELLATION_INVALID
      _latitude = LATITUDE_INVALID
      _longitude = LONGITUDE_INVALID
      _dilution = DILUTION_INVALID
      _altitude = ALTITUDE_INVALID
      _groundSpeed = GROUNDSPEED_INVALID
      _direction = DIRECTION_INVALID
      _unixTime = UNIXTIME_INVALID
      _dateTime = DATETIME_INVALID
      _utcOffset = UTCOFFSET_INVALID
      _command = COMMAND_INVALID
      _valueCallbackGps = Nothing
      REM --- (end of YGps attributes initialization)
    End Sub

    REM --- (YGps private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("isFixed") Then
        If (json_val.getInt("isFixed") > 0) Then _isFixed = 1 Else _isFixed = 0
      End If
      If json_val.has("satCount") Then
        _satCount = json_val.getLong("satCount")
      End If
      If json_val.has("coordSystem") Then
        _coordSystem = CInt(json_val.getLong("coordSystem"))
      End If
      If json_val.has("constellation") Then
        _constellation = CInt(json_val.getLong("constellation"))
      End If
      If json_val.has("latitude") Then
        _latitude = json_val.getString("latitude")
      End If
      If json_val.has("longitude") Then
        _longitude = json_val.getString("longitude")
      End If
      If json_val.has("dilution") Then
        _dilution = Math.Round(json_val.getDouble("dilution") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("altitude") Then
        _altitude = Math.Round(json_val.getDouble("altitude") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("groundSpeed") Then
        _groundSpeed = Math.Round(json_val.getDouble("groundSpeed") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("direction") Then
        _direction = Math.Round(json_val.getDouble("direction") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("unixTime") Then
        _unixTime = json_val.getLong("unixTime")
      End If
      If json_val.has("dateTime") Then
        _dateTime = json_val.getString("dateTime")
      End If
      If json_val.has("utcOffset") Then
        _utcOffset = CInt(json_val.getLong("utcOffset"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YGps private methods declaration)

    REM --- (YGps public methods declaration)
    '''*
    ''' <summary>
    '''   Returns TRUE if the receiver has found enough satellites to work.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ISFIXED_FALSE</c> or <c>Y_ISFIXED_TRUE</c>, according to TRUE if the receiver has found
    '''   enough satellites to work
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ISFIXED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_isFixed() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ISFIXED_INVALID
        End If
      End If
      res = Me._isFixed
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the count of visible satellites.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the count of visible satellites
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SATCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_satCount() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SATCOUNT_INVALID
        End If
      End If
      res = Me._satCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the representation system used for positioning data.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_COORDSYSTEM_GPS_DMS</c>, <c>Y_COORDSYSTEM_GPS_DM</c> and
    '''   <c>Y_COORDSYSTEM_GPS_D</c> corresponding to the representation system used for positioning data
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_COORDSYSTEM_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_coordSystem() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COORDSYSTEM_INVALID
        End If
      End If
      res = Me._coordSystem
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the representation system used for positioning data.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_COORDSYSTEM_GPS_DMS</c>, <c>Y_COORDSYSTEM_GPS_DM</c> and
    '''   <c>Y_COORDSYSTEM_GPS_D</c> corresponding to the representation system used for positioning data
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
    Public Function set_coordSystem(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("coordSystem", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the the satellites constellation used to compute
    '''   positioning data.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_CONSTELLATION_GPS</c>, <c>Y_CONSTELLATION_GLONASS</c>,
    '''   <c>Y_CONSTELLATION_GALLILEO</c>, <c>Y_CONSTELLATION_GNSS</c>, <c>Y_CONSTELLATION_GPS_GLONASS</c>,
    '''   <c>Y_CONSTELLATION_GPS_GALLILEO</c> and <c>Y_CONSTELLATION_GLONASS_GALLELIO</c> corresponding to
    '''   the the satellites constellation used to compute
    '''   positioning data
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CONSTELLATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_constellation() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CONSTELLATION_INVALID
        End If
      End If
      res = Me._constellation
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the satellites constellation used to compute
    '''   positioning data.
    ''' <para>
    '''   Possible  constellations are GPS, Glonass, Galileo ,
    '''   GNSS ( = GPS + Glonass + Galileo) and the 3 possible pairs. This seeting has effect on Yocto-GPS rev A.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_CONSTELLATION_GPS</c>, <c>Y_CONSTELLATION_GLONASS</c>,
    '''   <c>Y_CONSTELLATION_GALLILEO</c>, <c>Y_CONSTELLATION_GNSS</c>, <c>Y_CONSTELLATION_GPS_GLONASS</c>,
    '''   <c>Y_CONSTELLATION_GPS_GALLILEO</c> and <c>Y_CONSTELLATION_GLONASS_GALLELIO</c> corresponding to
    '''   the satellites constellation used to compute
    '''   positioning data
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
    Public Function set_constellation(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("constellation", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current latitude.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current latitude
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LATITUDE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_latitude() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LATITUDE_INVALID
        End If
      End If
      res = Me._latitude
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current longitude.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current longitude
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LONGITUDE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_longitude() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LONGITUDE_INVALID
        End If
      End If
      res = Me._longitude
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current horizontal dilution of precision,
    '''   the smaller that number is, the better .
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current horizontal dilution of precision,
    '''   the smaller that number is, the better
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DILUTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_dilution() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DILUTION_INVALID
        End If
      End If
      res = Me._dilution
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current altitude.
    ''' <para>
    '''   Beware:  GPS technology
    '''   is very inaccurate regarding altitude.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current altitude
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ALTITUDE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_altitude() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ALTITUDE_INVALID
        End If
      End If
      res = Me._altitude
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current ground speed in Km/h.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current ground speed in Km/h
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_GROUNDSPEED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_groundSpeed() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return GROUNDSPEED_INVALID
        End If
      End If
      res = Me._groundSpeed
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current move bearing in degrees, zero
    '''   is the true (geographic) north.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current move bearing in degrees, zero
    '''   is the true (geographic) north
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DIRECTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_direction() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DIRECTION_INVALID
        End If
      End If
      res = Me._direction
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current time in Unix format (number of
    '''   seconds elapsed since Jan 1st, 1970).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current time in Unix format (number of
    '''   seconds elapsed since Jan 1st, 1970)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_UNIXTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_unixTime() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return UNIXTIME_INVALID
        End If
      End If
      res = Me._unixTime
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current time in the form "YYYY/MM/DD hh:mm:ss".
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current time in the form "YYYY/MM/DD hh:mm:ss"
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DATETIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_dateTime() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DATETIME_INVALID
        End If
      End If
      res = Me._dateTime
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of seconds between current time and UTC time (time zone).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of seconds between current time and UTC time (time zone)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_UTCOFFSET_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_utcOffset() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return UTCOFFSET_INVALID
        End If
      End If
      res = Me._utcOffset
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of seconds between current time and UTC time (time zone).
    ''' <para>
    '''   The timezone is automatically rounded to the nearest multiple of 15 minutes.
    '''   If current UTC time is known, the current time is automatically be updated according to the selected time zone.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the number of seconds between current time and UTC time (time zone)
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
    Public Function set_utcOffset(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("utcOffset", rest_val)
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
    '''   Retrieves a GPS for a given identifier.
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
    '''   This function does not require that the GPS is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YGps.isOnline()</c> to test if the GPS is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a GPS by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the GPS
    ''' </param>
    ''' <returns>
    '''   a <c>YGps</c> object allowing you to drive the GPS.
    ''' </returns>
    '''/
    Public Shared Function FindGps(func As String) As YGps
      Dim obj As YGps
      obj = CType(YFunction._FindFromCache("Gps", func), YGps)
      If ((obj Is Nothing)) Then
        obj = New YGps(func)
        YFunction._AddToCache("Gps", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YGpsValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackGps = callback
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
      If (Not (Me._valueCallbackGps Is Nothing)) Then
        Me._valueCallbackGps(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of GPS started using <c>yFirstGps()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned GPS order.
    '''   If you want to find a specific a GPS, use <c>Gps.findGps()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGps</c> object, corresponding to
    '''   a GPS currently online, or a <c>Nothing</c> pointer
    '''   if there are no more GPS to enumerate.
    ''' </returns>
    '''/
    Public Function nextGps() As YGps
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YGps.FindGps(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of GPS currently accessible.
    ''' <para>
    '''   Use the method <c>YGps.nextGps()</c> to iterate on
    '''   next GPS.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGps</c> object, corresponding to
    '''   the first GPS currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstGps() As YGps
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Gps", 0, p, size, neededsize, errmsg)
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
      Return YGps.FindGps(serial + "." + funcId)
    End Function

    REM --- (end of YGps public methods declaration)

  End Class

  REM --- (YGps functions)

  '''*
  ''' <summary>
  '''   Retrieves a GPS for a given identifier.
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
  '''   This function does not require that the GPS is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YGps.isOnline()</c> to test if the GPS is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a GPS by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the GPS
  ''' </param>
  ''' <returns>
  '''   a <c>YGps</c> object allowing you to drive the GPS.
  ''' </returns>
  '''/
  Public Function yFindGps(ByVal func As String) As YGps
    Return YGps.FindGps(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of GPS currently accessible.
  ''' <para>
  '''   Use the method <c>YGps.nextGps()</c> to iterate on
  '''   next GPS.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YGps</c> object, corresponding to
  '''   the first GPS currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstGps() As YGps
    Return YGps.FirstGps()
  End Function


  REM --- (end of YGps functions)

End Module
