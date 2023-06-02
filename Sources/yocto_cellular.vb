'*********************************************************************
'*
'* $Id: yocto_cellular.vb 54160 2023-04-21 07:33:49Z seb $
'*
'* Implements yFindCellular(), the high-level API for Cellular functions
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

Module yocto_cellular


    REM --- (generated code: YCellRecord return codes)
    REM --- (end of generated code: YCellRecord return codes)
    REM --- (generated code: YCellRecord dlldef)
    REM --- (end of generated code: YCellRecord dlldef)
  REM --- (generated code: YCellRecord globals)

  REM --- (end of generated code: YCellRecord globals)

  REM --- (generated code: YCellRecord class start)

  '''*
  ''' <c>YCellRecord</c> objects are used to describe a wireless network.
  ''' These objects are used in particular in conjunction with the
  ''' <c>YCellular</c> class.
  '''/
  Public Class YCellRecord
    REM --- (end of generated code: YCellRecord class start)

    REM --- (generated code: YCellRecord definitions)
    REM --- (end of generated code: YCellRecord definitions)

    REM --- (generated code: YCellRecord attributes declaration)
    Protected _oper As String
    Protected _mcc As Integer
    Protected _mnc As Integer
    Protected _lac As Integer
    Protected _cid As Integer
    Protected _dbm As Integer
    Protected _tad As Integer
    REM --- (end of generated code: YCellRecord attributes declaration)

    Public Sub New(ByVal mcc As Integer, ByVal mnc As Integer, ByVal lac As Integer, ByVal cellId As Integer, ByVal dbm As Integer, ByVal tad As Integer, ByVal oper As String)
      REM --- (generated code: YCellRecord attributes initialization)
      _mcc = 0
      _mnc = 0
      _lac = 0
      _cid = 0
      _dbm = 0
      _tad = 0
      REM --- (end of generated code: YCellRecord attributes initialization)
      _oper = oper
      _mcc = mcc
      _mnc = mnc
      _lac = lac
      _cid = cellId
      _dbm = dbm
      _tad = tad
    End Sub

    REM --- (generated code: YCellRecord private methods declaration)

    REM --- (end of generated code: YCellRecord private methods declaration)

    REM --- (generated code: YCellRecord public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the name of the the cell operator, as received from the network.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with the name of the the cell operator.
    ''' </returns>
    '''/
    Public Overridable Function get_cellOperator() As String
      Return Me._oper
    End Function

    '''*
    ''' <summary>
    '''   Returns the Mobile Country Code (MCC).
    ''' <para>
    '''   The MCC is a unique identifier for each country.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the Mobile Country Code (MCC).
    ''' </returns>
    '''/
    Public Overridable Function get_mobileCountryCode() As Integer
      Return Me._mcc
    End Function

    '''*
    ''' <summary>
    '''   Returns the Mobile Network Code (MNC).
    ''' <para>
    '''   The MNC is a unique identifier for each phone
    '''   operator within a country.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the Mobile Network Code (MNC).
    ''' </returns>
    '''/
    Public Overridable Function get_mobileNetworkCode() As Integer
      Return Me._mnc
    End Function

    '''*
    ''' <summary>
    '''   Returns the Location Area Code (LAC).
    ''' <para>
    '''   The LAC is a unique identifier for each
    '''   place within a country.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the Location Area Code (LAC).
    ''' </returns>
    '''/
    Public Overridable Function get_locationAreaCode() As Integer
      Return Me._lac
    End Function

    '''*
    ''' <summary>
    '''   Returns the Cell ID.
    ''' <para>
    '''   The Cell ID is a unique identifier for each
    '''   base transmission station within a LAC.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the Cell Id.
    ''' </returns>
    '''/
    Public Overridable Function get_cellId() As Integer
      Return Me._cid
    End Function

    '''*
    ''' <summary>
    '''   Returns the signal strength, measured in dBm.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the signal strength.
    ''' </returns>
    '''/
    Public Overridable Function get_signalStrength() As Integer
      Return Me._dbm
    End Function

    '''*
    ''' <summary>
    '''   Returns the Timing Advance (TA).
    ''' <para>
    '''   The TA corresponds to the time necessary
    '''   for the signal to reach the base station from the device.
    '''   Each increment corresponds about to 550m of distance.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the Timing Advance (TA).
    ''' </returns>
    '''/
    Public Overridable Function get_timingAdvance() As Integer
      Return Me._tad
    End Function



    REM --- (end of generated code: YCellRecord public methods declaration)

  End Class

  REM --- (generated code: YCellRecord functions)


  REM --- (end of generated code: YCellRecord functions)



    REM --- (generated code: YCellular return codes)
    REM --- (end of generated code: YCellular return codes)
    REM --- (generated code: YCellular dlldef)
    REM --- (end of generated code: YCellular dlldef)
  REM --- (generated code: YCellular globals)

  Public Const Y_LINKQUALITY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CELLOPERATOR_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CELLIDENTIFIER_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CELLTYPE_GPRS As Integer = 0
  Public Const Y_CELLTYPE_EGPRS As Integer = 1
  Public Const Y_CELLTYPE_WCDMA As Integer = 2
  Public Const Y_CELLTYPE_HSDPA As Integer = 3
  Public Const Y_CELLTYPE_NONE As Integer = 4
  Public Const Y_CELLTYPE_CDMA As Integer = 5
  Public Const Y_CELLTYPE_LTE_M As Integer = 6
  Public Const Y_CELLTYPE_NB_IOT As Integer = 7
  Public Const Y_CELLTYPE_EC_GSM_IOT As Integer = 8
  Public Const Y_CELLTYPE_INVALID As Integer = -1
  Public Const Y_IMSI_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_MESSAGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PIN_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_RADIOCONFIG_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_LOCKEDOPERATOR_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_AIRPLANEMODE_OFF As Integer = 0
  Public Const Y_AIRPLANEMODE_ON As Integer = 1
  Public Const Y_AIRPLANEMODE_INVALID As Integer = -1
  Public Const Y_ENABLEDATA_HOMENETWORK As Integer = 0
  Public Const Y_ENABLEDATA_ROAMING As Integer = 1
  Public Const Y_ENABLEDATA_NEVER As Integer = 2
  Public Const Y_ENABLEDATA_NEUTRALITY As Integer = 3
  Public Const Y_ENABLEDATA_INVALID As Integer = -1
  Public Const Y_APN_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_APNSECRET_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PINGINTERVAL_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_DATASENT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_DATARECEIVED_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YCellularValueCallback(ByVal func As YCellular, ByVal value As String)
  Public Delegate Sub YCellularTimedReportCallback(ByVal func As YCellular, ByVal measure As YMeasure)
  REM --- (end of generated code: YCellular globals)

  REM --- (generated code: YCellular class start)

  '''*
  ''' <summary>
  '''   The <c>YCellular</c> class provides control over cellular network parameters
  '''   and status for devices that are GSM-enabled.
  ''' <para>
  '''   Note that TCP/IP parameters are configured separately, using class <c>YNetwork</c>.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YCellular
    Inherits YFunction
    REM --- (end of generated code: YCellular class start)

    REM --- (generated code: YCellular definitions)
    Public Const LINKQUALITY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CELLOPERATOR_INVALID As String = YAPI.INVALID_STRING
    Public Const CELLIDENTIFIER_INVALID As String = YAPI.INVALID_STRING
    Public Const CELLTYPE_GPRS As Integer = 0
    Public Const CELLTYPE_EGPRS As Integer = 1
    Public Const CELLTYPE_WCDMA As Integer = 2
    Public Const CELLTYPE_HSDPA As Integer = 3
    Public Const CELLTYPE_NONE As Integer = 4
    Public Const CELLTYPE_CDMA As Integer = 5
    Public Const CELLTYPE_LTE_M As Integer = 6
    Public Const CELLTYPE_NB_IOT As Integer = 7
    Public Const CELLTYPE_EC_GSM_IOT As Integer = 8
    Public Const CELLTYPE_INVALID As Integer = -1
    Public Const IMSI_INVALID As String = YAPI.INVALID_STRING
    Public Const MESSAGE_INVALID As String = YAPI.INVALID_STRING
    Public Const PIN_INVALID As String = YAPI.INVALID_STRING
    Public Const RADIOCONFIG_INVALID As String = YAPI.INVALID_STRING
    Public Const LOCKEDOPERATOR_INVALID As String = YAPI.INVALID_STRING
    Public Const AIRPLANEMODE_OFF As Integer = 0
    Public Const AIRPLANEMODE_ON As Integer = 1
    Public Const AIRPLANEMODE_INVALID As Integer = -1
    Public Const ENABLEDATA_HOMENETWORK As Integer = 0
    Public Const ENABLEDATA_ROAMING As Integer = 1
    Public Const ENABLEDATA_NEVER As Integer = 2
    Public Const ENABLEDATA_NEUTRALITY As Integer = 3
    Public Const ENABLEDATA_INVALID As Integer = -1
    Public Const APN_INVALID As String = YAPI.INVALID_STRING
    Public Const APNSECRET_INVALID As String = YAPI.INVALID_STRING
    Public Const PINGINTERVAL_INVALID As Integer = YAPI.INVALID_UINT
    Public Const DATASENT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const DATARECEIVED_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of generated code: YCellular definitions)

    REM --- (generated code: YCellular attributes declaration)
    Protected _linkQuality As Integer
    Protected _cellOperator As String
    Protected _cellIdentifier As String
    Protected _cellType As Integer
    Protected _imsi As String
    Protected _message As String
    Protected _pin As String
    Protected _radioConfig As String
    Protected _lockedOperator As String
    Protected _airplaneMode As Integer
    Protected _enableData As Integer
    Protected _apn As String
    Protected _apnSecret As String
    Protected _pingInterval As Integer
    Protected _dataSent As Integer
    Protected _dataReceived As Integer
    Protected _command As String
    Protected _valueCallbackCellular As YCellularValueCallback
    REM --- (end of generated code: YCellular attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Cellular"
      REM --- (generated code: YCellular attributes initialization)
      _linkQuality = LINKQUALITY_INVALID
      _cellOperator = CELLOPERATOR_INVALID
      _cellIdentifier = CELLIDENTIFIER_INVALID
      _cellType = CELLTYPE_INVALID
      _imsi = IMSI_INVALID
      _message = MESSAGE_INVALID
      _pin = PIN_INVALID
      _radioConfig = RADIOCONFIG_INVALID
      _lockedOperator = LOCKEDOPERATOR_INVALID
      _airplaneMode = AIRPLANEMODE_INVALID
      _enableData = ENABLEDATA_INVALID
      _apn = APN_INVALID
      _apnSecret = APNSECRET_INVALID
      _pingInterval = PINGINTERVAL_INVALID
      _dataSent = DATASENT_INVALID
      _dataReceived = DATARECEIVED_INVALID
      _command = COMMAND_INVALID
      _valueCallbackCellular = Nothing
      REM --- (end of generated code: YCellular attributes initialization)
    End Sub

    REM --- (generated code: YCellular private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("linkQuality") Then
        _linkQuality = CInt(json_val.getLong("linkQuality"))
      End If
      If json_val.has("cellOperator") Then
        _cellOperator = json_val.getString("cellOperator")
      End If
      If json_val.has("cellIdentifier") Then
        _cellIdentifier = json_val.getString("cellIdentifier")
      End If
      If json_val.has("cellType") Then
        _cellType = CInt(json_val.getLong("cellType"))
      End If
      If json_val.has("imsi") Then
        _imsi = json_val.getString("imsi")
      End If
      If json_val.has("message") Then
        _message = json_val.getString("message")
      End If
      If json_val.has("pin") Then
        _pin = json_val.getString("pin")
      End If
      If json_val.has("radioConfig") Then
        _radioConfig = json_val.getString("radioConfig")
      End If
      If json_val.has("lockedOperator") Then
        _lockedOperator = json_val.getString("lockedOperator")
      End If
      If json_val.has("airplaneMode") Then
        If (json_val.getInt("airplaneMode") > 0) Then _airplaneMode = 1 Else _airplaneMode = 0
      End If
      If json_val.has("enableData") Then
        _enableData = CInt(json_val.getLong("enableData"))
      End If
      If json_val.has("apn") Then
        _apn = json_val.getString("apn")
      End If
      If json_val.has("apnSecret") Then
        _apnSecret = json_val.getString("apnSecret")
      End If
      If json_val.has("pingInterval") Then
        _pingInterval = CInt(json_val.getLong("pingInterval"))
      End If
      If json_val.has("dataSent") Then
        _dataSent = CInt(json_val.getLong("dataSent"))
      End If
      If json_val.has("dataReceived") Then
        _dataReceived = CInt(json_val.getLong("dataReceived"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YCellular private methods declaration)

    REM --- (generated code: YCellular public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the link quality, expressed in percent.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the link quality, expressed in percent
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.LINKQUALITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_linkQuality() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LINKQUALITY_INVALID
        End If
      End If
      res = Me._linkQuality
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the name of the cell operator currently in use.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of the cell operator currently in use
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.CELLOPERATOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_cellOperator() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CELLOPERATOR_INVALID
        End If
      End If
      res = Me._cellOperator
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the unique identifier of the cellular antenna in use: MCC, MNC, LAC and Cell ID.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the unique identifier of the cellular antenna in use: MCC, MNC, LAC and Cell ID
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.CELLIDENTIFIER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_cellIdentifier() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CELLIDENTIFIER_INVALID
        End If
      End If
      res = Me._cellIdentifier
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Active cellular connection type.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YCellular.CELLTYPE_GPRS</c>, <c>YCellular.CELLTYPE_EGPRS</c>,
    '''   <c>YCellular.CELLTYPE_WCDMA</c>, <c>YCellular.CELLTYPE_HSDPA</c>, <c>YCellular.CELLTYPE_NONE</c>,
    '''   <c>YCellular.CELLTYPE_CDMA</c>, <c>YCellular.CELLTYPE_LTE_M</c>, <c>YCellular.CELLTYPE_NB_IOT</c>
    '''   and <c>YCellular.CELLTYPE_EC_GSM_IOT</c>
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.CELLTYPE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_cellType() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CELLTYPE_INVALID
        End If
      End If
      res = Me._cellType
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the International Mobile Subscriber Identity (MSI) that uniquely identifies
    '''   the SIM card.
    ''' <para>
    '''   The first 3 digits represent the mobile country code (MCC), which
    '''   is followed by the mobile network code (MNC), either 2-digit (European standard)
    '''   or 3-digit (North American standard)
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the International Mobile Subscriber Identity (MSI) that uniquely identifies
    '''   the SIM card
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.IMSI_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_imsi() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return IMSI_INVALID
        End If
      End If
      res = Me._imsi
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the latest status message from the wireless interface.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest status message from the wireless interface
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.MESSAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_message() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MESSAGE_INVALID
        End If
      End If
      res = Me._message
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns an opaque string if a PIN code has been configured in the device to access
    '''   the SIM card, or an empty string if none has been configured or if the code provided
    '''   was rejected by the SIM card.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to an opaque string if a PIN code has been configured in the device to access
    '''   the SIM card, or an empty string if none has been configured or if the code provided
    '''   was rejected by the SIM card
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.PIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pin() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PIN_INVALID
        End If
      End If
      res = Me._pin
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the PIN code used by the module to access the SIM card.
    ''' <para>
    '''   This function does not change the code on the SIM card itself, but only changes
    '''   the parameter used by the device to try to get access to it. If the SIM code
    '''   does not work immediately on first try, it will be automatically forgotten
    '''   and the message will be set to "Enter SIM PIN". The method should then be
    '''   invoked again with right correct PIN code. After three failed attempts in a row,
    '''   the message is changed to "Enter SIM PUK" and the SIM card PUK code must be
    '''   provided using method <c>sendPUK</c>.
    ''' </para>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module to save the
    '''   new value in the device flash.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the PIN code used by the module to access the SIM card
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
    Public Function set_pin(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("pin", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the type of protocol used over the serial line, as a string.
    ''' <para>
    '''   Possible values are "Line" for ASCII messages separated by CR and/or LF,
    '''   "Frame:[timeout]ms" for binary messages separated by a delay time,
    '''   "Char" for a continuous ASCII stream or
    '''   "Byte" for a continuous binary stream.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the type of protocol used over the serial line, as a string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.RADIOCONFIG_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_radioConfig() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RADIOCONFIG_INVALID
        End If
      End If
      res = Me._radioConfig
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the type of protocol used over the serial line.
    ''' <para>
    '''   Possible values are "Line" for ASCII messages separated by CR and/or LF,
    '''   "Frame:[timeout]ms" for binary messages separated by a delay time,
    '''   "Char" for a continuous ASCII stream or
    '''   "Byte" for a continuous binary stream.
    '''   The suffix "/[wait]ms" can be added to reduce the transmit rate so that there
    '''   is always at lest the specified number of milliseconds between each bytes sent.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the type of protocol used over the serial line
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
    Public Function set_radioConfig(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("radioConfig", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the name of the only cell operator to use if automatic choice is disabled,
    '''   or an empty string if the SIM card will automatically choose among available
    '''   cell operators.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of the only cell operator to use if automatic choice is disabled,
    '''   or an empty string if the SIM card will automatically choose among available
    '''   cell operators
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.LOCKEDOPERATOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lockedOperator() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LOCKEDOPERATOR_INVALID
        End If
      End If
      res = Me._lockedOperator
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the name of the cell operator to be used.
    ''' <para>
    '''   If the name is an empty
    '''   string, the choice will be made automatically based on the SIM card. Otherwise,
    '''   the selected operator is the only one that will be used.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the name of the cell operator to be used
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
    Public Function set_lockedOperator(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("lockedOperator", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns true if the airplane mode is active (radio turned off).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YCellular.AIRPLANEMODE_OFF</c> or <c>YCellular.AIRPLANEMODE_ON</c>, according to true if
    '''   the airplane mode is active (radio turned off)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.AIRPLANEMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_airplaneMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return AIRPLANEMODE_INVALID
        End If
      End If
      res = Me._airplaneMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the activation state of airplane mode (radio turned off).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YCellular.AIRPLANEMODE_OFF</c> or <c>YCellular.AIRPLANEMODE_ON</c>, according to the
    '''   activation state of airplane mode (radio turned off)
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
    Public Function set_airplaneMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("airplaneMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the condition for enabling IP data services (GPRS).
    ''' <para>
    '''   When data services are disabled, SMS are the only mean of communication.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YCellular.ENABLEDATA_HOMENETWORK</c>, <c>YCellular.ENABLEDATA_ROAMING</c>,
    '''   <c>YCellular.ENABLEDATA_NEVER</c> and <c>YCellular.ENABLEDATA_NEUTRALITY</c> corresponding to the
    '''   condition for enabling IP data services (GPRS)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.ENABLEDATA_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_enableData() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ENABLEDATA_INVALID
        End If
      End If
      res = Me._enableData
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the condition for enabling IP data services (GPRS).
    ''' <para>
    '''   The service can be either fully deactivated, or limited to the SIM home network,
    '''   or enabled for all partner networks (roaming). Caution: enabling data services
    '''   on roaming networks may cause prohibitive communication costs !
    ''' </para>
    ''' <para>
    '''   When data services are disabled, SMS are the only mean of communication.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YCellular.ENABLEDATA_HOMENETWORK</c>, <c>YCellular.ENABLEDATA_ROAMING</c>,
    '''   <c>YCellular.ENABLEDATA_NEVER</c> and <c>YCellular.ENABLEDATA_NEUTRALITY</c> corresponding to the
    '''   condition for enabling IP data services (GPRS)
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
    Public Function set_enableData(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("enableData", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the Access Point Name (APN) to be used, if needed.
    ''' <para>
    '''   When left blank, the APN suggested by the cell operator will be used.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the Access Point Name (APN) to be used, if needed
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.APN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_apn() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return APN_INVALID
        End If
      End If
      res = Me._apn
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Returns the Access Point Name (APN) to be used, if needed.
    ''' <para>
    '''   When left blank, the APN suggested by the cell operator will be used.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
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
    Public Function set_apn(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("apn", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns an opaque string if APN authentication parameters have been configured
    '''   in the device, or an empty string otherwise.
    ''' <para>
    '''   To configure these parameters, use <c>set_apnAuth()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to an opaque string if APN authentication parameters have been configured
    '''   in the device, or an empty string otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.APNSECRET_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_apnSecret() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return APNSECRET_INVALID
        End If
      End If
      res = Me._apnSecret
      Return res
    End Function


    Public Function set_apnSecret(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("apnSecret", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the automated connectivity check interval, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the automated connectivity check interval, in seconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.PINGINTERVAL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pingInterval() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PINGINTERVAL_INVALID
        End If
      End If
      res = Me._pingInterval
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the automated connectivity check interval, in seconds.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the automated connectivity check interval, in seconds
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
    Public Function set_pingInterval(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pingInterval", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of bytes sent so far.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of bytes sent so far
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.DATASENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_dataSent() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DATASENT_INVALID
        End If
      End If
      res = Me._dataSent
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the value of the outgoing data counter.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the value of the outgoing data counter
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
    Public Function set_dataSent(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("dataSent", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of bytes received so far.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of bytes received so far
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YCellular.DATARECEIVED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_dataReceived() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DATARECEIVED_INVALID
        End If
      End If
      res = Me._dataReceived
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the value of the incoming data counter.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the value of the incoming data counter
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
    Public Function set_dataReceived(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("dataReceived", rest_val)
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
    '''   Retrieves a cellular interface for a given identifier.
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
    '''   This function does not require that the cellular interface is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YCellular.isOnline()</c> to test if the cellular interface is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a cellular interface by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the cellular interface, for instance
    '''   <c>YHUBGSM1.cellular</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YCellular</c> object allowing you to drive the cellular interface.
    ''' </returns>
    '''/
    Public Shared Function FindCellular(func As String) As YCellular
      Dim obj As YCellular
      obj = CType(YFunction._FindFromCache("Cellular", func), YCellular)
      If ((obj Is Nothing)) Then
        obj = New YCellular(func)
        YFunction._AddToCache("Cellular", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YCellularValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackCellular = callback
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
      If (Not (Me._valueCallbackCellular Is Nothing)) Then
        Me._valueCallbackCellular(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Sends a PUK code to unlock the SIM card after three failed PIN code attempts, and
    '''   setup a new PIN into the SIM card.
    ''' <para>
    '''   Only ten consecutive tentatives are permitted:
    '''   after that, the SIM card will be blocked permanently without any mean of recovery
    '''   to use it again. Note that after calling this method, you have usually to invoke
    '''   method <c>set_pin()</c> to tell the YoctoHub which PIN to use in the future.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="puk">
    '''   the SIM PUK code
    ''' </param>
    ''' <param name="newPin">
    '''   new PIN code to configure into the SIM card
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function sendPUK(puk As String, newPin As String) As Integer
      Dim gsmMsg As String
      gsmMsg = Me.get_message()
      If Not((gsmMsg).Substring(0, 13) = "Enter SIM PUK") Then
        me._throw(YAPI.INVALID_ARGUMENT,  "PUK not expected at this time")
        return YAPI.INVALID_ARGUMENT
      end if
      If (newPin = "") Then
        Return Me.set_command("AT+CPIN=" + puk + ",0000;+CLCK=SC,0,0000")
      End If
      Return Me.set_command("AT+CPIN=" + puk + "," + newPin)
    End Function

    '''*
    ''' <summary>
    '''   Configure authentication parameters to connect to the APN.
    ''' <para>
    '''   Both
    '''   PAP and CHAP authentication are supported.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="username">
    '''   APN username
    ''' </param>
    ''' <param name="password">
    '''   APN password
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_apnAuth(username As String, password As String) As Integer
      Return Me.set_apnSecret("" + username + "," + password)
    End Function

    '''*
    ''' <summary>
    '''   Clear the transmitted data counters.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function clearDataCounters() As Integer
      Dim retcode As Integer = 0

      retcode = Me.set_dataReceived(0)
      If (retcode <> YAPI.SUCCESS) Then
        Return retcode
      End If
      retcode = Me.set_dataSent(0)
      Return retcode
    End Function

    '''*
    ''' <summary>
    '''   Sends an AT command to the GSM module and returns the command output.
    ''' <para>
    '''   The command will only execute when the GSM module is in standard
    '''   command state, and should leave it in the exact same state.
    '''   Use this function with great care !
    ''' </para>
    ''' </summary>
    ''' <param name="cmd">
    '''   the AT command to execute, like for instance: "+CCLK?".
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   a string with the result of the commands. Empty lines are
    '''   automatically removed from the output.
    ''' </returns>
    '''/
    Public Overridable Function _AT(cmd As String) As String
      Dim chrPos As Integer = 0
      Dim cmdLen As Integer = 0
      Dim waitMore As Integer = 0
      Dim res As String
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim buffstr As String
      Dim buffstrlen As Integer = 0
      Dim idx As Integer = 0
      Dim suffixlen As Integer = 0
      REM // quote dangerous characters used in AT commands
      cmdLen = (cmd).Length
      chrPos = cmd.IndexOf("#")
      While (chrPos >= 0)
        cmd = "" +  (cmd).Substring( 0, chrPos) + "" + Chr( 37) + "23" + (cmd).Substring( chrPos+1, cmdLen-chrPos-1)
        cmdLen = cmdLen + 2
        chrPos = cmd.IndexOf("#")
      End While
      chrPos = cmd.IndexOf("+")
      While (chrPos >= 0)
        cmd = "" +  (cmd).Substring( 0, chrPos) + "" + Chr( 37) + "2B" + (cmd).Substring( chrPos+1, cmdLen-chrPos-1)
        cmdLen = cmdLen + 2
        chrPos = cmd.IndexOf("+")
      End While
      chrPos = cmd.IndexOf("=")
      While (chrPos >= 0)
        cmd = "" +  (cmd).Substring( 0, chrPos) + "" + Chr( 37) + "3D" + (cmd).Substring( chrPos+1, cmdLen-chrPos-1)
        cmdLen = cmdLen + 2
        chrPos = cmd.IndexOf("=")
      End While
      cmd = "at.txt?cmd=" + cmd
      res = ""
      REM // max 2 minutes (each iteration may take up to 5 seconds if waiting)
      waitMore = 24
      While (waitMore > 0)
        buff = Me._download(cmd)
        bufflen = (buff).Length
        buffstr = YAPI.DefaultEncoding.GetString(buff)
        buffstrlen = (buffstr).Length
        idx = bufflen - 1
        While ((idx > 0) AndAlso (buff(idx) <> 64) AndAlso (buff(idx) <> 10) AndAlso (buff(idx) <> 13))
          idx = idx - 1
        End While
        If (buff(idx) = 64) Then
          REM // continuation detected
          suffixlen = bufflen - idx
          cmd = "at.txt?cmd=" + (buffstr).Substring( buffstrlen - suffixlen, suffixlen)
          buffstr = (buffstr).Substring( 0, buffstrlen - suffixlen)
          waitMore = waitMore - 1
        Else
          REM // request complete
          waitMore = 0
        End If
        res = "" +  res + "" + buffstr
      End While
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the list detected cell operators in the neighborhood.
    ''' <para>
    '''   This function will typically take between 30 seconds to 1 minute to
    '''   return. Note that any SIM card can usually only connect to specific
    '''   operators. All networks returned by this function might therefore
    '''   not be available for connection.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of string (cell operator names).
    ''' </returns>
    '''/
    Public Overridable Function get_availableOperators() As List(Of String)
      Dim cops As String
      Dim idx As Integer = 0
      Dim slen As Integer = 0
      Dim res As List(Of String) = New List(Of String)()

      cops = Me._AT("+COPS=?")
      slen = (cops).Length
      res.Clear()
      idx = cops.IndexOf("(")
      While (idx >= 0)
        slen = slen - (idx+1)
        cops = (cops).Substring( idx+1, slen)
        idx = cops.IndexOf("""")
        If (idx > 0) Then
          slen = slen - (idx+1)
          cops = (cops).Substring( idx+1, slen)
          idx = cops.IndexOf("""")
          If (idx > 0) Then
            res.Add((cops).Substring( 0, idx))
          End If
        End If
        idx = cops.IndexOf("(")
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of nearby cellular antennas, as required for quick
    '''   geolocation of the device.
    ''' <para>
    '''   The first cell listed is the serving
    '''   cell, and the next ones are the neighbor cells reported by the
    '''   serving cell.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of <c>YCellRecords</c>.
    ''' </returns>
    '''/
    Public Overridable Function quickCellSurvey() As List(Of YCellRecord)
      Dim moni As String
      Dim recs As List(Of String) = New List(Of String)()
      Dim llen As Integer = 0
      Dim mccs As String
      Dim mcc As Integer = 0
      Dim mncs As String
      Dim mnc As Integer = 0
      Dim lac As Integer = 0
      Dim cellId As Integer = 0
      Dim dbms As String
      Dim dbm As Integer = 0
      Dim tads As String
      Dim tad As Integer = 0
      Dim oper As String
      Dim res As List(Of YCellRecord) = New List(Of YCellRecord)()

      moni = Me._AT("+CCED=0;#MONI=7;#MONI")
      mccs = (moni).Substring(7, 3)
      If ((mccs).Substring(0, 1) = "0") Then
        mccs = (mccs).Substring(1, 2)
      End If
      If ((mccs).Substring(0, 1) = "0") Then
        mccs = (mccs).Substring(1, 1)
      End If
      mcc = YAPI._atoi(mccs)
      mncs = (moni).Substring(11, 3)
      If ((mncs).Substring(2, 1) = ",") Then
        mncs = (mncs).Substring(0, 2)
      End If
      If ((mncs).Substring(0, 1) = "0") Then
        mncs = (mncs).Substring(1, (mncs).Length-1)
      End If
      mnc = YAPI._atoi(mncs)
      recs = New List(Of String)(moni.Split(New Char() {"#"c}))
      REM // process each line in turn
      res.Clear()
      Dim ii_0 As Integer
      For ii_0 = 0 To recs.Count - 1
        llen = (recs(ii_0)).Length - 2
        If (llen >= 44) Then
          If ((recs(ii_0)).Substring(41, 3) = "dbm") Then
            lac = Convert.ToInt32((recs(ii_0)).Substring(16, 4), 16)
            cellId = Convert.ToInt32((recs(ii_0)).Substring(23, 4), 16)
            dbms = (recs(ii_0)).Substring(37, 4)
            If ((dbms).Substring(0, 1) = " ") Then
              dbms = (dbms).Substring(1, 3)
            End If
            dbm = YAPI._atoi(dbms)
            If (llen > 66) Then
              tads = (recs(ii_0)).Substring(54, 2)
              If ((tads).Substring(0, 1) = " ") Then
                tads = (tads).Substring(1, 3)
              End If
              tad = YAPI._atoi(tads)
              oper = (recs(ii_0)).Substring(66, llen-66)
            Else
              tad = -1
              oper = ""
            End If
            If (lac < 65535) Then
              res.Add(New YCellRecord(mcc, mnc, lac, cellId, dbm, tad, oper))
            End If
          End If
        End If
      Next ii_0
      Return res
    End Function

    Public Overridable Function imm_decodePLMN(mccmnc As String) As String
      Dim inputlen As Integer = 0
      Dim mcc As Integer = 0
      Dim npos As Integer = 0
      Dim nval As Integer = 0
      Dim ch As String
      Dim plmnid As Integer = 0
      REM // Make sure we have a valid MCC/MNC pair
      inputlen = (mccmnc).Length
      If (inputlen < 5) Then
        Return mccmnc
      End If
      mcc = YAPI._atoi((mccmnc).Substring(0, 3))
      If (mcc < 200) Then
        Return mccmnc
      End If
      If ((mccmnc).Substring(3, 1) = " ") Then
        npos = 4
      Else
        npos = 3
      End If
      plmnid = mcc
      While (plmnid < 100000 AndAlso npos < inputlen)
        ch = (mccmnc).Substring(npos, 1)
        nval = YAPI._atoi(ch)
        If (ch = (nval).ToString()) Then
          plmnid = plmnid * 10 + nval
          npos = npos + 1
        Else
          npos = inputlen
        End If
      End While
      REM // Search for PLMN operator brand, if known
      If (plmnid < 20201) Then
        Return mccmnc
      End If
      If (plmnid < 50503) Then
        If (plmnid < 40407) Then
          If (plmnid < 25008) Then
            If (plmnid < 23102) Then
              If (plmnid < 21601) Then
                If (plmnid < 20809) Then
                  If (plmnid < 20408) Then
                    If (plmnid < 20210) Then
                      If (plmnid = 20201) Then
                        Return "Cosmote"
                      End If
                      If (plmnid = 20202) Then
                        Return "Cosmote"
                      End If
                      If (plmnid = 20205) Then
                        Return "Vodafone GR"
                      End If
                      If (plmnid = 20209) Then
                        Return "Wind GR"
                      End If
                    Else
                      If (plmnid = 20210) Then
                        Return "Wind GR"
                      End If
                      If (plmnid = 20402) Then
                        Return "Tele2 NL"
                      End If
                      If (plmnid = 20403) Then
                        Return "Voiceworks"
                      End If
                      If (plmnid = 20404) Then
                        Return "Vodafone NL"
                      End If
                    End If
                  Else
                    If (plmnid < 20601) Then
                      If (plmnid = 20408) Then
                        Return "KPN"
                      End If
                      If (plmnid = 20410) Then
                        Return "KPN"
                      End If
                      If (plmnid = 20416) Then
                        Return "T-Mobile (BEN)"
                      End If
                      If (plmnid = 20420) Then
                        Return "T-Mobile NL"
                      End If
                    Else
                      If (plmnid < 20620) Then
                        If (plmnid = 20601) Then
                          Return "Proximus"
                        End If
                        If (plmnid = 20610) Then
                          Return "Orange Belgium"
                        End If
                      Else
                        If (plmnid = 20620) Then
                          Return "Base"
                        End If
                        If (plmnid = 20801) Then
                          Return "Orange FR"
                        End If
                        If (plmnid = 20802) Then
                          Return "Orange FR"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 20836) Then
                    If (plmnid < 20815) Then
                      If (plmnid = 20809) Then
                        Return "SFR"
                      End If
                      If (plmnid = 20810) Then
                        Return "SFR"
                      End If
                      If (plmnid = 20813) Then
                        Return "SFR"
                      End If
                      If (plmnid = 20814) Then
                        Return "SNCF Réseau"
                      End If
                    Else
                      If (plmnid = 20815) Then
                        Return "Free FR"
                      End If
                      If (plmnid = 20816) Then
                        Return "Free FR"
                      End If
                      If (plmnid = 20820) Then
                        Return "Bouygues"
                      End If
                      If (plmnid = 20835) Then
                        Return "Free FR"
                      End If
                    End If
                  Else
                    If (plmnid < 21401) Then
                      If (plmnid = 20836) Then
                        Return "Free FR"
                      End If
                      If (plmnid = 20888) Then
                        Return "Bouygues"
                      End If
                      If (plmnid = 21210) Then
                        Return "Office des Telephones"
                      End If
                      If (plmnid = 21303) Then
                        Return "Som, Mobiland"
                      End If
                    Else
                      If (plmnid < 21404) Then
                        If (plmnid = 21401) Then
                          Return "Vodafone ES"
                        End If
                        If (plmnid = 21403) Then
                          Return "Orange ES"
                        End If
                      Else
                        If (plmnid = 21404) Then
                          Return "Yoigo"
                        End If
                        If (plmnid = 21407) Then
                          Return "Movistar ES"
                        End If
                        If (plmnid = 21451) Then
                          Return "ADIF"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 22210) Then
                  If (plmnid < 21901) Then
                    If (plmnid < 21699) Then
                      If (plmnid = 21601) Then
                        Return "Telenor Hungary"
                      End If
                      If (plmnid = 21603) Then
                        Return "DIGI"
                      End If
                      If (plmnid = 21630) Then
                        Return "Telekom HU"
                      End If
                      If (plmnid = 21670) Then
                        Return "Vodafone HU"
                      End If
                    Else
                      If (plmnid = 21699) Then
                        Return "MAV GSM-R"
                      End If
                      If (plmnid = 21803) Then
                        Return "HT-ERONET"
                      End If
                      If (plmnid = 21805) Then
                        Return "m:tel BiH"
                      End If
                      If (plmnid = 21890) Then
                        Return "BH Mobile"
                      End If
                    End If
                  Else
                    If (plmnid < 22003) Then
                      If (plmnid = 21901) Then
                        Return "T-Mobile HR"
                      End If
                      If (plmnid = 21902) Then
                        Return "Tele2 HR"
                      End If
                      If (plmnid = 21910) Then
                        Return "A1 HR"
                      End If
                      If (plmnid = 22001) Then
                        Return "Telenor RS"
                      End If
                    Else
                      If (plmnid < 22101) Then
                        If (plmnid = 22003) Then
                          Return "mts"
                        End If
                        If (plmnid = 22005) Then
                          Return "VIP"
                        End If
                      Else
                        If (plmnid = 22101) Then
                          Return "Vala"
                        End If
                        If (plmnid = 22102) Then
                          Return "IPKO"
                        End If
                        If (plmnid = 22201) Then
                          Return "TIM IT"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 22801) Then
                    If (plmnid < 22299) Then
                      If (plmnid = 22210) Then
                        Return "Vodafone IT"
                      End If
                      If (plmnid = 22230) Then
                        Return "RFI"
                      End If
                      If (plmnid = 22250) Then
                        Return "Iliad"
                      End If
                      If (plmnid = 22288) Then
                        Return "Wind IT"
                      End If
                    Else
                      If (plmnid < 22603) Then
                        If (plmnid = 22299) Then
                          Return "3 Italia"
                        End If
                        If (plmnid = 22601) Then
                          Return "Vodafone RO"
                        End If
                      Else
                        If (plmnid = 22603) Then
                          Return "Telekom RO"
                        End If
                        If (plmnid = 22605) Then
                          Return "Digi.Mobil"
                        End If
                        If (plmnid = 22610) Then
                          Return "Orange RO"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 22808) Then
                      If (plmnid = 22801) Then
                        Return "Swisscom CH"
                      End If
                      If (plmnid = 22802) Then
                        Return "Sunrise"
                      End If
                      If (plmnid = 22803) Then
                        Return "Salt"
                      End If
                      If (plmnid = 22806) Then
                        Return "SBB-CFF-FFS"
                      End If
                    Else
                      If (plmnid < 23002) Then
                        If (plmnid = 22808) Then
                          Return "Tele4u"
                        End If
                        If (plmnid = 23001) Then
                          Return "T-Mobile CZ"
                        End If
                      Else
                        If (plmnid = 23002) Then
                          Return "O2 CZ"
                        End If
                        If (plmnid = 23003) Then
                          Return "Vodafone CZ"
                        End If
                        If (plmnid = 23101) Then
                          Return "Orange SK"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If (plmnid < 23458) Then
                If (plmnid < 23411) Then
                  If (plmnid < 23205) Then
                    If (plmnid < 23106) Then
                      If (plmnid = 23102) Then
                        Return "Telekom SK"
                      End If
                      If (plmnid = 23103) Then
                        Return "4ka"
                      End If
                      If (plmnid = 23104) Then
                        Return "Telekom SK"
                      End If
                      If (plmnid = 23105) Then
                        Return "Orange SK"
                      End If
                    Else
                      If (plmnid = 23106) Then
                        Return "O2 SK"
                      End If
                      If (plmnid = 23199) Then
                        Return "?SR"
                      End If
                      If (plmnid = 23201) Then
                        Return "A1.net"
                      End If
                      If (plmnid = 23203) Then
                        Return "Magenta"
                      End If
                    End If
                  Else
                    If (plmnid < 23402) Then
                      If (plmnid = 23205) Then
                        Return "3 AT"
                      End If
                      If (plmnid = 23210) Then
                        Return "3 AT"
                      End If
                      If (plmnid = 23291) Then
                        Return "GSM-R A"
                      End If
                      If (plmnid = 23400) Then
                        Return "BT"
                      End If
                    Else
                      If (plmnid < 23403) Then
                        If (plmnid = 23402) Then
                          Return "O2 (UK)"
                        End If
                        If (plmnid = 23403) Then
                          Return "Airtel-Vodafone GG"
                        End If
                      Else
                        If (plmnid = 23403) Then
                          Return "Airtel-Vodafone JE"
                        End If
                        If (plmnid = 23403) Then
                          Return "Airtel-Vodafone GB"
                        End If
                        If (plmnid = 23410) Then
                          Return "O2 (UK)"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 23436) Then
                    If (plmnid < 23419) Then
                      If (plmnid = 23411) Then
                        Return "O2 (UK)"
                      End If
                      If (plmnid = 23412) Then
                        Return "Railtrack"
                      End If
                      If (plmnid = 23413) Then
                        Return "Railtrack"
                      End If
                      If (plmnid = 23415) Then
                        Return "Vodafone UK"
                      End If
                    Else
                      If (plmnid < 23430) Then
                        If (plmnid = 23419) Then
                          Return "Private Mobile Networks PMN"
                        End If
                        If (plmnid = 23420) Then
                          Return "3 GB"
                        End If
                      Else
                        If (plmnid = 23430) Then
                          Return "T-Mobile UK"
                        End If
                        If (plmnid = 23433) Then
                          Return "Orange GB"
                        End If
                        If (plmnid = 23434) Then
                          Return "Orange GB"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 23450) Then
                      If (plmnid = 23436) Then
                        Return "Sure Mobile IM"
                      End If
                      If (plmnid = 23436) Then
                        Return "Sure Mobile GB"
                      End If
                      If (plmnid = 23450) Then
                        Return "JT GG"
                      End If
                      If (plmnid = 23450) Then
                        Return "JT JE"
                      End If
                    Else
                      If (plmnid < 23455) Then
                        If (plmnid = 23450) Then
                          Return "JT GB"
                        End If
                        If (plmnid = 23455) Then
                          Return "Sure Mobile GG"
                        End If
                      Else
                        If (plmnid = 23455) Then
                          Return "Sure Mobile JE"
                        End If
                        If (plmnid = 23455) Then
                          Return "Sure Mobile GB"
                        End If
                        If (plmnid = 23458) Then
                          Return "Pronto GSM IM"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 24405) Then
                  If (plmnid < 24001) Then
                    If (plmnid < 23806) Then
                      If (plmnid = 23458) Then
                        Return "Pronto GSM GB"
                      End If
                      If (plmnid = 23476) Then
                        Return "BT"
                      End If
                      If (plmnid = 23801) Then
                        Return "TDC"
                      End If
                      If (plmnid = 23802) Then
                        Return "Telenor DK"
                      End If
                    Else
                      If (plmnid = 23806) Then
                        Return "3 DK"
                      End If
                      If (plmnid = 23820) Then
                        Return "Telia DK"
                      End If
                      If (plmnid = 23823) Then
                        Return "GSM-R DK"
                      End If
                      If (plmnid = 23877) Then
                        Return "Telenor DK"
                      End If
                    End If
                  Else
                    If (plmnid < 24024) Then
                      If (plmnid = 24001) Then
                        Return "Telia SE"
                      End If
                      If (plmnid = 24002) Then
                        Return "3 SE"
                      End If
                      If (plmnid = 24007) Then
                        Return "Tele2 SE"
                      End If
                      If (plmnid = 24021) Then
                        Return "MobiSir"
                      End If
                    Else
                      If (plmnid < 24202) Then
                        If (plmnid = 24024) Then
                          Return "Sweden 2G"
                        End If
                        If (plmnid = 24201) Then
                          Return "Telenor NO"
                        End If
                      Else
                        If (plmnid = 24202) Then
                          Return "Telia NO"
                        End If
                        If (plmnid = 24214) Then
                          Return "ice"
                        End If
                        If (plmnid = 24403) Then
                          Return "DNA"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 24605) Then
                    If (plmnid < 24436) Then
                      If (plmnid = 24405) Then
                        Return "Elisa FI"
                      End If
                      If (plmnid = 24407) Then
                        Return "Nokia"
                      End If
                      If (plmnid = 24412) Then
                        Return "DNA"
                      End If
                      If (plmnid = 24414) Then
                        Return "Ålcom"
                      End If
                    Else
                      If (plmnid < 24601) Then
                        If (plmnid = 24436) Then
                          Return "Telia / DNA"
                        End If
                        If (plmnid = 24491) Then
                          Return "Telia FI"
                        End If
                      Else
                        If (plmnid = 24601) Then
                          Return "Telia LT"
                        End If
                        If (plmnid = 24602) Then
                          Return "BIT?"
                        End If
                        If (plmnid = 24603) Then
                          Return "Tele2 LT"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 24801) Then
                      If (plmnid = 24605) Then
                        Return "LitRail"
                      End If
                      If (plmnid = 24701) Then
                        Return "LMT"
                      End If
                      If (plmnid = 24702) Then
                        Return "Tele2 LV"
                      End If
                      If (plmnid = 24705) Then
                        Return "Bite"
                      End If
                    Else
                      If (plmnid < 24803) Then
                        If (plmnid = 24801) Then
                          Return "Telia EE"
                        End If
                        If (plmnid = 24802) Then
                          Return "Elisa EE"
                        End If
                      Else
                        If (plmnid = 24803) Then
                          Return "Tele2 EE"
                        End If
                        If (plmnid = 25001) Then
                          Return "MTS RU"
                        End If
                        If (plmnid = 25002) Then
                          Return "MegaFon RU"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          Else
            If (plmnid < 28405) Then
              If (plmnid < 26801) Then
                If (plmnid < 25706) Then
                  If (plmnid < 25099) Then
                    If (plmnid < 25033) Then
                      If (plmnid = 25008) Then
                        Return "Vainah Telecom"
                      End If
                      If (plmnid = 25020) Then
                        Return "Tele2 RU"
                      End If
                      If (plmnid = 25027) Then
                        Return "Letai"
                      End If
                      If (plmnid = 25032) Then
                        Return "Win Mobile"
                      End If
                    Else
                      If (plmnid = 25033) Then
                        Return "Sevmobile"
                      End If
                      If (plmnid = 25034) Then
                        Return "Krymtelekom"
                      End If
                      If (plmnid = 25035) Then
                        Return "MOTIV"
                      End If
                      If (plmnid = 25060) Then
                        Return "Volna mobile"
                      End If
                    End If
                  Else
                    If (plmnid < 25506) Then
                      If (plmnid = 25099) Then
                        Return "Beeline RU"
                      End If
                      If (plmnid = 25501) Then
                        Return "Vodafone UA"
                      End If
                      If (plmnid = 25502) Then
                        Return "Kyivstar"
                      End If
                      If (plmnid = 25503) Then
                        Return "Kyivstar"
                      End If
                    Else
                      If (plmnid < 25701) Then
                        If (plmnid = 25506) Then
                          Return "lifecell"
                        End If
                        If (plmnid = 25599) Then
                          Return "Phoenix UA"
                        End If
                      Else
                        If (plmnid = 25701) Then
                          Return "A1 BY"
                        End If
                        If (plmnid = 25702) Then
                          Return "MTS BY"
                        End If
                        If (plmnid = 25704) Then
                          Return "life:)"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 26015) Then
                    If (plmnid < 25915) Then
                      If (plmnid = 25706) Then
                        Return "beCloud"
                      End If
                      If (plmnid = 25901) Then
                        Return "Orange MD"
                      End If
                      If (plmnid = 25902) Then
                        Return "Moldcell"
                      End If
                      If (plmnid = 25905) Then
                        Return "Unité"
                      End If
                    Else
                      If (plmnid < 26002) Then
                        If (plmnid = 25915) Then
                          Return "IDC"
                        End If
                        If (plmnid = 26001) Then
                          Return "Plus"
                        End If
                      Else
                        If (plmnid = 26002) Then
                          Return "T-Mobile PL"
                        End If
                        If (plmnid = 26003) Then
                          Return "Orange PL"
                        End If
                        If (plmnid = 26006) Then
                          Return "Play"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 26202) Then
                      If (plmnid = 26015) Then
                        Return "Aero2"
                      End If
                      If (plmnid = 26016) Then
                        Return "Aero2"
                      End If
                      If (plmnid = 26034) Then
                        Return "NetWorkS!"
                      End If
                      If (plmnid = 26201) Then
                        Return "Telekom DE"
                      End If
                    Else
                      If (plmnid < 26209) Then
                        If (plmnid = 26202) Then
                          Return "Vodafone DE"
                        End If
                        If (plmnid = 26203) Then
                          Return "O2 DE"
                        End If
                      Else
                        If (plmnid = 26209) Then
                          Return "Vodafone DE"
                        End If
                        If (plmnid = 26601) Then
                          Return "GibTel"
                        End If
                        If (plmnid = 26609) Then
                          Return "Shine"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 27601) Then
                  If (plmnid < 27202) Then
                    If (plmnid < 27071) Then
                      If (plmnid = 26801) Then
                        Return "Vodafone PT"
                      End If
                      If (plmnid = 26803) Then
                        Return "NOS"
                      End If
                      If (plmnid = 26806) Then
                        Return "MEO"
                      End If
                      If (plmnid = 27001) Then
                        Return "POST"
                      End If
                    Else
                      If (plmnid = 27071) Then
                        Return "CFL"
                      End If
                      If (plmnid = 27077) Then
                        Return "Tango"
                      End If
                      If (plmnid = 27099) Then
                        Return "Orange LU"
                      End If
                      If (plmnid = 27201) Then
                        Return "Vodafone IE"
                      End If
                    End If
                  Else
                    If (plmnid < 27401) Then
                      If (plmnid = 27202) Then
                        Return "3 IE"
                      End If
                      If (plmnid = 27203) Then
                        Return "Eir"
                      End If
                      If (plmnid = 27205) Then
                        Return "3 IE"
                      End If
                      If (plmnid = 27207) Then
                        Return "Eir"
                      End If
                    Else
                      If (plmnid < 27404) Then
                        If (plmnid = 27401) Then
                          Return "Síminn"
                        End If
                        If (plmnid = 27402) Then
                          Return "Vodafone IS"
                        End If
                      Else
                        If (plmnid = 27404) Then
                          Return "Viking"
                        End If
                        If (plmnid = 27408) Then
                          Return "On-waves"
                        End If
                        If (plmnid = 27411) Then
                          Return "Nova"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 28201) Then
                    If (plmnid < 27821) Then
                      If (plmnid = 27601) Then
                        Return "Telekom.al"
                      End If
                      If (plmnid = 27602) Then
                        Return "Vodafone AL"
                      End If
                      If (plmnid = 27603) Then
                        Return "Eagle Mobile"
                      End If
                      If (plmnid = 27801) Then
                        Return "Vodafone MT"
                      End If
                    Else
                      If (plmnid < 28001) Then
                        If (plmnid = 27821) Then
                          Return "GO"
                        End If
                        If (plmnid = 27877) Then
                          Return "Melita"
                        End If
                      Else
                        If (plmnid = 28001) Then
                          Return "Cytamobile-Vodafone"
                        End If
                        If (plmnid = 28010) Then
                          Return "Epic"
                        End If
                        If (plmnid = 28020) Then
                          Return "PrimeTel"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 28304) Then
                      If (plmnid = 28201) Then
                        Return "Geocell"
                      End If
                      If (plmnid = 28202) Then
                        Return "Magti"
                      End If
                      If (plmnid = 28204) Then
                        Return "Beeline GE"
                      End If
                      If (plmnid = 28301) Then
                        Return "Beeline AM"
                      End If
                    Else
                      If (plmnid < 28310) Then
                        If (plmnid = 28304) Then
                          Return "Karabakh Telecom"
                        End If
                        If (plmnid = 28305) Then
                          Return "VivaCell-MTS"
                        End If
                      Else
                        If (plmnid = 28310) Then
                          Return "Ucom"
                        End If
                        If (plmnid = 28401) Then
                          Return "A1 BG"
                        End If
                        If (plmnid = 28403) Then
                          Return "Vivacom"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If (plmnid < 36251) Then
                If (plmnid < 29501) Then
                  If (plmnid < 28967) Then
                    If (plmnid < 28602) Then
                      If (plmnid = 28405) Then
                        Return "Telenor BG"
                      End If
                      If (plmnid = 28407) Then
                        Return "????"
                      End If
                      If (plmnid = 28413) Then
                        Return "??.???"
                      End If
                      If (plmnid = 28601) Then
                        Return "Turkcell"
                      End If
                    Else
                      If (plmnid = 28602) Then
                        Return "Vodafone TR"
                      End If
                      If (plmnid = 28603) Then
                        Return "Türk Telekom"
                      End If
                      If (plmnid = 28801) Then
                        Return "Føroya Tele"
                      End If
                      If (plmnid = 28802) Then
                        Return "Hey"
                      End If
                    End If
                  Else
                    If (plmnid < 29341) Then
                      If (plmnid = 28967) Then
                        Return "Aquafon"
                      End If
                      If (plmnid = 28988) Then
                        Return "A-Mobile"
                      End If
                      If (plmnid = 29201) Then
                        Return "PRIMA"
                      End If
                      If (plmnid = 29340) Then
                        Return "A1 SI"
                      End If
                    Else
                      If (plmnid < 29401) Then
                        If (plmnid = 29341) Then
                          Return "Mobitel SI"
                        End If
                        If (plmnid = 29370) Then
                          Return "Telemach"
                        End If
                      Else
                        If (plmnid = 29401) Then
                          Return "Telekom.mk"
                        End If
                        If (plmnid = 29402) Then
                          Return "vip"
                        End If
                        If (plmnid = 29403) Then
                          Return "vip"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 34001) Then
                    If (plmnid < 29702) Then
                      If (plmnid = 29501) Then
                        Return "Swisscom LI"
                      End If
                      If (plmnid = 29502) Then
                        Return "7acht"
                      End If
                      If (plmnid = 29505) Then
                        Return "FL1"
                      End If
                      If (plmnid = 29701) Then
                        Return "Telenor ME"
                      End If
                    Else
                      If (plmnid < 30801) Then
                        If (plmnid = 29702) Then
                          Return "T-Mobile ME"
                        End If
                        If (plmnid = 29703) Then
                          Return "m:tel CG"
                        End If
                      Else
                        If (plmnid = 30801) Then
                          Return "Ameris"
                        End If
                        If (plmnid = 30802) Then
                          Return "GLOBALTEL"
                        End If
                        If (plmnid = 34001) Then
                          Return "Orange BL/GF/GP/MF/MQ"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 34008) Then
                      If (plmnid = 34001) Then
                        Return "Orange GF"
                      End If
                      If (plmnid = 34002) Then
                        Return "SFR Caraïbe BL/GF/GP/MF/MQ"
                      End If
                      If (plmnid = 34002) Then
                        Return "SFR Caraïbe GF"
                      End If
                      If (plmnid = 34003) Then
                        Return "Chippie BL/GF/GP/MF/MQ"
                      End If
                    Else
                      If (plmnid < 34020) Then
                        If (plmnid = 34008) Then
                          Return "Dauphin"
                        End If
                        If (plmnid = 34020) Then
                          Return "Digicel BL/GF/GP/MF/MQ"
                        End If
                      Else
                        If (plmnid = 34020) Then
                          Return "Digicel GF"
                        End If
                        If (plmnid = 35000) Then
                          Return "One"
                        End If
                        If (plmnid = 35002) Then
                          Return "Mobility"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 37202) Then
                  If (plmnid < 36291) Then
                    If (plmnid < 36268) Then
                      If (plmnid = 36251) Then
                        Return "Telcell"
                      End If
                      If (plmnid = 36254) Then
                        Return "ECC"
                      End If
                      If (plmnid = 36259) Then
                        Return "Chippie BQ/CW/SX"
                      End If
                      If (plmnid = 36260) Then
                        Return "Chippie BQ/CW/SX"
                      End If
                    Else
                      If (plmnid = 36268) Then
                        Return "Digicel BQ/CW/SX"
                      End If
                      If (plmnid = 36269) Then
                        Return "Digicel BQ/CW/SX"
                      End If
                      If (plmnid = 36276) Then
                        Return "Digicel BQ/CW/SX"
                      End If
                      If (plmnid = 36278) Then
                        Return "Telbo"
                      End If
                    End If
                  Else
                    If (plmnid < 36449) Then
                      If (plmnid = 36291) Then
                        Return "Chippie BQ/CW/SX"
                      End If
                      If (plmnid = 36301) Then
                        Return "SETAR"
                      End If
                      If (plmnid = 36302) Then
                        Return "Digicel AW"
                      End If
                      If (plmnid = 36439) Then
                        Return "BTC"
                      End If
                    Else
                      If (plmnid < 37001) Then
                        If (plmnid = 36449) Then
                          Return "Aliv"
                        End If
                        If (plmnid = 36801) Then
                          Return "CUBACEL"
                        End If
                      Else
                        If (plmnid = 37001) Then
                          Return "Altice"
                        End If
                        If (plmnid = 37002) Then
                          Return "Claro DO"
                        End If
                        If (plmnid = 37004) Then
                          Return "Viva DO"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 40107) Then
                    If (plmnid < 40002) Then
                      If (plmnid = 37202) Then
                        Return "Digicel HT"
                      End If
                      If (plmnid = 37203) Then
                        Return "Natcom"
                      End If
                      If (plmnid = 37412) Then
                        Return "bmobile TT"
                      End If
                      If (plmnid = 40001) Then
                        Return "Azercell"
                      End If
                    Else
                      If (plmnid < 40006) Then
                        If (plmnid = 40002) Then
                          Return "Bakcell"
                        End If
                        If (plmnid = 40004) Then
                          Return "Nar Mobile"
                        End If
                      Else
                        If (plmnid = 40006) Then
                          Return "Naxtel"
                        End If
                        If (plmnid = 40101) Then
                          Return "Beeline KZ"
                        End If
                        If (plmnid = 40102) Then
                          Return "Kcell"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 40401) Then
                      If (plmnid = 40107) Then
                        Return "Altel"
                      End If
                      If (plmnid = 40177) Then
                        Return "Tele2.kz"
                      End If
                      If (plmnid = 40211) Then
                        Return "B-Mobile"
                      End If
                      If (plmnid = 40277) Then
                        Return "TashiCell"
                      End If
                    Else
                      If (plmnid < 40403) Then
                        If (plmnid = 40401) Then
                          Return "Vodafone India"
                        End If
                        If (plmnid = 40402) Then
                          Return "AirTel"
                        End If
                      Else
                        If (plmnid = 40403) Then
                          Return "AirTel"
                        End If
                        If (plmnid = 40404) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 40405) Then
                          Return "Vodafone India"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        Else
          If (plmnid < 42001) Then
            If (plmnid < 40493) Then
              If (plmnid < 40450) Then
                If (plmnid < 40427) Then
                  If (plmnid < 40416) Then
                    If (plmnid < 40412) Then
                      If (plmnid = 40407) Then
                        Return "IDEA"
                      End If
                      If (plmnid = 40409) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40410) Then
                        Return "AirTel"
                      End If
                      If (plmnid = 40411) Then
                        Return "Vodafone India"
                      End If
                    Else
                      If (plmnid = 40412) Then
                        Return "IDEA"
                      End If
                      If (plmnid = 40413) Then
                        Return "Vodafone India"
                      End If
                      If (plmnid = 40414) Then
                        Return "IDEA"
                      End If
                      If (plmnid = 40415) Then
                        Return "Vodafone India"
                      End If
                    End If
                  Else
                    If (plmnid < 40420) Then
                      If (plmnid = 40416) Then
                        Return "Airtel IN"
                      End If
                      If (plmnid = 40417) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 40418) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40419) Then
                        Return "IDEA"
                      End If
                    Else
                      If (plmnid < 40422) Then
                        If (plmnid = 40420) Then
                          Return "Vodafone India"
                        End If
                        If (plmnid = 40421) Then
                          Return "Loop Mobile"
                        End If
                      Else
                        If (plmnid = 40422) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 40424) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 40425) Then
                          Return "AIRCEL"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 40438) Then
                    If (plmnid < 40431) Then
                      If (plmnid = 40427) Then
                        Return "Vodafone India"
                      End If
                      If (plmnid = 40428) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 40429) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 40430) Then
                        Return "Vodafone India"
                      End If
                    Else
                      If (plmnid < 40435) Then
                        If (plmnid = 40431) Then
                          Return "AirTel"
                        End If
                        If (plmnid = 40434) Then
                          Return "cellone"
                        End If
                      Else
                        If (plmnid = 40435) Then
                          Return "Aircel"
                        End If
                        If (plmnid = 40436) Then
                          Return "Reliance"
                        End If
                        If (plmnid = 40437) Then
                          Return "Aircel"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 40444) Then
                      If (plmnid = 40438) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40441) Then
                        Return "Aircel"
                      End If
                      If (plmnid = 40442) Then
                        Return "Aircel"
                      End If
                      If (plmnid = 40443) Then
                        Return "Vodafone India"
                      End If
                    Else
                      If (plmnid < 40446) Then
                        If (plmnid = 40444) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 40445) Then
                          Return "Airtel IN"
                        End If
                      Else
                        If (plmnid = 40446) Then
                          Return "Vodafone India"
                        End If
                        If (plmnid = 40448) Then
                          Return "Dishnet Wireless"
                        End If
                        If (plmnid = 40449) Then
                          Return "Airtel IN"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 40471) Then
                  If (plmnid < 40458) Then
                    If (plmnid < 40454) Then
                      If (plmnid = 40450) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40451) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40452) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40453) Then
                        Return "cellone"
                      End If
                    Else
                      If (plmnid = 40454) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40455) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40456) Then
                        Return "IDEA"
                      End If
                      If (plmnid = 40457) Then
                        Return "cellone"
                      End If
                    End If
                  Else
                    If (plmnid < 40464) Then
                      If (plmnid = 40458) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40459) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40460) Then
                        Return "Vodafone India"
                      End If
                      If (plmnid = 40462) Then
                        Return "cellone"
                      End If
                    Else
                      If (plmnid < 40467) Then
                        If (plmnid = 40464) Then
                          Return "cellone"
                        End If
                        If (plmnid = 40466) Then
                          Return "cellone"
                        End If
                      Else
                        If (plmnid = 40467) Then
                          Return "Reliance"
                        End If
                        If (plmnid = 40468) Then
                          Return "DOLPHIN"
                        End If
                        If (plmnid = 40469) Then
                          Return "DOLPHIN"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 40480) Then
                    If (plmnid < 40475) Then
                      If (plmnid = 40471) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40472) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40473) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40474) Then
                        Return "cellone"
                      End If
                    Else
                      If (plmnid < 40477) Then
                        If (plmnid = 40475) Then
                          Return "cellone"
                        End If
                        If (plmnid = 40476) Then
                          Return "cellone"
                        End If
                      Else
                        If (plmnid = 40477) Then
                          Return "cellone"
                        End If
                        If (plmnid = 40478) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 40479) Then
                          Return "cellone"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 40485) Then
                      If (plmnid = 40480) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40481) Then
                        Return "cellone"
                      End If
                      If (plmnid = 40483) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40484) Then
                        Return "Vodafone India"
                      End If
                    Else
                      If (plmnid < 40490) Then
                        If (plmnid = 40485) Then
                          Return "Reliance"
                        End If
                        If (plmnid = 40486) Then
                          Return "Vodafone India"
                        End If
                      Else
                        If (plmnid = 40490) Then
                          Return "AirTel"
                        End If
                        If (plmnid = 40491) Then
                          Return "AIRCEL"
                        End If
                        If (plmnid = 40492) Then
                          Return "AirTel"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If (plmnid < 41004) Then
                If (plmnid < 40517) Then
                  If (plmnid < 40507) Then
                    If (plmnid < 40503) Then
                      If (plmnid = 40493) Then
                        Return "AirTel"
                      End If
                      If (plmnid = 40495) Then
                        Return "AirTel"
                      End If
                      If (plmnid = 40496) Then
                        Return "AirTel"
                      End If
                      If (plmnid = 40501) Then
                        Return "Reliance"
                      End If
                    Else
                      If (plmnid = 40503) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40504) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40505) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40506) Then
                        Return "Reliance"
                      End If
                    End If
                  Else
                    If (plmnid < 40511) Then
                      If (plmnid = 40507) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40508) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40509) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40510) Then
                        Return "Reliance"
                      End If
                    Else
                      If (plmnid < 40513) Then
                        If (plmnid = 40511) Then
                          Return "Reliance"
                        End If
                        If (plmnid = 40512) Then
                          Return "Reliance"
                        End If
                      Else
                        If (plmnid = 40513) Then
                          Return "Reliance"
                        End If
                        If (plmnid = 40514) Then
                          Return "Reliance"
                        End If
                        If (plmnid = 40515) Then
                          Return "Reliance"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 40553) Then
                    If (plmnid < 40521) Then
                      If (plmnid = 40517) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40518) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40519) Then
                        Return "Reliance"
                      End If
                      If (plmnid = 40520) Then
                        Return "Reliance"
                      End If
                    Else
                      If (plmnid < 40523) Then
                        If (plmnid = 40521) Then
                          Return "Reliance"
                        End If
                        If (plmnid = 40522) Then
                          Return "Reliance"
                        End If
                      Else
                        If (plmnid = 40523) Then
                          Return "Reliance"
                        End If
                        If (plmnid = 40551) Then
                          Return "AirTel"
                        End If
                        If (plmnid = 40552) Then
                          Return "AirTel"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 40566) Then
                      If (plmnid = 40553) Then
                        Return "AirTel"
                      End If
                      If (plmnid = 40554) Then
                        Return "AirTel"
                      End If
                      If (plmnid = 40555) Then
                        Return "Airtel IN"
                      End If
                      If (plmnid = 40556) Then
                        Return "AirTel"
                      End If
                    Else
                      If (plmnid < 41001) Then
                        If (plmnid = 40566) Then
                          Return "Vodafone India"
                        End If
                        If (plmnid = 40570) Then
                          Return "IDEA"
                        End If
                      Else
                        If (plmnid = 41001) Then
                          Return "Jazz"
                        End If
                        If (plmnid = 41002) Then
                          Return "3G EVO / CharJi 4G"
                        End If
                        If (plmnid = 41003) Then
                          Return "Ufone"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 41406) Then
                  If (plmnid < 41250) Then
                    If (plmnid < 41008) Then
                      If (plmnid = 41004) Then
                        Return "Zong"
                      End If
                      If (plmnid = 41005) Then
                        Return "SCO Mobile"
                      End If
                      If (plmnid = 41006) Then
                        Return "Telenor PK"
                      End If
                      If (plmnid = 41007) Then
                        Return "Jazz"
                      End If
                    Else
                      If (plmnid = 41008) Then
                        Return "SCO Mobile"
                      End If
                      If (plmnid = 41201) Then
                        Return "AWCC"
                      End If
                      If (plmnid = 41220) Then
                        Return "Roshan"
                      End If
                      If (plmnid = 41240) Then
                        Return "MTN AF"
                      End If
                    End If
                  Else
                    If (plmnid < 41302) Then
                      If (plmnid = 41250) Then
                        Return "Etisalat AF"
                      End If
                      If (plmnid = 41280) Then
                        Return "Salaam"
                      End If
                      If (plmnid = 41288) Then
                        Return "Salaam"
                      End If
                      If (plmnid = 41301) Then
                        Return "Mobitel LK"
                      End If
                    Else
                      If (plmnid < 41309) Then
                        If (plmnid = 41302) Then
                          Return "Dialog"
                        End If
                        If (plmnid = 41305) Then
                          Return "Airtel LK"
                        End If
                      Else
                        If (plmnid = 41309) Then
                          Return "Hutch"
                        End If
                        If (plmnid = 41401) Then
                          Return "MPT"
                        End If
                        If (plmnid = 41405) Then
                          Return "Ooredoo MM"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 41800) Then
                    If (plmnid < 41601) Then
                      If (plmnid = 41406) Then
                        Return "Telenor MM"
                      End If
                      If (plmnid = 41409) Then
                        Return "Mytel"
                      End If
                      If (plmnid = 41501) Then
                        Return "Alfa"
                      End If
                      If (plmnid = 41503) Then
                        Return "Touch"
                      End If
                    Else
                      If (plmnid < 41677) Then
                        If (plmnid = 41601) Then
                          Return "zain JO"
                        End If
                        If (plmnid = 41603) Then
                          Return "Umniah"
                        End If
                      Else
                        If (plmnid = 41677) Then
                          Return "Orange JO"
                        End If
                        If (plmnid = 41701) Then
                          Return "Syriatel"
                        End If
                        If (plmnid = 41702) Then
                          Return "MTN SY"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 41830) Then
                      If (plmnid = 41800) Then
                        Return "Asia Cell"
                      End If
                      If (plmnid = 41805) Then
                        Return "Asia Cell"
                      End If
                      If (plmnid = 41808) Then
                        Return "SanaTel"
                      End If
                      If (plmnid = 41820) Then
                        Return "Zain IQ"
                      End If
                    Else
                      If (plmnid < 41902) Then
                        If (plmnid = 41830) Then
                          Return "Zain IQ"
                        End If
                        If (plmnid = 41840) Then
                          Return "Korek"
                        End If
                      Else
                        If (plmnid = 41902) Then
                          Return "zain KW"
                        End If
                        If (plmnid = 41903) Then
                          Return "K.S.C Ooredoo"
                        End If
                        If (plmnid = 41904) Then
                          Return "STC KW"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          Else
            If (plmnid < 44054) Then
              If (plmnid < 43211) Then
                If (plmnid < 42506) Then
                  If (plmnid < 42203) Then
                    If (plmnid < 42101) Then
                      If (plmnid = 42001) Then
                        Return "Al Jawal (STC )"
                      End If
                      If (plmnid = 42003) Then
                        Return "Mobily"
                      End If
                      If (plmnid = 42004) Then
                        Return "Zain SA"
                      End If
                      If (plmnid = 42021) Then
                        Return "RGSM"
                      End If
                    Else
                      If (plmnid = 42101) Then
                        Return "SabaFon"
                      End If
                      If (plmnid = 42102) Then
                        Return "MTN YE"
                      End If
                      If (plmnid = 42104) Then
                        Return "Y"
                      End If
                      If (plmnid = 42202) Then
                        Return "Omantel"
                      End If
                    End If
                  Else
                    If (plmnid < 42502) Then
                      If (plmnid = 42203) Then
                        Return "ooredoo OM"
                      End If
                      If (plmnid = 42402) Then
                        Return "Etisalat AE"
                      End If
                      If (plmnid = 42403) Then
                        Return "du"
                      End If
                      If (plmnid = 42501) Then
                        Return "Partner"
                      End If
                    Else
                      If (plmnid < 42505) Then
                        If (plmnid = 42502) Then
                          Return "Cellcom IL"
                        End If
                        If (plmnid = 42503) Then
                          Return "Pelephone"
                        End If
                      Else
                        If (plmnid = 42505) Then
                          Return "Jawwal IL"
                        End If
                        If (plmnid = 42505) Then
                          Return "Jawwal PS"
                        End If
                        If (plmnid = 42506) Then
                          Return "Wataniya Mobile"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 42701) Then
                    If (plmnid < 42510) Then
                      If (plmnid = 42506) Then
                        Return "Wataniya"
                      End If
                      If (plmnid = 42507) Then
                        Return "Hot Mobile"
                      End If
                      If (plmnid = 42508) Then
                        Return "Golan Telecom"
                      End If
                      If (plmnid = 42509) Then
                        Return "We4G"
                      End If
                    Else
                      If (plmnid < 42602) Then
                        If (plmnid = 42510) Then
                          Return "Partner"
                        End If
                        If (plmnid = 42601) Then
                          Return "Batelco"
                        End If
                      Else
                        If (plmnid = 42602) Then
                          Return "zain BH"
                        End If
                        If (plmnid = 42604) Then
                          Return "STC BH"
                        End If
                        If (plmnid = 42605) Then
                          Return "Batelco"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 42898) Then
                      If (plmnid = 42701) Then
                        Return "ooredoo QA"
                      End If
                      If (plmnid = 42702) Then
                        Return "Vodafone QA"
                      End If
                      If (plmnid = 42888) Then
                        Return "Unitel MN"
                      End If
                      If (plmnid = 42891) Then
                        Return "Skytel"
                      End If
                    Else
                      If (plmnid < 42901) Then
                        If (plmnid = 42898) Then
                          Return "G-Mobile"
                        End If
                        If (plmnid = 42899) Then
                          Return "Mobicom"
                        End If
                      Else
                        If (plmnid = 42901) Then
                          Return "Namaste / NT Mobile / Sky Phone"
                        End If
                        If (plmnid = 42902) Then
                          Return "Ncell"
                        End If
                        If (plmnid = 42904) Then
                          Return "SmartCell"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 43408) Then
                  If (plmnid < 43270) Then
                    If (plmnid < 43220) Then
                      If (plmnid = 43211) Then
                        Return "IR-TCI (Hamrah-e-Avval)"
                      End If
                      If (plmnid = 43212) Then
                        Return "Avacell(HiWEB)"
                      End If
                      If (plmnid = 43214) Then
                        Return "TKC/KFZO"
                      End If
                      If (plmnid = 43219) Then
                        Return "Espadan (JV-PJS)"
                      End If
                    Else
                      If (plmnid = 43220) Then
                        Return "RighTel"
                      End If
                      If (plmnid = 43221) Then
                        Return "RighTel"
                      End If
                      If (plmnid = 43232) Then
                        Return "Taliya"
                      End If
                      If (plmnid = 43235) Then
                        Return "MTN Irancell"
                      End If
                    End If
                  Else
                    If (plmnid < 43293) Then
                      If (plmnid = 43270) Then
                        Return "MTCE"
                      End If
                      If (plmnid = 43271) Then
                        Return "KOOHE NOOR"
                      End If
                      If (plmnid = 43290) Then
                        Return "Iraphone"
                      End If
                      If (plmnid = 43293) Then
                        Return "Iraphone"
                      End If
                    Else
                      If (plmnid < 43404) Then
                        If (plmnid = 43293) Then
                          Return "Farzanegan Pars"
                        End If
                        If (plmnid = 43299) Then
                          Return "TCI (GSM WLL)"
                        End If
                      Else
                        If (plmnid = 43404) Then
                          Return "Beeline UZ"
                        End If
                        If (plmnid = 43405) Then
                          Return "Ucell"
                        End If
                        If (plmnid = 43407) Then
                          Return "Mobiuz"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 43802) Then
                    If (plmnid < 43604) Then
                      If (plmnid = 43408) Then
                        Return "UzMobile"
                      End If
                      If (plmnid = 43601) Then
                        Return "Tcell"
                      End If
                      If (plmnid = 43602) Then
                        Return "Tcell"
                      End If
                      If (plmnid = 43603) Then
                        Return "MegaFon TJ"
                      End If
                    Else
                      If (plmnid < 43701) Then
                        If (plmnid = 43604) Then
                          Return "Babilon-M"
                        End If
                        If (plmnid = 43605) Then
                          Return "ZET-Mobile"
                        End If
                      Else
                        If (plmnid = 43701) Then
                          Return "Beeline KG"
                        End If
                        If (plmnid = 43705) Then
                          Return "MegaCom"
                        End If
                        If (plmnid = 43709) Then
                          Return "O!"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 44021) Then
                      If (plmnid = 43802) Then
                        Return "TM-Cell"
                      End If
                      If (plmnid = 44010) Then
                        Return "NTT docomo"
                      End If
                      If (plmnid = 44011) Then
                        Return "Rakuten Mobile"
                      End If
                      If (plmnid = 44020) Then
                        Return "SoftBank"
                      End If
                    Else
                      If (plmnid < 44051) Then
                        If (plmnid = 44021) Then
                          Return "SoftBank"
                        End If
                        If (plmnid = 44050) Then
                          Return "au"
                        End If
                      Else
                        If (plmnid = 44051) Then
                          Return "au"
                        End If
                        If (plmnid = 44052) Then
                          Return "au"
                        End If
                        If (plmnid = 44053) Then
                          Return "au"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If (plmnid < 45606) Then
                If (plmnid < 45205) Then
                  If (plmnid < 44101) Then
                    If (plmnid < 44073) Then
                      If (plmnid = 44054) Then
                        Return "au"
                      End If
                      If (plmnid = 44070) Then
                        Return "au"
                      End If
                      If (plmnid = 44071) Then
                        Return "au"
                      End If
                      If (plmnid = 44072) Then
                        Return "au"
                      End If
                    Else
                      If (plmnid = 44073) Then
                        Return "au"
                      End If
                      If (plmnid = 44074) Then
                        Return "au"
                      End If
                      If (plmnid = 44075) Then
                        Return "au"
                      End If
                      If (plmnid = 44076) Then
                        Return "au"
                      End If
                    End If
                  Else
                    If (plmnid < 45008) Then
                      If (plmnid = 44101) Then
                        Return "SoftBank"
                      End If
                      If (plmnid = 45004) Then
                        Return "KT"
                      End If
                      If (plmnid = 45005) Then
                        Return "SKTelecom"
                      End If
                      If (plmnid = 45006) Then
                        Return "LG U+"
                      End If
                    Else
                      If (plmnid < 45201) Then
                        If (plmnid = 45008) Then
                          Return "olleh"
                        End If
                        If (plmnid = 45012) Then
                          Return "SKTelecom"
                        End If
                      Else
                        If (plmnid = 45201) Then
                          Return "MobiFone"
                        End If
                        If (plmnid = 45202) Then
                          Return "Vinaphone"
                        End If
                        If (plmnid = 45204) Then
                          Return "Viettel Mobile"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 45500) Then
                    If (plmnid < 45404) Then
                      If (plmnid = 45205) Then
                        Return "Vietnamobile"
                      End If
                      If (plmnid = 45207) Then
                        Return "Gmobile"
                      End If
                      If (plmnid = 45400) Then
                        Return "1O1O / One2Free / New World Mobility / SUNMobile"
                      End If
                      If (plmnid = 45403) Then
                        Return "3 HK"
                      End If
                    Else
                      If (plmnid < 45412) Then
                        If (plmnid = 45404) Then
                          Return "3 (2G)"
                        End If
                        If (plmnid = 45406) Then
                          Return "SmarTone HK"
                        End If
                      Else
                        If (plmnid = 45412) Then
                          Return "CMCC HK"
                        End If
                        If (plmnid = 45416) Then
                          Return "PCCW Mobile (2G)"
                        End If
                        If (plmnid = 45420) Then
                          Return "PCCW Mobile (4G)"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 45601) Then
                      If (plmnid = 45500) Then
                        Return "SmarTone MO"
                      End If
                      If (plmnid = 45501) Then
                        Return "CTM"
                      End If
                      If (plmnid = 45505) Then
                        Return "3 MO"
                      End If
                      If (plmnid = 45507) Then
                        Return "China Telecom MO"
                      End If
                    Else
                      If (plmnid < 45603) Then
                        If (plmnid = 45601) Then
                          Return "Cellcard"
                        End If
                        If (plmnid = 45602) Then
                          Return "Smart KH"
                        End If
                      Else
                        If (plmnid = 45603) Then
                          Return "qb"
                        End If
                        If (plmnid = 45604) Then
                          Return "qb"
                        End If
                        If (plmnid = 45605) Then
                          Return "Smart KH"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 46697) Then
                  If (plmnid < 45708) Then
                    If (plmnid < 45618) Then
                      If (plmnid = 45606) Then
                        Return "Smart KH"
                      End If
                      If (plmnid = 45608) Then
                        Return "Metfone"
                      End If
                      If (plmnid = 45609) Then
                        Return "Metfone"
                      End If
                      If (plmnid = 45611) Then
                        Return "SEATEL"
                      End If
                    Else
                      If (plmnid = 45618) Then
                        Return "Cellcard"
                      End If
                      If (plmnid = 45701) Then
                        Return "LaoTel"
                      End If
                      If (plmnid = 45702) Then
                        Return "ETL"
                      End If
                      If (plmnid = 45703) Then
                        Return "Unitel LA"
                      End If
                    End If
                  Else
                    If (plmnid < 46020) Then
                      If (plmnid = 45708) Then
                        Return "Beeline LA"
                      End If
                      If (plmnid = 46000) Then
                        Return "China Mobile"
                      End If
                      If (plmnid = 46001) Then
                        Return "China Unicom"
                      End If
                      If (plmnid = 46003) Then
                        Return "China Telecom CN"
                      End If
                    Else
                      If (plmnid < 46605) Then
                        If (plmnid = 46020) Then
                          Return "China Tietong"
                        End If
                        If (plmnid = 46601) Then
                          Return "FarEasTone"
                        End If
                      Else
                        If (plmnid = 46605) Then
                          Return "APTG"
                        End If
                        If (plmnid = 46689) Then
                          Return "T Star"
                        End If
                        If (plmnid = 46692) Then
                          Return "Chunghwa"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 50211) Then
                    If (plmnid < 47004) Then
                      If (plmnid = 46697) Then
                        Return "Taiwan Mobile"
                      End If
                      If (plmnid = 47001) Then
                        Return "Grameenphone"
                      End If
                      If (plmnid = 47002) Then
                        Return "Robi"
                      End If
                      If (plmnid = 47003) Then
                        Return "Banglalink"
                      End If
                    Else
                      If (plmnid < 47009) Then
                        If (plmnid = 47004) Then
                          Return "TeleTalk"
                        End If
                        If (plmnid = 47007) Then
                          Return "Airtel BD"
                        End If
                      Else
                        If (plmnid = 47009) Then
                          Return "ollo"
                        End If
                        If (plmnid = 47201) Then
                          Return "Dhiraagu"
                        End If
                        If (plmnid = 47202) Then
                          Return "Ooredoo MV"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 50217) Then
                      If (plmnid = 50211) Then
                        Return "TM Homeline"
                      End If
                      If (plmnid = 50212) Then
                        Return "Maxis"
                      End If
                      If (plmnid = 50213) Then
                        Return "Celcom"
                      End If
                      If (plmnid = 50216) Then
                        Return "DiGi"
                      End If
                    Else
                      If (plmnid < 50219) Then
                        If (plmnid = 50217) Then
                          Return "Maxis"
                        End If
                        If (plmnid = 50218) Then
                          Return "U Mobile"
                        End If
                      Else
                        If (plmnid = 50219) Then
                          Return "Celcom"
                        End If
                        If (plmnid = 50501) Then
                          Return "Telstra"
                        End If
                        If (plmnid = 50502) Then
                          Return "Optus"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      Else
        If (plmnid < 72402) Then
          If (plmnid < 62303) Then
            If (plmnid < 60501) Then
              If (plmnid < 53901) Then
                If (plmnid < 52001) Then
                  If (plmnid < 51011) Then
                    If (plmnid < 50516) Then
                      If (plmnid = 50503) Then
                        Return "Vodafone AU"
                      End If
                      If (plmnid = 50510) Then
                        Return "Norfolk Is."
                      End If
                      If (plmnid = 50510) Then
                        Return "Norfolk Telecom"
                      End If
                      If (plmnid = 50513) Then
                        Return "RailCorp"
                      End If
                    Else
                      If (plmnid = 50516) Then
                        Return "VicTrack"
                      End If
                      If (plmnid = 51001) Then
                        Return "Indosat Ooredoo"
                      End If
                      If (plmnid = 51009) Then
                        Return "Smartfren"
                      End If
                      If (plmnid = 51010) Then
                        Return "Telkomsel"
                      End If
                    End If
                  Else
                    If (plmnid < 51402) Then
                      If (plmnid = 51011) Then
                        Return "XL"
                      End If
                      If (plmnid = 51028) Then
                        Return "Fren/Hepi"
                      End If
                      If (plmnid = 51089) Then
                        Return "3 ID"
                      End If
                      If (plmnid = 51401) Then
                        Return "Telkomcel"
                      End If
                    Else
                      If (plmnid < 51502) Then
                        If (plmnid = 51402) Then
                          Return "TT"
                        End If
                        If (plmnid = 51403) Then
                          Return "Telemor"
                        End If
                      Else
                        If (plmnid = 51502) Then
                          Return "Globe"
                        End If
                        If (plmnid = 51503) Then
                          Return "SMART PH"
                        End If
                        If (plmnid = 51505) Then
                          Return "Sun Cellular"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 52505) Then
                    If (plmnid < 52018) Then
                      If (plmnid = 52001) Then
                        Return "AIS"
                      End If
                      If (plmnid = 52003) Then
                        Return "AIS"
                      End If
                      If (plmnid = 52004) Then
                        Return "TrueMove H"
                      End If
                      If (plmnid = 52005) Then
                        Return "dtac"
                      End If
                    Else
                      If (plmnid = 52018) Then
                        Return "dtac"
                      End If
                      If (plmnid = 52099) Then
                        Return "TrueMove"
                      End If
                      If (plmnid = 52501) Then
                        Return "SingTel"
                      End If
                      If (plmnid = 52503) Then
                        Return "M1"
                      End If
                    End If
                  Else
                    If (plmnid < 53024) Then
                      If (plmnid = 52505) Then
                        Return "StarHub"
                      End If
                      If (plmnid = 52811) Then
                        Return "DST"
                      End If
                      If (plmnid = 53001) Then
                        Return "Vodafone NZ"
                      End If
                      If (plmnid = 53005) Then
                        Return "Spark"
                      End If
                    Else
                      If (plmnid < 53701) Then
                        If (plmnid = 53024) Then
                          Return "2degrees"
                        End If
                        If (plmnid = 53602) Then
                          Return "Digicel NR"
                        End If
                      Else
                        If (plmnid = 53701) Then
                          Return "bmobile PG"
                        End If
                        If (plmnid = 53702) Then
                          Return "citifon"
                        End If
                        If (plmnid = 53703) Then
                          Return "Digicel PG"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 54801) Then
                  If (plmnid < 54202) Then
                    If (plmnid < 54100) Then
                      If (plmnid = 53901) Then
                        Return "U-Call"
                      End If
                      If (plmnid = 53988) Then
                        Return "Digicel TO"
                      End If
                      If (plmnid = 54001) Then
                        Return "BREEZE"
                      End If
                      If (plmnid = 54002) Then
                        Return "BeMobile"
                      End If
                    Else
                      If (plmnid = 54100) Then
                        Return "AIL"
                      End If
                      If (plmnid = 54101) Then
                        Return "SMILE"
                      End If
                      If (plmnid = 54105) Then
                        Return "Digicel VU"
                      End If
                      If (plmnid = 54201) Then
                        Return "Vodafone FJ"
                      End If
                    End If
                  Else
                    If (plmnid < 54509) Then
                      If (plmnid = 54202) Then
                        Return "Digicel FJ"
                      End If
                      If (plmnid = 54203) Then
                        Return "TFL"
                      End If
                      If (plmnid = 54411) Then
                        Return "Bluesky AS"
                      End If
                      If (plmnid = 54501) Then
                        Return "Kiribati - ATH"
                      End If
                    Else
                      If (plmnid < 54705) Then
                        If (plmnid = 54509) Then
                          Return "Kiribati - Frigate Net"
                        End If
                        If (plmnid = 54601) Then
                          Return "Mobilis NC"
                        End If
                      Else
                        If (plmnid = 54705) Then
                          Return "Ora"
                        End If
                        If (plmnid = 54715) Then
                          Return "Vodafone PF"
                        End If
                        If (plmnid = 54720) Then
                          Return "Vini"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 60203) Then
                    If (plmnid < 55202) Then
                      If (plmnid = 54801) Then
                        Return "Bluesky CK"
                      End If
                      If (plmnid = 54901) Then
                        Return "Digicel WS"
                      End If
                      If (plmnid = 54927) Then
                        Return "Bluesky WS"
                      End If
                      If (plmnid = 55201) Then
                        Return "PNCC"
                      End If
                    Else
                      If (plmnid < 55501) Then
                        If (plmnid = 55202) Then
                          Return "PT Waves"
                        End If
                        If (plmnid = 55301) Then
                          Return "TTC"
                        End If
                      Else
                        If (plmnid = 55501) Then
                          Return "Telecom Niue"
                        End If
                        If (plmnid = 60201) Then
                          Return "Orange EG"
                        End If
                        If (plmnid = 60202) Then
                          Return "Vodafone EG"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 60303) Then
                      If (plmnid = 60203) Then
                        Return "Etisalat EG"
                      End If
                      If (plmnid = 60204) Then
                        Return "WE"
                      End If
                      If (plmnid = 60301) Then
                        Return "Mobilis DZ"
                      End If
                      If (plmnid = 60302) Then
                        Return "Djezzy"
                      End If
                    Else
                      If (plmnid < 60401) Then
                        If (plmnid = 60303) Then
                          Return "Ooredoo DZ"
                        End If
                        If (plmnid = 60400) Then
                          Return "Orange Morocco"
                        End If
                      Else
                        If (plmnid = 60401) Then
                          Return "IAM"
                        End If
                        If (plmnid = 60402) Then
                          Return "INWI"
                        End If
                        If (plmnid = 60405) Then
                          Return "INWI"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If (plmnid < 61403) Then
                If (plmnid < 61002) Then
                  If (plmnid < 60703) Then
                    If (plmnid < 60601) Then
                      If (plmnid = 60501) Then
                        Return "Orange TN"
                      End If
                      If (plmnid = 60502) Then
                        Return "Tunicell"
                      End If
                      If (plmnid = 60503) Then
                        Return "OOREDOO TN"
                      End If
                      If (plmnid = 60600) Then
                        Return "Libyana"
                      End If
                    Else
                      If (plmnid = 60601) Then
                        Return "Madar"
                      End If
                      If (plmnid = 60603) Then
                        Return "Libya Phone"
                      End If
                      If (plmnid = 60701) Then
                        Return "Gamcel"
                      End If
                      If (plmnid = 60702) Then
                        Return "Africell GM"
                      End If
                    End If
                  Else
                    If (plmnid < 60803) Then
                      If (plmnid = 60703) Then
                        Return "Comium"
                      End If
                      If (plmnid = 60704) Then
                        Return "QCell"
                      End If
                      If (plmnid = 60801) Then
                        Return "Orange SN"
                      End If
                      If (plmnid = 60802) Then
                        Return "Free SN"
                      End If
                    Else
                      If (plmnid < 60902) Then
                        If (plmnid = 60803) Then
                          Return "Expresso"
                        End If
                        If (plmnid = 60901) Then
                          Return "Mattel"
                        End If
                      Else
                        If (plmnid = 60902) Then
                          Return "Chinguitel"
                        End If
                        If (plmnid = 60910) Then
                          Return "Mauritel"
                        End If
                        If (plmnid = 61001) Then
                          Return "Malitel"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 61204) Then
                    If (plmnid < 61103) Then
                      If (plmnid = 61002) Then
                        Return "Orange ML"
                      End If
                      If (plmnid = 61003) Then
                        Return "Telecel ML"
                      End If
                      If (plmnid = 61101) Then
                        Return "Orange GN"
                      End If
                      If (plmnid = 61102) Then
                        Return "Sotelgui"
                      End If
                    Else
                      If (plmnid < 61105) Then
                        If (plmnid = 61103) Then
                          Return "Telecel Guinee"
                        End If
                        If (plmnid = 61104) Then
                          Return "MTN GN"
                        End If
                      Else
                        If (plmnid = 61105) Then
                          Return "Cellcom GN"
                        End If
                        If (plmnid = 61202) Then
                          Return "Moov CI"
                        End If
                        If (plmnid = 61203) Then
                          Return "Orange CI"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 61301) Then
                      If (plmnid = 61204) Then
                        Return "KoZ"
                      End If
                      If (plmnid = 61205) Then
                        Return "MTN CI"
                      End If
                      If (plmnid = 61206) Then
                        Return "GreenN"
                      End If
                      If (plmnid = 61207) Then
                        Return "café"
                      End If
                    Else
                      If (plmnid < 61303) Then
                        If (plmnid = 61301) Then
                          Return "Telmob"
                        End If
                        If (plmnid = 61302) Then
                          Return "Orange BF"
                        End If
                      Else
                        If (plmnid = 61303) Then
                          Return "Telecel Faso"
                        End If
                        If (plmnid = 61401) Then
                          Return "SahelCom"
                        End If
                        If (plmnid = 61402) Then
                          Return "Airtel NE"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 61909) Then
                  If (plmnid < 61701) Then
                    If (plmnid < 61601) Then
                      If (plmnid = 61403) Then
                        Return "Moov NE"
                      End If
                      If (plmnid = 61404) Then
                        Return "Orange NE"
                      End If
                      If (plmnid = 61501) Then
                        Return "Togo Cell"
                      End If
                      If (plmnid = 61503) Then
                        Return "Moov TG"
                      End If
                    Else
                      If (plmnid = 61601) Then
                        Return "Libercom"
                      End If
                      If (plmnid = 61602) Then
                        Return "Moov BJ"
                      End If
                      If (plmnid = 61603) Then
                        Return "MTN BJ"
                      End If
                      If (plmnid = 61604) Then
                        Return "BBCOM"
                      End If
                    End If
                  Else
                    If (plmnid < 61804) Then
                      If (plmnid = 61701) Then
                        Return "my.t"
                      End If
                      If (plmnid = 61703) Then
                        Return "CHILI"
                      End If
                      If (plmnid = 61710) Then
                        Return "Emtel"
                      End If
                      If (plmnid = 61801) Then
                        Return "Lonestar Cell MTN"
                      End If
                    Else
                      If (plmnid < 61901) Then
                        If (plmnid = 61804) Then
                          Return "Novafone"
                        End If
                        If (plmnid = 61807) Then
                          Return "Orange LBR"
                        End If
                      Else
                        If (plmnid = 61901) Then
                          Return "Orange SL"
                        End If
                        If (plmnid = 61903) Then
                          Return "Africell SL"
                        End If
                        If (plmnid = 61905) Then
                          Return "Africell SL"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 62130) Then
                    If (plmnid < 62006) Then
                      If (plmnid = 61909) Then
                        Return "Smart Mobile SL"
                      End If
                      If (plmnid = 62001) Then
                        Return "MTN GH"
                      End If
                      If (plmnid = 62002) Then
                        Return "Vodafone GH"
                      End If
                      If (plmnid = 62003) Then
                        Return "AirtelTigo"
                      End If
                    Else
                      If (plmnid < 62120) Then
                        If (plmnid = 62006) Then
                          Return "AirtelTigo"
                        End If
                        If (plmnid = 62007) Then
                          Return "Globacom"
                        End If
                      Else
                        If (plmnid = 62120) Then
                          Return "Airtel NG"
                        End If
                        If (plmnid = 62122) Then
                          Return "InterC"
                        End If
                        If (plmnid = 62127) Then
                          Return "Smile NG"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 62201) Then
                      If (plmnid = 62130) Then
                        Return "MTN NG"
                      End If
                      If (plmnid = 62140) Then
                        Return "Ntel"
                      End If
                      If (plmnid = 62150) Then
                        Return "Glo"
                      End If
                      If (plmnid = 62160) Then
                        Return "9mobile"
                      End If
                    Else
                      If (plmnid < 62207) Then
                        If (plmnid = 62201) Then
                          Return "Airtel TD"
                        End If
                        If (plmnid = 62203) Then
                          Return "Tigo TD"
                        End If
                      Else
                        If (plmnid = 62207) Then
                          Return "Salam"
                        End If
                        If (plmnid = 62301) Then
                          Return "Moov CF"
                        End If
                        If (plmnid = 62302) Then
                          Return "TC"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          Else
            If (plmnid < 64201) Then
              If (plmnid < 63409) Then
                If (plmnid < 62910) Then
                  If (plmnid < 62601) Then
                    If (plmnid < 62403) Then
                      If (plmnid = 62303) Then
                        Return "Orange CF"
                      End If
                      If (plmnid = 62304) Then
                        Return "Azur"
                      End If
                      If (plmnid = 62401) Then
                        Return "MTN Cameroon"
                      End If
                      If (plmnid = 62402) Then
                        Return "Orange CM"
                      End If
                    Else
                      If (plmnid = 62403) Then
                        Return "Camtel"
                      End If
                      If (plmnid = 62404) Then
                        Return "Nexttel"
                      End If
                      If (plmnid = 62501) Then
                        Return "CVMOVEL"
                      End If
                      If (plmnid = 62502) Then
                        Return "T+"
                      End If
                    End If
                  Else
                    If (plmnid < 62801) Then
                      If (plmnid = 62601) Then
                        Return "CSTmovel"
                      End If
                      If (plmnid = 62602) Then
                        Return "Unitel STP"
                      End If
                      If (plmnid = 62701) Then
                        Return "Orange GQ"
                      End If
                      If (plmnid = 62703) Then
                        Return "Muni"
                      End If
                    Else
                      If (plmnid < 62803) Then
                        If (plmnid = 62801) Then
                          Return "Libertis"
                        End If
                        If (plmnid = 62802) Then
                          Return "Moov GA"
                        End If
                      Else
                        If (plmnid = 62803) Then
                          Return "Airtel GA"
                        End If
                        If (plmnid = 62901) Then
                          Return "Airtel CG"
                        End If
                        If (plmnid = 62907) Then
                          Return "Airtel CG"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 63201) Then
                    If (plmnid < 63086) Then
                      If (plmnid = 62910) Then
                        Return "Libertis Telecom"
                      End If
                      If (plmnid = 63001) Then
                        Return "Vodacom CD"
                      End If
                      If (plmnid = 63002) Then
                        Return "Airtel CD"
                      End If
                      If (plmnid = 63005) Then
                        Return "Supercell"
                      End If
                    Else
                      If (plmnid < 63090) Then
                        If (plmnid = 63086) Then
                          Return "Orange RDC"
                        End If
                        If (plmnid = 63089) Then
                          Return "Orange RDC"
                        End If
                      Else
                        If (plmnid = 63090) Then
                          Return "Africell CD"
                        End If
                        If (plmnid = 63102) Then
                          Return "UNITEL AO"
                        End If
                        If (plmnid = 63104) Then
                          Return "MOVICEL"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 63301) Then
                      If (plmnid = 63201) Then
                        Return "Guinetel"
                      End If
                      If (plmnid = 63202) Then
                        Return "MTN Areeba"
                      End If
                      If (plmnid = 63203) Then
                        Return "Orange GW"
                      End If
                      If (plmnid = 63207) Then
                        Return "Guinetel"
                      End If
                    Else
                      If (plmnid < 63401) Then
                        If (plmnid = 63301) Then
                          Return "Cable & Wireless SC"
                        End If
                        If (plmnid = 63310) Then
                          Return "Airtel SC"
                        End If
                      Else
                        If (plmnid = 63401) Then
                          Return "Zain SD"
                        End If
                        If (plmnid = 63402) Then
                          Return "MTN SD"
                        End If
                        If (plmnid = 63407) Then
                          Return "Sudani One"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 63902) Then
                  If (plmnid < 63720) Then
                    If (plmnid < 63601) Then
                      If (plmnid = 63409) Then
                        Return "khartoum INC"
                      End If
                      If (plmnid = 63510) Then
                        Return "MTN RW"
                      End If
                      If (plmnid = 63513) Then
                        Return "Airtel RW"
                      End If
                      If (plmnid = 63517) Then
                        Return "Olleh"
                      End If
                    Else
                      If (plmnid = 63601) Then
                        Return "MTN ET"
                      End If
                      If (plmnid = 63701) Then
                        Return "Telesom"
                      End If
                      If (plmnid = 63704) Then
                        Return "Somafone"
                      End If
                      If (plmnid = 63710) Then
                        Return "Nationlink"
                      End If
                    End If
                  Else
                    If (plmnid < 63760) Then
                      If (plmnid = 63720) Then
                        Return "SOMNET"
                      End If
                      If (plmnid = 63730) Then
                        Return "Golis"
                      End If
                      If (plmnid = 63750) Then
                        Return "Hormuud"
                      End If
                      If (plmnid = 63757) Then
                        Return "UNITEL SO"
                      End If
                    Else
                      If (plmnid < 63771) Then
                        If (plmnid = 63760) Then
                          Return "Nationlink"
                        End If
                        If (plmnid = 63767) Then
                          Return "Horntel Group"
                        End If
                      Else
                        If (plmnid = 63771) Then
                          Return "Somtel"
                        End If
                        If (plmnid = 63782) Then
                          Return "Telcom"
                        End If
                        If (plmnid = 63801) Then
                          Return "Evatis"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 64009) Then
                    If (plmnid < 64002) Then
                      If (plmnid = 63902) Then
                        Return "Safaricom"
                      End If
                      If (plmnid = 63903) Then
                        Return "Airtel KE"
                      End If
                      If (plmnid = 63907) Then
                        Return "Telkom KE"
                      End If
                      If (plmnid = 63910) Then
                        Return "Faiba 4G"
                      End If
                    Else
                      If (plmnid < 64004) Then
                        If (plmnid = 64002) Then
                          Return "tiGO"
                        End If
                        If (plmnid = 64003) Then
                          Return "Zantel"
                        End If
                      Else
                        If (plmnid = 64004) Then
                          Return "Vodacom TZ"
                        End If
                        If (plmnid = 64005) Then
                          Return "Airtel TZ"
                        End If
                        If (plmnid = 64007) Then
                          Return "TTCL Mobile"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 64111) Then
                      If (plmnid = 64009) Then
                        Return "Halotel"
                      End If
                      If (plmnid = 64011) Then
                        Return "SmileCom"
                      End If
                      If (plmnid = 64101) Then
                        Return "Airtel UG"
                      End If
                      If (plmnid = 64110) Then
                        Return "MTN UG"
                      End If
                    Else
                      If (plmnid < 64118) Then
                        If (plmnid = 64111) Then
                          Return "Uganda Telecom"
                        End If
                        If (plmnid = 64114) Then
                          Return "Africell UG"
                        End If
                      Else
                        If (plmnid = 64118) Then
                          Return "Smart UG"
                        End If
                        If (plmnid = 64122) Then
                          Return "Airtel UG"
                        End If
                        If (plmnid = 64133) Then
                          Return "Smile UG"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If (plmnid < 65501) Then
                If (plmnid < 64703) Then
                  If (plmnid < 64501) Then
                    If (plmnid < 64282) Then
                      If (plmnid = 64201) Then
                        Return "econet Leo"
                      End If
                      If (plmnid = 64203) Then
                        Return "Onatel"
                      End If
                      If (plmnid = 64207) Then
                        Return "Smart Mobile BI"
                      End If
                      If (plmnid = 64208) Then
                        Return "Lumitel"
                      End If
                    Else
                      If (plmnid = 64282) Then
                        Return "econet Leo"
                      End If
                      If (plmnid = 64301) Then
                        Return "mCel"
                      End If
                      If (plmnid = 64303) Then
                        Return "Movitel"
                      End If
                      If (plmnid = 64304) Then
                        Return "Vodacom MZ"
                      End If
                    End If
                  Else
                    If (plmnid < 64602) Then
                      If (plmnid = 64501) Then
                        Return "Airtel ZM"
                      End If
                      If (plmnid = 64502) Then
                        Return "MTN ZM"
                      End If
                      If (plmnid = 64503) Then
                        Return "ZAMTEL"
                      End If
                      If (plmnid = 64601) Then
                        Return "Airtel MG"
                      End If
                    Else
                      If (plmnid < 64700) Then
                        If (plmnid = 64602) Then
                          Return "Orange MG"
                        End If
                        If (plmnid = 64604) Then
                          Return "Telma"
                        End If
                      Else
                        If (plmnid = 64700) Then
                          Return "Orange YT/RE"
                        End If
                        If (plmnid = 64701) Then
                          Return "Maoré Mobile"
                        End If
                        If (plmnid = 64702) Then
                          Return "Only"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 65010) Then
                    If (plmnid < 64804) Then
                      If (plmnid = 64703) Then
                        Return "Free YT/RE"
                      End If
                      If (plmnid = 64710) Then
                        Return "SFR Réunion"
                      End If
                      If (plmnid = 64801) Then
                        Return "Net*One"
                      End If
                      If (plmnid = 64803) Then
                        Return "Telecel ZW"
                      End If
                    Else
                      If (plmnid < 64903) Then
                        If (plmnid = 64804) Then
                          Return "Econet"
                        End If
                        If (plmnid = 64901) Then
                          Return "MTC"
                        End If
                      Else
                        If (plmnid = 64903) Then
                          Return "TN Mobile"
                        End If
                        If (plmnid = 65001) Then
                          Return "TNM"
                        End If
                        If (plmnid = 65002) Then
                          Return "Access"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 65202) Then
                      If (plmnid = 65010) Then
                        Return "Airtel MW"
                      End If
                      If (plmnid = 65101) Then
                        Return "Vodacom LS"
                      End If
                      If (plmnid = 65102) Then
                        Return "Econet Telecom"
                      End If
                      If (plmnid = 65201) Then
                        Return "Mascom"
                      End If
                    Else
                      If (plmnid < 65310) Then
                        If (plmnid = 65202) Then
                          Return "Orange BW"
                        End If
                        If (plmnid = 65204) Then
                          Return "beMobile"
                        End If
                      Else
                        If (plmnid = 65310) Then
                          Return "Swazi MTN"
                        End If
                        If (plmnid = 65401) Then
                          Return "HURI"
                        End If
                        If (plmnid = 65402) Then
                          Return "TELCO SA"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 70602) Then
                  If (plmnid < 65902) Then
                    If (plmnid < 65514) Then
                      If (plmnid = 65501) Then
                        Return "Vodacom ZA"
                      End If
                      If (plmnid = 65502) Then
                        Return "Telkom ZA"
                      End If
                      If (plmnid = 65507) Then
                        Return "Cell C"
                      End If
                      If (plmnid = 65510) Then
                        Return "MTN ZA"
                      End If
                    Else
                      If (plmnid = 65514) Then
                        Return "Neotel"
                      End If
                      If (plmnid = 65519) Then
                        Return "Rain"
                      End If
                      If (plmnid = 65701) Then
                        Return "Eritel"
                      End If
                      If (plmnid = 65801) Then
                        Return "Sure SH"
                      End If
                    End If
                  Else
                    If (plmnid < 70269) Then
                      If (plmnid = 65902) Then
                        Return "MTN SS"
                      End If
                      If (plmnid = 65903) Then
                        Return "Gemtel"
                      End If
                      If (plmnid = 65906) Then
                        Return "Zain SS"
                      End If
                      If (plmnid = 70267) Then
                        Return "DigiCell"
                      End If
                    Else
                      If (plmnid < 70402) Then
                        If (plmnid = 70269) Then
                          Return "SMART BZ"
                        End If
                        If (plmnid = 70401) Then
                          Return "Claro GT"
                        End If
                      Else
                        If (plmnid = 70402) Then
                          Return "Tigo GT"
                        End If
                        If (plmnid = 70403) Then
                          Return "movistar GT"
                        End If
                        If (plmnid = 70601) Then
                          Return "Claro SV"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 71204) Then
                    If (plmnid < 71030) Then
                      If (plmnid = 70602) Then
                        Return "Digicel SV"
                      End If
                      If (plmnid = 70603) Then
                        Return "Tigo SV"
                      End If
                      If (plmnid = 70604) Then
                        Return "movistar SV"
                      End If
                      If (plmnid = 71021) Then
                        Return "Claro NI"
                      End If
                    Else
                      If (plmnid < 71201) Then
                        If (plmnid = 71030) Then
                          Return "movistar NI"
                        End If
                        If (plmnid = 71073) Then
                          Return "Claro NI"
                        End If
                      Else
                        If (plmnid = 71201) Then
                          Return "Kölbi ICE"
                        End If
                        If (plmnid = 71202) Then
                          Return "Kölbi ICE"
                        End If
                        If (plmnid = 71203) Then
                          Return "Claro CR"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 71404) Then
                      If (plmnid = 71204) Then
                        Return "movistar CR"
                      End If
                      If (plmnid = 71401) Then
                        Return "Cable & Wireless PA"
                      End If
                      If (plmnid = 71402) Then
                        Return "movistar PA"
                      End If
                      If (plmnid = 71403) Then
                        Return "Claro PA"
                      End If
                    Else
                      If (plmnid < 71610) Then
                        If (plmnid = 71404) Then
                          Return "Digicel PA"
                        End If
                        If (plmnid = 71606) Then
                          Return "Movistar PE"
                        End If
                      Else
                        If (plmnid = 71610) Then
                          Return "Claro PE"
                        End If
                        If (plmnid = 71615) Then
                          Return "Bitel"
                        End If
                        If (plmnid = 71617) Then
                          Return "Entel PE"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        Else
          If (plmnid < 312330) Then
            If (plmnid < 310080) Then
              If (plmnid < 74602) Then
                If (plmnid < 73009) Then
                  If (plmnid < 72423) Then
                    If (plmnid < 72406) Then
                      If (plmnid = 72402) Then
                        Return "TIM BR"
                      End If
                      If (plmnid = 72403) Then
                        Return "TIM BR"
                      End If
                      If (plmnid = 72404) Then
                        Return "TIM BR"
                      End If
                      If (plmnid = 72405) Then
                        Return "Claro BR"
                      End If
                    Else
                      If (plmnid = 72406) Then
                        Return "Vivo"
                      End If
                      If (plmnid = 72410) Then
                        Return "Vivo"
                      End If
                      If (plmnid = 72411) Then
                        Return "Vivo"
                      End If
                      If (plmnid = 72415) Then
                        Return "Sercomtel"
                      End If
                    End If
                  Else
                    If (plmnid < 72434) Then
                      If (plmnid = 72423) Then
                        Return "Vivo"
                      End If
                      If (plmnid = 72431) Then
                        Return "Oi"
                      End If
                      If (plmnid = 72432) Then
                        Return "Algar Telecom"
                      End If
                      If (plmnid = 72433) Then
                        Return "Algar Telecom"
                      End If
                    Else
                      If (plmnid < 73001) Then
                        If (plmnid = 72434) Then
                          Return "Algar Telecom"
                        End If
                        If (plmnid = 72439) Then
                          Return "Nextel"
                        End If
                      Else
                        If (plmnid = 73001) Then
                          Return "entel"
                        End If
                        If (plmnid = 73002) Then
                          Return "movistar CL"
                        End If
                        If (plmnid = 73003) Then
                          Return "CLARO CL"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 73801) Then
                    If (plmnid < 73404) Then
                      If (plmnid = 73009) Then
                        Return "WOM"
                      End If
                      If (plmnid = 73010) Then
                        Return "entel"
                      End If
                      If (plmnid = 73099) Then
                        Return "Will"
                      End If
                      If (plmnid = 73402) Then
                        Return "Digitel GSM"
                      End If
                    Else
                      If (plmnid < 73601) Then
                        If (plmnid = 73404) Then
                          Return "movistar VE"
                        End If
                        If (plmnid = 73406) Then
                          Return "Movilnet"
                        End If
                      Else
                        If (plmnid = 73601) Then
                          Return "Viva BO"
                        End If
                        If (plmnid = 73602) Then
                          Return "Entel BO"
                        End If
                        If (plmnid = 73603) Then
                          Return "Tigo BO"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 74401) Then
                      If (plmnid = 73801) Then
                        Return "Digicel GY"
                      End If
                      If (plmnid = 74000) Then
                        Return "Movistar EC"
                      End If
                      If (plmnid = 74001) Then
                        Return "Claro EC"
                      End If
                      If (plmnid = 74002) Then
                        Return "CNT Mobile"
                      End If
                    Else
                      If (plmnid < 74404) Then
                        If (plmnid = 74401) Then
                          Return "VOX"
                        End If
                        If (plmnid = 74402) Then
                          Return "Claro PY"
                        End If
                      Else
                        If (plmnid = 74404) Then
                          Return "Tigo PY"
                        End If
                        If (plmnid = 74405) Then
                          Return "Personal PY"
                        End If
                        If (plmnid = 74406) Then
                          Return "Copaco"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 302270) Then
                  If (plmnid < 90115) Then
                    If (plmnid < 74810) Then
                      If (plmnid = 74602) Then
                        Return "Telesur"
                      End If
                      If (plmnid = 74603) Then
                        Return "Digicel SR"
                      End If
                      If (plmnid = 74801) Then
                        Return "Antel"
                      End If
                      If (plmnid = 74807) Then
                        Return "Movistar UY"
                      End If
                    Else
                      If (plmnid = 74810) Then
                        Return "Claro UY"
                      End If
                      If (plmnid = 90112) Then
                        Return "Telenor INTL"
                      End If
                      If (plmnid = 90113) Then
                        Return "GSM.AQ"
                      End If
                      If (plmnid = 90114) Then
                        Return "AeroMobile"
                      End If
                    End If
                  Else
                    If (plmnid < 90127) Then
                      If (plmnid = 90115) Then
                        Return "OnAir"
                      End If
                      If (plmnid = 90118) Then
                        Return "Cellular @Sea"
                      End If
                      If (plmnid = 90119) Then
                        Return "Vodafone Malta Maritime"
                      End If
                      If (plmnid = 90126) Then
                        Return "TIM@sea"
                      End If
                    Else
                      If (plmnid < 90132) Then
                        If (plmnid = 90127) Then
                          Return "OnMarine"
                        End If
                        If (plmnid = 90131) Then
                          Return "Orange INTL"
                        End If
                      Else
                        If (plmnid = 90132) Then
                          Return "Sky High"
                        End If
                        If (plmnid = 99501) Then
                          Return "FonePlus"
                        End If
                        If (plmnid = 302220) Then
                          Return "Telus Mobility, Koodo Mobile, Public Mobile"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 302720) Then
                    If (plmnid < 302510) Then
                      If (plmnid = 302270) Then
                        Return "EastLink"
                      End If
                      If (plmnid = 302480) Then
                        Return "Qiniq"
                      End If
                      If (plmnid = 302490) Then
                        Return "Freedom Mobile"
                      End If
                      If (plmnid = 302500) Then
                        Return "Videotron"
                      End If
                    Else
                      If (plmnid < 302610) Then
                        If (plmnid = 302510) Then
                          Return "Videotron"
                        End If
                        If (plmnid = 302530) Then
                          Return "Keewaytinook Mobile"
                        End If
                      Else
                        If (plmnid = 302610) Then
                          Return "Bell Mobility"
                        End If
                        If (plmnid = 302620) Then
                          Return "ICE Wireless"
                        End If
                        If (plmnid = 302660) Then
                          Return "MTS CA"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 310030) Then
                      If (plmnid = 302720) Then
                        Return "Rogers Wireless"
                      End If
                      If (plmnid = 302780) Then
                        Return "SaskTel"
                      End If
                      If (plmnid = 310012) Then
                        Return "Verizon"
                      End If
                      If (plmnid = 310020) Then
                        Return "Union Wireless"
                      End If
                    Else
                      If (plmnid < 310032) Then
                        If (plmnid = 310030) Then
                          Return "AT&T US"
                        End If
                        If (plmnid = 310032) Then
                          Return "IT&E Wireless GU"
                        End If
                      Else
                        If (plmnid = 310032) Then
                          Return "IT&E Wireless US"
                        End If
                        If (plmnid = 310066) Then
                          Return "U.S. Cellular"
                        End If
                        If (plmnid = 310070) Then
                          Return "AT&T US"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If (plmnid < 311080) Then
                If (plmnid < 310370) Then
                  If (plmnid < 310150) Then
                    If (plmnid < 310110) Then
                      If (plmnid = 310080) Then
                        Return "AT&T US"
                      End If
                      If (plmnid = 310090) Then
                        Return "AT&T US"
                      End If
                      If (plmnid = 310100) Then
                        Return "Plateau Wireless"
                      End If
                      If (plmnid = 310110) Then
                        Return "IT&E Wireless MP"
                      End If
                    Else
                      If (plmnid = 310110) Then
                        Return "IT&E Wireless US"
                      End If
                      If (plmnid = 310120) Then
                        Return "Sprint"
                      End If
                      If (plmnid = 310140) Then
                        Return "GTA Wireless GU"
                      End If
                      If (plmnid = 310140) Then
                        Return "GTA Wireless US"
                      End If
                    End If
                  Else
                    If (plmnid < 310190) Then
                      If (plmnid = 310150) Then
                        Return "AT&T US"
                      End If
                      If (plmnid = 310160) Then
                        Return "T-Mobile US"
                      End If
                      If (plmnid = 310170) Then
                        Return "AT&T US"
                      End If
                      If (plmnid = 310180) Then
                        Return "West Central"
                      End If
                    Else
                      If (plmnid < 310320) Then
                        If (plmnid = 310190) Then
                          Return "GCI"
                        End If
                        If (plmnid = 310260) Then
                          Return "T-Mobile US"
                        End If
                      Else
                        If (plmnid = 310320) Then
                          Return "Cellular One"
                        End If
                        If (plmnid = 310370) Then
                          Return "Docomo GU"
                        End If
                        If (plmnid = 310370) Then
                          Return "Docomo MP"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 310740) Then
                    If (plmnid < 310410) Then
                      If (plmnid = 310370) Then
                        Return "Docomo US"
                      End If
                      If (plmnid = 310390) Then
                        Return "Cellular One of East Texas"
                      End If
                      If (plmnid = 310400) Then
                        Return "iConnect GU"
                      End If
                      If (plmnid = 310400) Then
                        Return "iConnect US"
                      End If
                    Else
                      If (plmnid < 310450) Then
                        If (plmnid = 310410) Then
                          Return "AT&T US"
                        End If
                        If (plmnid = 310430) Then
                          Return "GCI"
                        End If
                      Else
                        If (plmnid = 310450) Then
                          Return "Viaero"
                        End If
                        If (plmnid = 310540) Then
                          Return "Phoenix US"
                        End If
                        If (plmnid = 310680) Then
                          Return "AT&T US"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 310990) Then
                      If (plmnid = 310740) Then
                        Return "Viaero"
                      End If
                      If (plmnid = 310770) Then
                        Return "iWireless"
                      End If
                      If (plmnid = 310790) Then
                        Return "BLAZE"
                      End If
                      If (plmnid = 310950) Then
                        Return "AT&T US"
                      End If
                    Else
                      If (plmnid < 311030) Then
                        If (plmnid = 310990) Then
                          Return "Evolve Broadband"
                        End If
                        If (plmnid = 311020) Then
                          Return "Chariton Valley"
                        End If
                      Else
                        If (plmnid = 311030) Then
                          Return "Indigo Wireless"
                        End If
                        If (plmnid = 311040) Then
                          Return "Choice Wireless"
                        End If
                        If (plmnid = 311070) Then
                          Return "AT&T US"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 311670) Then
                  If (plmnid < 311470) Then
                    If (plmnid < 311320) Then
                      If (plmnid = 311080) Then
                        Return "Pine Cellular"
                      End If
                      If (plmnid = 311090) Then
                        Return "AT&T US"
                      End If
                      If (plmnid = 311230) Then
                        Return "C Spire Wireless"
                      End If
                      If (plmnid = 311290) Then
                        Return "BLAZE"
                      End If
                    Else
                      If (plmnid = 311320) Then
                        Return "Choice Wireless"
                      End If
                      If (plmnid = 311330) Then
                        Return "Bug Tussel Wireless"
                      End If
                      If (plmnid = 311370) Then
                        Return "GCI Wireless"
                      End If
                      If (plmnid = 311450) Then
                        Return "PTCI"
                      End If
                    End If
                  Else
                    If (plmnid < 311550) Then
                      If (plmnid = 311470) Then
                        Return "Viya"
                      End If
                      If (plmnid = 311480) Then
                        Return "Verizon"
                      End If
                      If (plmnid = 311490) Then
                        Return "Sprint"
                      End If
                      If (plmnid = 311530) Then
                        Return "NewCore"
                      End If
                    Else
                      If (plmnid < 311580) Then
                        If (plmnid = 311550) Then
                          Return "Choice Wireless"
                        End If
                        If (plmnid = 311560) Then
                          Return "OTZ Cellular"
                        End If
                      Else
                        If (plmnid = 311580) Then
                          Return "U.S. Cellular"
                        End If
                        If (plmnid = 311640) Then
                          Return "Rock Wireless"
                        End If
                        If (plmnid = 311650) Then
                          Return "United Wireless"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 312120) Then
                    If (plmnid < 311850) Then
                      If (plmnid = 311670) Then
                        Return "Pine Belt Wireless"
                      End If
                      If (plmnid = 311780) Then
                        Return "ASTCA US"
                      End If
                      If (plmnid = 311780) Then
                        Return "ASTCA AS"
                      End If
                      If (plmnid = 311840) Then
                        Return "Cellcom US"
                      End If
                    Else
                      If (plmnid < 311950) Then
                        If (plmnid = 311850) Then
                          Return "Cellcom US"
                        End If
                        If (plmnid = 311860) Then
                          Return "STRATA"
                        End If
                      Else
                        If (plmnid = 311950) Then
                          Return "ETC"
                        End If
                        If (plmnid = 311970) Then
                          Return "Big River Broadband"
                        End If
                        If (plmnid = 312030) Then
                          Return "Bravado Wireless"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 312170) Then
                      If (plmnid = 312120) Then
                        Return "Appalachian Wireless"
                      End If
                      If (plmnid = 312130) Then
                        Return "Appalachian Wireless"
                      End If
                      If (plmnid = 312150) Then
                        Return "NorthwestCell"
                      End If
                      If (plmnid = 312160) Then
                        Return "Chat Mobility"
                      End If
                    Else
                      If (plmnid < 312260) Then
                        If (plmnid = 312170) Then
                          Return "Chat Mobility"
                        End If
                        If (plmnid = 312220) Then
                          Return "Chariton Valley"
                        End If
                      Else
                        If (plmnid = 312260) Then
                          Return "NewCore"
                        End If
                        If (plmnid = 312270) Then
                          Return "Pioneer Cellular"
                        End If
                        If (plmnid = 312280) Then
                          Return "Pioneer Cellular"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          Else
            If (plmnid < 405803) Then
              If (plmnid < 360110) Then
                If (plmnid < 338180) Then
                  If (plmnid < 330120) Then
                    If (plmnid < 312860) Then
                      If (plmnid = 312330) Then
                        Return "Nemont"
                      End If
                      If (plmnid = 312400) Then
                        Return "Mid-Rivers Wireless"
                      End If
                      If (plmnid = 312470) Then
                        Return "Carolina West Wireless"
                      End If
                      If (plmnid = 312720) Then
                        Return "Southern LINC"
                      End If
                    Else
                      If (plmnid = 312860) Then
                        Return "ClearTalk"
                      End If
                      If (plmnid = 312900) Then
                        Return "ClearTalk"
                      End If
                      If (plmnid = 313100) Then
                        Return "FirstNet"
                      End If
                      If (plmnid = 330110) Then
                        Return "Claro Puerto Rico"
                      End If
                    End If
                  Else
                    If (plmnid < 334140) Then
                      If (plmnid = 330120) Then
                        Return "Open Mobile"
                      End If
                      If (plmnid = 334020) Then
                        Return "Telcel"
                      End If
                      If (plmnid = 334050) Then
                        Return "AT&T / Unefon"
                      End If
                      If (plmnid = 334090) Then
                        Return "AT&T MX"
                      End If
                    Else
                      If (plmnid < 338050) Then
                        If (plmnid = 334140) Then
                          Return "Altan Redes"
                        End If
                        If (plmnid = 338050) Then
                          Return "Digicel JM"
                        End If
                      Else
                        If (plmnid = 338050) Then
                          Return "Digicel LC"
                        End If
                        If (plmnid = 338050) Then
                          Return "Digicel TC"
                        End If
                        If (plmnid = 338050) Then
                          Return "Digicel Bermuda"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 348570) Then
                    If (plmnid < 344050) Then
                      If (plmnid = 338180) Then
                        Return "FLOW JM"
                      End If
                      If (plmnid = 342600) Then
                        Return "FLOW BB"
                      End If
                      If (plmnid = 342750) Then
                        Return "Digicel BB"
                      End If
                      If (plmnid = 344030) Then
                        Return "APUA"
                      End If
                    Else
                      If (plmnid < 346050) Then
                        If (plmnid = 344050) Then
                          Return "Digicel AG"
                        End If
                        If (plmnid = 344920) Then
                          Return "FLOW AG"
                        End If
                      Else
                        If (plmnid = 346050) Then
                          Return "Digicel KY"
                        End If
                        If (plmnid = 346140) Then
                          Return "FLOW KY"
                        End If
                        If (plmnid = 348170) Then
                          Return "FLOW VG"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 354860) Then
                      If (plmnid = 348570) Then
                        Return "CCT Boatphone"
                      End If
                      If (plmnid = 348770) Then
                        Return "Digicel VG"
                      End If
                      If (plmnid = 352030) Then
                        Return "Digicel GD"
                      End If
                      If (plmnid = 352110) Then
                        Return "FLOW GD"
                      End If
                    Else
                      If (plmnid < 356110) Then
                        If (plmnid = 354860) Then
                          Return "FLOW MS"
                        End If
                        If (plmnid = 356050) Then
                          Return "Digicel KN"
                        End If
                      Else
                        If (plmnid = 356110) Then
                          Return "FLOW KN"
                        End If
                        If (plmnid = 358110) Then
                          Return "FLOW LC"
                        End If
                        If (plmnid = 360050) Then
                          Return "Digicel VC"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 405039) Then
                  If (plmnid < 405028) Then
                    If (plmnid < 374130) Then
                      If (plmnid = 360110) Then
                        Return "FLOW VC"
                      End If
                      If (plmnid = 365840) Then
                        Return "FLOW AI"
                      End If
                      If (plmnid = 366020) Then
                        Return "Digicel DM"
                      End If
                      If (plmnid = 366110) Then
                        Return "FLOW DM"
                      End If
                    Else
                      If (plmnid = 374130) Then
                        Return "Digicel TT"
                      End If
                      If (plmnid = 376350) Then
                        Return "FLOW TC"
                      End If
                      If (plmnid = 405025) Then
                        Return "TATA DOCOMO"
                      End If
                      If (plmnid = 405027) Then
                        Return "TATA DOCOMO"
                      End If
                    End If
                  Else
                    If (plmnid < 405034) Then
                      If (plmnid = 405028) Then
                        Return "TATA DOCOMO"
                      End If
                      If (plmnid = 405030) Then
                        Return "TATA DOCOMO"
                      End If
                      If (plmnid = 405031) Then
                        Return "TATA DOCOMO"
                      End If
                      If (plmnid = 405032) Then
                        Return "TATA DOCOMO"
                      End If
                    Else
                      If (plmnid < 405036) Then
                        If (plmnid = 405034) Then
                          Return "TATA DOCOMO"
                        End If
                        If (plmnid = 405035) Then
                          Return "TATA DOCOMO"
                        End If
                      Else
                        If (plmnid = 405036) Then
                          Return "TATA DOCOMO"
                        End If
                        If (plmnid = 405037) Then
                          Return "TATA DOCOMO"
                        End If
                        If (plmnid = 405038) Then
                          Return "TATA DOCOMO"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 405751) Then
                    If (plmnid < 405044) Then
                      If (plmnid = 405039) Then
                        Return "TATA DOCOMO"
                      End If
                      If (plmnid = 405041) Then
                        Return "TATA DOCOMO"
                      End If
                      If (plmnid = 405042) Then
                        Return "TATA DOCOMO"
                      End If
                      If (plmnid = 405043) Then
                        Return "TATA DOCOMO"
                      End If
                    Else
                      If (plmnid < 405046) Then
                        If (plmnid = 405044) Then
                          Return "TATA DOCOMO"
                        End If
                        If (plmnid = 405045) Then
                          Return "TATA DOCOMO"
                        End If
                      Else
                        If (plmnid = 405046) Then
                          Return "TATA DOCOMO"
                        End If
                        If (plmnid = 405047) Then
                          Return "TATA DOCOMO"
                        End If
                        If (plmnid = 405750) Then
                          Return "Vodafone India"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 405755) Then
                      If (plmnid = 405751) Then
                        Return "Vodafone India"
                      End If
                      If (plmnid = 405752) Then
                        Return "Vodafone India"
                      End If
                      If (plmnid = 405753) Then
                        Return "Vodafone India"
                      End If
                      If (plmnid = 405754) Then
                        Return "Vodafone India"
                      End If
                    Else
                      If (plmnid < 405799) Then
                        If (plmnid = 405755) Then
                          Return "Vodafone India"
                        End If
                        If (plmnid = 405756) Then
                          Return "Vodafone India"
                        End If
                      Else
                        If (plmnid = 405799) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 405800) Then
                          Return "AIRCEL"
                        End If
                        If (plmnid = 405801) Then
                          Return "AIRCEL"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If (plmnid < 405866) Then
                If (plmnid < 405848) Then
                  If (plmnid < 405819) Then
                    If (plmnid < 405809) Then
                      If (plmnid = 405803) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 405804) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 405805) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 405806) Then
                        Return "AIRCEL"
                      End If
                    Else
                      If (plmnid = 405809) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 405810) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 405811) Then
                        Return "AIRCEL"
                      End If
                      If (plmnid = 405818) Then
                        Return "Uninor"
                      End If
                    End If
                  Else
                    If (plmnid < 405827) Then
                      If (plmnid = 405819) Then
                        Return "Uninor"
                      End If
                      If (plmnid = 405820) Then
                        Return "Uninor"
                      End If
                      If (plmnid = 405821) Then
                        Return "Uninor"
                      End If
                      If (plmnid = 405822) Then
                        Return "Uninor"
                      End If
                    Else
                      If (plmnid < 405845) Then
                        If (plmnid = 405827) Then
                          Return "Videocon Datacom"
                        End If
                        If (plmnid = 405840) Then
                          Return "Jio"
                        End If
                      Else
                        If (plmnid = 405845) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 405846) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 405847) Then
                          Return "IDEA"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 405857) Then
                    If (plmnid < 405852) Then
                      If (plmnid = 405848) Then
                        Return "IDEA"
                      End If
                      If (plmnid = 405849) Then
                        Return "IDEA"
                      End If
                      If (plmnid = 405850) Then
                        Return "IDEA"
                      End If
                      If (plmnid = 405851) Then
                        Return "IDEA"
                      End If
                    Else
                      If (plmnid < 405854) Then
                        If (plmnid = 405852) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 405853) Then
                          Return "IDEA"
                        End If
                      Else
                        If (plmnid = 405854) Then
                          Return "Jio"
                        End If
                        If (plmnid = 405855) Then
                          Return "Jio"
                        End If
                        If (plmnid = 405856) Then
                          Return "Jio"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 405861) Then
                      If (plmnid = 405857) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405858) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405859) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405860) Then
                        Return "Jio"
                      End If
                    Else
                      If (plmnid < 405863) Then
                        If (plmnid = 405861) Then
                          Return "Jio"
                        End If
                        If (plmnid = 405862) Then
                          Return "Jio"
                        End If
                      Else
                        If (plmnid = 405863) Then
                          Return "Jio"
                        End If
                        If (plmnid = 405864) Then
                          Return "Jio"
                        End If
                        If (plmnid = 405865) Then
                          Return "Jio"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                If (plmnid < 708002) Then
                  If (plmnid < 405874) Then
                    If (plmnid < 405870) Then
                      If (plmnid = 405866) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405867) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405868) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405869) Then
                        Return "Jio"
                      End If
                    Else
                      If (plmnid = 405870) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405871) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405872) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405873) Then
                        Return "Jio"
                      End If
                    End If
                  Else
                    If (plmnid < 405910) Then
                      If (plmnid = 405874) Then
                        Return "Jio"
                      End If
                      If (plmnid = 405880) Then
                        Return "Uninor"
                      End If
                      If (plmnid = 405908) Then
                        Return "IDEA"
                      End If
                      If (plmnid = 405909) Then
                        Return "IDEA"
                      End If
                    Else
                      If (plmnid < 405929) Then
                        If (plmnid = 405910) Then
                          Return "IDEA"
                        End If
                        If (plmnid = 405911) Then
                          Return "IDEA"
                        End If
                      Else
                        If (plmnid = 405929) Then
                          Return "Uninor"
                        End If
                        If (plmnid = 502153) Then
                          Return "unifi"
                        End If
                        If (plmnid = 708001) Then
                          Return "Claro HN"
                        End If
                      End If
                    End If
                  End If
                Else
                  If (plmnid < 732099) Then
                    If (plmnid < 722070) Then
                      If (plmnid = 708002) Then
                        Return "Tigo HN"
                      End If
                      If (plmnid = 708030) Then
                        Return "Hondutel"
                      End If
                      If (plmnid = 714020) Then
                        Return "movistar PA"
                      End If
                      If (plmnid = 722010) Then
                        Return "Movistar AR"
                      End If
                    Else
                      If (plmnid < 722320) Then
                        If (plmnid = 722070) Then
                          Return "Movistar AR"
                        End If
                        If (plmnid = 722310) Then
                          Return "Claro AR"
                        End If
                      Else
                        If (plmnid = 722320) Then
                          Return "Claro AR"
                        End If
                        If (plmnid = 722330) Then
                          Return "Claro AR"
                        End If
                        If (plmnid = 722341) Then
                          Return "Personal AR"
                        End If
                      End If
                    End If
                  Else
                    If (plmnid < 732123) Then
                      If (plmnid = 732099) Then
                        Return "EMCALI"
                      End If
                      If (plmnid = 732101) Then
                        Return "Claro CO"
                      End If
                      If (plmnid = 732103) Then
                        Return "Tigo CO"
                      End If
                      If (plmnid = 732111) Then
                        Return "Tigo CO"
                      End If
                    Else
                      If (plmnid < 732187) Then
                        If (plmnid = 732123) Then
                          Return "movistar CO"
                        End If
                        If (plmnid = 732130) Then
                          Return "AVANTEL"
                        End If
                      Else
                        If (plmnid = 732187) Then
                          Return "eTb"
                        End If
                        If (plmnid = 738002) Then
                          Return "GT&T Cellink Plus"
                        End If
                        If (plmnid = 750001) Then
                          Return "Sure FK"
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
      Return mccmnc
    End Function

    '''*
    ''' <summary>
    '''   Returns the cell operator brand for a given MCC/MNC pair.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="mccmnc">
    '''   a string starting with a MCC code followed by a MNC code,
    ''' </param>
    ''' <returns>
    '''   a string containing the corresponding cell operator brand name.
    ''' </returns>
    '''/
    Public Overridable Function decodePLMN(mccmnc As String) As String
      Return Me.imm_decodePLMN(mccmnc)
    End Function

    '''*
    ''' <summary>
    '''   Returns the list available radio communication profiles, as a string array
    '''   (YoctoHub-GSM-4G only).
    ''' <para>
    '''   Each string is a made of a numerical ID, followed by a colon,
    '''   followed by the profile description.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of string describing available radio communication profiles.
    ''' </returns>
    '''/
    Public Overridable Function get_communicationProfiles() As List(Of String)
      Dim profiles As String
      Dim lines As List(Of String) = New List(Of String)()
      Dim nlines As Integer = 0
      Dim idx As Integer = 0
      Dim line As String
      Dim cpos As Integer = 0
      Dim profno As Integer = 0
      Dim res As List(Of String) = New List(Of String)()

      profiles = Me._AT("+UMNOPROF=?")
      lines = New List(Of String)(profiles.Split(vbLf.ToCharArray()))
      nlines = lines.Count
      If Not(nlines > 0) Then
        me._throw( YAPI.IO_ERROR,  "fail to retrieve profile list")
        return res
      end if
      res.Clear()
      idx = 0
      While (idx < nlines)
        line = lines(idx)
        cpos = line.IndexOf(":")
        If (cpos > 0) Then
          profno = YAPI._atoi((line).Substring( 0, cpos))
          If (profno > 1) Then
            res.Add(line)
          End If
        End If
        idx = idx + 1
      End While

      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of cellular interfaces started using <c>yFirstCellular()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned cellular interfaces order.
    '''   If you want to find a specific a cellular interface, use <c>Cellular.findCellular()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCellular</c> object, corresponding to
    '''   a cellular interface currently online, or a <c>Nothing</c> pointer
    '''   if there are no more cellular interfaces to enumerate.
    ''' </returns>
    '''/
    Public Function nextCellular() As YCellular
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YCellular.FindCellular(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of cellular interfaces currently accessible.
    ''' <para>
    '''   Use the method <c>YCellular.nextCellular()</c> to iterate on
    '''   next cellular interfaces.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCellular</c> object, corresponding to
    '''   the first cellular interface currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstCellular() As YCellular
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Cellular", 0, p, size, neededsize, errmsg)
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
      Return YCellular.FindCellular(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YCellular public methods declaration)

  End Class

  REM --- (generated code: YCellular functions)

  '''*
  ''' <summary>
  '''   Retrieves a cellular interface for a given identifier.
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
  '''   This function does not require that the cellular interface is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YCellular.isOnline()</c> to test if the cellular interface is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a cellular interface by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the cellular interface, for instance
  '''   <c>YHUBGSM1.cellular</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YCellular</c> object allowing you to drive the cellular interface.
  ''' </returns>
  '''/
  Public Function yFindCellular(ByVal func As String) As YCellular
    Return YCellular.FindCellular(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of cellular interfaces currently accessible.
  ''' <para>
  '''   Use the method <c>YCellular.nextCellular()</c> to iterate on
  '''   next cellular interfaces.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YCellular</c> object, corresponding to
  '''   the first cellular interface currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstCellular() As YCellular
    Return YCellular.FirstCellular()
  End Function


  REM --- (end of generated code: YCellular functions)

End Module
