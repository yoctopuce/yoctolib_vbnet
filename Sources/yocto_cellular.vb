'*********************************************************************
'*
'* $Id: yocto_cellular.vb 19745 2015-03-17 09:47:45Z seb $
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
    Public Overridable Function get_cellOperator() As String
      Return Me._oper
    End Function

    Public Overridable Function get_mobileCountryCode() As Integer
      Return Me._mcc
    End Function

    Public Overridable Function get_mobileNetworkCode() As Integer
      Return Me._mnc
    End Function

    Public Overridable Function get_locationAreaCode() As Integer
      Return Me._lac
    End Function

    Public Overridable Function get_cellId() As Integer
      Return Me._cid
    End Function

    Public Overridable Function get_signalStrength() As Integer
      Return Me._dbm
    End Function

    Public Overridable Function get_timingAdvance() As Integer
      Return Me._tad
    End Function



    REM --- (end of generated code: YCellRecord public methods declaration)

  End Class

  REM --- (generated code: CellRecord functions)


  REM --- (end of generated code: CellRecord functions)



    REM --- (generated code: YCellular return codes)
    REM --- (end of generated code: YCellular return codes)
    REM --- (generated code: YCellular dlldef)
    REM --- (end of generated code: YCellular dlldef)
  REM --- (generated code: YCellular globals)

  Public Const Y_LINKQUALITY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CELLOPERATOR_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_MESSAGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PIN_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_LOCKEDOPERATOR_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ENABLEDATA_HOMENETWORK As Integer = 0
  Public Const Y_ENABLEDATA_ROAMING As Integer = 1
  Public Const Y_ENABLEDATA_NEVER As Integer = 2
  Public Const Y_ENABLEDATA_INVALID As Integer = -1
  Public Const Y_APN_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_APNSECRET_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YCellularValueCallback(ByVal func As YCellular, ByVal value As String)
  Public Delegate Sub YCellularTimedReportCallback(ByVal func As YCellular, ByVal measure As YMeasure)
  REM --- (end of generated code: YCellular globals)

  REM --- (generated code: YCellular class start)

  '''*
  ''' <summary>
  '''   YCellular functions provides control over cellular network parameters
  '''   and status for devices that are GSM-enabled.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YCellular
    Inherits YFunction
    REM --- (end of generated code: YCellular class start)

    REM --- (generated code: YCellular definitions)
    Public Const LINKQUALITY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CELLOPERATOR_INVALID As String = YAPI.INVALID_STRING
    Public Const MESSAGE_INVALID As String = YAPI.INVALID_STRING
    Public Const PIN_INVALID As String = YAPI.INVALID_STRING
    Public Const LOCKEDOPERATOR_INVALID As String = YAPI.INVALID_STRING
    Public Const ENABLEDATA_HOMENETWORK As Integer = 0
    Public Const ENABLEDATA_ROAMING As Integer = 1
    Public Const ENABLEDATA_NEVER As Integer = 2
    Public Const ENABLEDATA_INVALID As Integer = -1
    Public Const APN_INVALID As String = YAPI.INVALID_STRING
    Public Const APNSECRET_INVALID As String = YAPI.INVALID_STRING
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of generated code: YCellular definitions)

    REM --- (generated code: YCellular attributes declaration)
    Protected _linkQuality As Integer
    Protected _cellOperator As String
    Protected _message As String
    Protected _pin As String
    Protected _lockedOperator As String
    Protected _enableData As Integer
    Protected _apn As String
    Protected _apnSecret As String
    Protected _command As String
    Protected _valueCallbackCellular As YCellularValueCallback
    REM --- (end of generated code: YCellular attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Cellular"
      REM --- (generated code: YCellular attributes initialization)
      _linkQuality = LINKQUALITY_INVALID
      _cellOperator = CELLOPERATOR_INVALID
      _message = MESSAGE_INVALID
      _pin = PIN_INVALID
      _lockedOperator = LOCKEDOPERATOR_INVALID
      _enableData = ENABLEDATA_INVALID
      _apn = APN_INVALID
      _apnSecret = APNSECRET_INVALID
      _command = COMMAND_INVALID
      _valueCallbackCellular = Nothing
      REM --- (end of generated code: YCellular attributes initialization)
    End Sub

    REM --- (generated code: YCellular private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "linkQuality") Then
        _linkQuality = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "cellOperator") Then
        _cellOperator = member.svalue
        Return 1
      End If
      If (member.name = "message") Then
        _message = member.svalue
        Return 1
      End If
      If (member.name = "pin") Then
        _pin = member.svalue
        Return 1
      End If
      If (member.name = "lockedOperator") Then
        _lockedOperator = member.svalue
        Return 1
      End If
      If (member.name = "enableData") Then
        _enableData = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "apn") Then
        _apn = member.svalue
        Return 1
      End If
      If (member.name = "apnSecret") Then
        _apnSecret = member.svalue
        Return 1
      End If
      If (member.name = "command") Then
        _command = member.svalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
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
    '''   On failure, throws an exception or returns <c>Y_LINKQUALITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_linkQuality() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LINKQUALITY_INVALID
        End If
      End If
      Return Me._linkQuality
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
    '''   On failure, throws an exception or returns <c>Y_CELLOPERATOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_cellOperator() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CELLOPERATOR_INVALID
        End If
      End If
      Return Me._cellOperator
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
    '''   On failure, throws an exception or returns <c>Y_MESSAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_message() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MESSAGE_INVALID
        End If
      End If
      Return Me._message
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
    '''   On failure, throws an exception or returns <c>Y_PIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pin() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PIN_INVALID
        End If
      End If
      Return Me._pin
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   On failure, throws an exception or returns <c>Y_LOCKEDOPERATOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lockedOperator() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LOCKEDOPERATOR_INVALID
        End If
      End If
      Return Me._lockedOperator
    End Function


    '''*
    ''' <summary>
    '''   Changes the name of the cell operator to be used.
    ''' <para>
    '''   If the name is an empty
    '''   string, the choice will be made automatically based on the SIM card. Otherwise,
    '''   the selected operator is the only one that will be used.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   Returns the condition for enabling IP data services (GPRS).
    ''' <para>
    '''   When data services are disabled, SMS are the only mean of communication.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_ENABLEDATA_HOMENETWORK</c>, <c>Y_ENABLEDATA_ROAMING</c> and
    '''   <c>Y_ENABLEDATA_NEVER</c> corresponding to the condition for enabling IP data services (GPRS)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ENABLEDATA_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_enableData() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ENABLEDATA_INVALID
        End If
      End If
      Return Me._enableData
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
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_ENABLEDATA_HOMENETWORK</c>, <c>Y_ENABLEDATA_ROAMING</c> and
    '''   <c>Y_ENABLEDATA_NEVER</c> corresponding to the condition for enabling IP data services (GPRS)
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
    '''   On failure, throws an exception or returns <c>Y_APN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_apn() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return APN_INVALID
        End If
      End If
      Return Me._apn
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
    ''' <param name="newval">
    '''   a string
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
    '''   On failure, throws an exception or returns <c>Y_APNSECRET_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_apnSecret() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return APNSECRET_INVALID
        End If
      End If
      Return Me._apnSecret
    End Function


    Public Function set_apnSecret(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("apnSecret", rest_val)
    End Function
    Public Function get_command() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      Return Me._command
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
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the cellular interface
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
    Public Overloads Function registerValueCallback(callback As YCellularValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackCellular = callback
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
    '''   Only ten consecutives tentatives are permitted:
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function sendPUK(puk As String, newPin As String) As Integer
      Dim gsmMsg As String
      
      gsmMsg = Me.get_message()
      If Not(gsmMsg = "Enter SIM PUK") Then
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
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
      Dim content As Byte()
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
      
      REM // may throw an exception
      content = Me._download("at.txt?cmd=" + cmd)
      Return YAPI.DefaultEncoding.GetString(content)
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of nearby cellular antennas, as required for quick
    '''   geolocation of the device.
    ''' <para>
    '''   The first cell listed is the serving
    '''   cell, and the next ones are the neighboor cells reported by the
    '''   serving cell.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of YCellRecords.
    ''' </returns>
    '''/
    Public Overridable Function quickCellSurvey() As List(Of YCellRecord)
      Dim i_i As Integer
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
      REM // may throw an exception
      moni = Me._AT("+CCED=0;#MONI=7;#MONI")
      mccs = (moni).Substring(7, 3)
      If ((mccs).Substring(0, 1) = "0") Then
        mccs = (mccs).Substring(1, 2)
      End If
      If ((mccs).Substring(0, 1) = "0") Then
        mccs = (mccs).Substring(1, 1)
      End If
      mcc = Convert.ToInt32(mccs)
      mncs = (moni).Substring(11, 3)
      If ((mncs).Substring(2, 1) = ",") Then
        mncs = (mncs).Substring(0, 2)
      End If
      If ((mncs).Substring(0, 1) = "0") Then
        mncs = (mncs).Substring(1, (mncs).Length-1)
      End If
      mnc = Convert.ToInt32(mncs)
      recs = new List(Of String)(moni.Split(new Char() {"#"c}))
      REM // process each line in turn
      res.Clear()
      For i_i = 0 To recs.Count - 1
        llen = (recs(i_i)).Length - 2
        If (llen >= 44) Then
          If ((recs(i_i)).Substring(41, 3) = "dbm") Then
            lac = Convert.ToInt32((recs(i_i)).Substring(16, 4), 16)
            cellId = Convert.ToInt32((recs(i_i)).Substring(23, 4), 16)
            dbms = (recs(i_i)).Substring(37, 4)
            If ((dbms).Substring(0, 1) = " ") Then
              dbms = (dbms).Substring(1, 3)
            End If
            dbm = Convert.ToInt32(dbms)
            If (llen > 66) Then
              tads = (recs(i_i)).Substring(54, 2)
              If ((tads).Substring(0, 1) = " ") Then
                tads = (tads).Substring(1, 3)
              End If
              tad = Convert.ToInt32(tads)
              oper = (recs(i_i)).Substring(66, llen-66)
            Else
              tad = -1
              oper = ""
            End If
            If (lac < 65535) Then
              res.Add(New YCellRecord(mcc, mnc, lac, cellId, dbm, tad, oper))
            End If
          End If
        End If
      Next i_i
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of cellular interfaces started using <c>yFirstCellular()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YCellular</c> object, corresponding to
    '''   a cellular interface currently online, or a <c>null</c> pointer
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
    '''   the first cellular interface currently online, or a <c>null</c> pointer
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

  REM --- (generated code: Cellular functions)

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
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the cellular interface
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
  '''   the first cellular interface currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstCellular() As YCellular
    Return YCellular.FirstCellular()
  End Function


  REM --- (end of generated code: Cellular functions)

End Module
