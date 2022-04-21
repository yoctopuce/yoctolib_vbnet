' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindInputChain(), the high-level API for InputChain functions
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

Module yocto_inputchain

    REM --- (YInputChain return codes)
    REM --- (end of YInputChain return codes)
    REM --- (YInputChain dlldef)
    REM --- (end of YInputChain dlldef)
   REM --- (YInputChain yapiwrapper)
   REM --- (end of YInputChain yapiwrapper)
  REM --- (YInputChain globals)

  Public Const Y_EXPECTEDNODES_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_DETECTEDNODES_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_LOOPBACKTEST_OFF As Integer = 0
  Public Const Y_LOOPBACKTEST_ON As Integer = 1
  Public Const Y_LOOPBACKTEST_INVALID As Integer = -1
  Public Const Y_REFRESHRATE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BITCHAIN1_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_BITCHAIN2_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_BITCHAIN3_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_BITCHAIN4_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_BITCHAIN5_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_BITCHAIN6_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_BITCHAIN7_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_WATCHDOGPERIOD_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CHAINDIAGS_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YInputChainValueCallback(ByVal func As YInputChain, ByVal value As String)
  Public Delegate Sub YInputChainTimedReportCallback(ByVal func As YInputChain, ByVal measure As YMeasure)
  Public Delegate Sub YEventCallback(ByVal inputChain As YInputChain, ByVal timestamp As Integer, ByVal eventType As String, ByVal eventData As String, ByVal eventChange As String)

  Sub yInternalEventCallback(ByVal func As YInputChain, ByVal value As String)
    func._internalEventHandler(value)
  End Sub
  REM --- (end of YInputChain globals)

  REM --- (YInputChain class start)

  '''*
  ''' <summary>
  '''   The <c>YInputChain</c> class provides access to separate
  '''   digital inputs connected in a chain.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YInputChain
    Inherits YFunction
    REM --- (end of YInputChain class start)

    REM --- (YInputChain definitions)
    Public Const EXPECTEDNODES_INVALID As Integer = YAPI.INVALID_UINT
    Public Const DETECTEDNODES_INVALID As Integer = YAPI.INVALID_UINT
    Public Const LOOPBACKTEST_OFF As Integer = 0
    Public Const LOOPBACKTEST_ON As Integer = 1
    Public Const LOOPBACKTEST_INVALID As Integer = -1
    Public Const REFRESHRATE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BITCHAIN1_INVALID As String = YAPI.INVALID_STRING
    Public Const BITCHAIN2_INVALID As String = YAPI.INVALID_STRING
    Public Const BITCHAIN3_INVALID As String = YAPI.INVALID_STRING
    Public Const BITCHAIN4_INVALID As String = YAPI.INVALID_STRING
    Public Const BITCHAIN5_INVALID As String = YAPI.INVALID_STRING
    Public Const BITCHAIN6_INVALID As String = YAPI.INVALID_STRING
    Public Const BITCHAIN7_INVALID As String = YAPI.INVALID_STRING
    Public Const WATCHDOGPERIOD_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CHAINDIAGS_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YInputChain definitions)

    REM --- (YInputChain attributes declaration)
    Protected _expectedNodes As Integer
    Protected _detectedNodes As Integer
    Protected _loopbackTest As Integer
    Protected _refreshRate As Integer
    Protected _bitChain1 As String
    Protected _bitChain2 As String
    Protected _bitChain3 As String
    Protected _bitChain4 As String
    Protected _bitChain5 As String
    Protected _bitChain6 As String
    Protected _bitChain7 As String
    Protected _watchdogPeriod As Integer
    Protected _chainDiags As Integer
    Protected _valueCallbackInputChain As YInputChainValueCallback
    Protected _eventCallback As YEventCallback
    Protected _prevPos As Integer
    Protected _eventPos As Integer
    Protected _eventStamp As Integer
    Protected _eventChains As List(Of String)
    REM --- (end of YInputChain attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "InputChain"
      REM --- (YInputChain attributes initialization)
      _expectedNodes = EXPECTEDNODES_INVALID
      _detectedNodes = DETECTEDNODES_INVALID
      _loopbackTest = LOOPBACKTEST_INVALID
      _refreshRate = REFRESHRATE_INVALID
      _bitChain1 = BITCHAIN1_INVALID
      _bitChain2 = BITCHAIN2_INVALID
      _bitChain3 = BITCHAIN3_INVALID
      _bitChain4 = BITCHAIN4_INVALID
      _bitChain5 = BITCHAIN5_INVALID
      _bitChain6 = BITCHAIN6_INVALID
      _bitChain7 = BITCHAIN7_INVALID
      _watchdogPeriod = WATCHDOGPERIOD_INVALID
      _chainDiags = CHAINDIAGS_INVALID
      _valueCallbackInputChain = Nothing
      _prevPos = 0
      _eventPos = 0
      _eventStamp = 0
      _eventChains = New List(Of String)()
      REM --- (end of YInputChain attributes initialization)
    End Sub

    REM --- (YInputChain private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("expectedNodes") Then
        _expectedNodes = CInt(json_val.getLong("expectedNodes"))
      End If
      If json_val.has("detectedNodes") Then
        _detectedNodes = CInt(json_val.getLong("detectedNodes"))
      End If
      If json_val.has("loopbackTest") Then
        If (json_val.getInt("loopbackTest") > 0) Then _loopbackTest = 1 Else _loopbackTest = 0
      End If
      If json_val.has("refreshRate") Then
        _refreshRate = CInt(json_val.getLong("refreshRate"))
      End If
      If json_val.has("bitChain1") Then
        _bitChain1 = json_val.getString("bitChain1")
      End If
      If json_val.has("bitChain2") Then
        _bitChain2 = json_val.getString("bitChain2")
      End If
      If json_val.has("bitChain3") Then
        _bitChain3 = json_val.getString("bitChain3")
      End If
      If json_val.has("bitChain4") Then
        _bitChain4 = json_val.getString("bitChain4")
      End If
      If json_val.has("bitChain5") Then
        _bitChain5 = json_val.getString("bitChain5")
      End If
      If json_val.has("bitChain6") Then
        _bitChain6 = json_val.getString("bitChain6")
      End If
      If json_val.has("bitChain7") Then
        _bitChain7 = json_val.getString("bitChain7")
      End If
      If json_val.has("watchdogPeriod") Then
        _watchdogPeriod = CInt(json_val.getLong("watchdogPeriod"))
      End If
      If json_val.has("chainDiags") Then
        _chainDiags = CInt(json_val.getLong("chainDiags"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YInputChain private methods declaration)

    REM --- (YInputChain public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the number of nodes expected in the chain.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of nodes expected in the chain
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.EXPECTEDNODES_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_expectedNodes() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return EXPECTEDNODES_INVALID
        End If
      End If
      res = Me._expectedNodes
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of nodes expected in the chain.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the number of nodes expected in the chain
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
    Public Function set_expectedNodes(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("expectedNodes", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of nodes detected in the chain.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of nodes detected in the chain
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.DETECTEDNODES_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_detectedNodes() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DETECTEDNODES_INVALID
        End If
      End If
      res = Me._detectedNodes
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the activation state of the exhaustive chain connectivity test.
    ''' <para>
    '''   The connectivity test requires a cable connecting the end of the chain
    '''   to the loopback test connector.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YInputChain.LOOPBACKTEST_OFF</c> or <c>YInputChain.LOOPBACKTEST_ON</c>, according to the
    '''   activation state of the exhaustive chain connectivity test
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.LOOPBACKTEST_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_loopbackTest() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LOOPBACKTEST_INVALID
        End If
      End If
      res = Me._loopbackTest
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the activation state of the exhaustive chain connectivity test.
    ''' <para>
    '''   The connectivity test requires a cable connecting the end of the chain
    '''   to the loopback test connector.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YInputChain.LOOPBACKTEST_OFF</c> or <c>YInputChain.LOOPBACKTEST_ON</c>, according to the
    '''   activation state of the exhaustive chain connectivity test
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
    Public Function set_loopbackTest(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("loopbackTest", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the desired refresh rate, measured in Hz.
    ''' <para>
    '''   The higher the refresh rate is set, the higher the
    '''   communication speed on the chain will be.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the desired refresh rate, measured in Hz
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.REFRESHRATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_refreshRate() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return REFRESHRATE_INVALID
        End If
      End If
      res = Me._refreshRate
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the desired refresh rate, measured in Hz.
    ''' <para>
    '''   The higher the refresh rate is set, the higher the
    '''   communication speed on the chain will be.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the desired refresh rate, measured in Hz
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
    Public Function set_refreshRate(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("refreshRate", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the state of input 1 for all nodes of the input chain,
    '''   as a hexadecimal string.
    ''' <para>
    '''   The node nearest to the controller
    '''   is the lowest bit of the result.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the state of input 1 for all nodes of the input chain,
    '''   as a hexadecimal string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.BITCHAIN1_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bitChain1() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BITCHAIN1_INVALID
        End If
      End If
      res = Me._bitChain1
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of input 2 for all nodes of the input chain,
    '''   as a hexadecimal string.
    ''' <para>
    '''   The node nearest to the controller
    '''   is the lowest bit of the result.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the state of input 2 for all nodes of the input chain,
    '''   as a hexadecimal string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.BITCHAIN2_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bitChain2() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BITCHAIN2_INVALID
        End If
      End If
      res = Me._bitChain2
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of input 3 for all nodes of the input chain,
    '''   as a hexadecimal string.
    ''' <para>
    '''   The node nearest to the controller
    '''   is the lowest bit of the result.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the state of input 3 for all nodes of the input chain,
    '''   as a hexadecimal string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.BITCHAIN3_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bitChain3() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BITCHAIN3_INVALID
        End If
      End If
      res = Me._bitChain3
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of input 4 for all nodes of the input chain,
    '''   as a hexadecimal string.
    ''' <para>
    '''   The node nearest to the controller
    '''   is the lowest bit of the result.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the state of input 4 for all nodes of the input chain,
    '''   as a hexadecimal string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.BITCHAIN4_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bitChain4() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BITCHAIN4_INVALID
        End If
      End If
      res = Me._bitChain4
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of input 5 for all nodes of the input chain,
    '''   as a hexadecimal string.
    ''' <para>
    '''   The node nearest to the controller
    '''   is the lowest bit of the result.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the state of input 5 for all nodes of the input chain,
    '''   as a hexadecimal string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.BITCHAIN5_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bitChain5() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BITCHAIN5_INVALID
        End If
      End If
      res = Me._bitChain5
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of input 6 for all nodes of the input chain,
    '''   as a hexadecimal string.
    ''' <para>
    '''   The node nearest to the controller
    '''   is the lowest bit of the result.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the state of input 6 for all nodes of the input chain,
    '''   as a hexadecimal string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.BITCHAIN6_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bitChain6() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BITCHAIN6_INVALID
        End If
      End If
      res = Me._bitChain6
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of input 7 for all nodes of the input chain,
    '''   as a hexadecimal string.
    ''' <para>
    '''   The node nearest to the controller
    '''   is the lowest bit of the result.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the state of input 7 for all nodes of the input chain,
    '''   as a hexadecimal string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.BITCHAIN7_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bitChain7() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BITCHAIN7_INVALID
        End If
      End If
      res = Me._bitChain7
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the wait time in seconds before triggering an inactivity
    '''   timeout error.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the wait time in seconds before triggering an inactivity
    '''   timeout error
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.WATCHDOGPERIOD_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_watchdogPeriod() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return WATCHDOGPERIOD_INVALID
        End If
      End If
      res = Me._watchdogPeriod
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the wait time in seconds before triggering an inactivity
    '''   timeout error.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method
    '''   of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the wait time in seconds before triggering an inactivity
    '''   timeout error
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
    Public Function set_watchdogPeriod(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("watchdogPeriod", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the controller state diagnostics.
    ''' <para>
    '''   Bit 0 indicates a chain length
    '''   error, bit 1 indicates an inactivity timeout and bit 2 indicates
    '''   a loopback test failure.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the controller state diagnostics
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YInputChain.CHAINDIAGS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_chainDiags() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CHAINDIAGS_INVALID
        End If
      End If
      res = Me._chainDiags
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a digital input chain for a given identifier.
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
    '''   This function does not require that the digital input chain is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YInputChain.isOnline()</c> to test if the digital input chain is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a digital input chain by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the digital input chain, for instance
    '''   <c>MyDevice.inputChain</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YInputChain</c> object allowing you to drive the digital input chain.
    ''' </returns>
    '''/
    Public Shared Function FindInputChain(func As String) As YInputChain
      Dim obj As YInputChain
      obj = CType(YFunction._FindFromCache("InputChain", func), YInputChain)
      If ((obj Is Nothing)) Then
        obj = New YInputChain(func)
        YFunction._AddToCache("InputChain", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YInputChainValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackInputChain = callback
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
      If (Not (Me._valueCallbackInputChain Is Nothing)) Then
        Me._valueCallbackInputChain(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Resets the application watchdog countdown.
    ''' <para>
    '''   If you have setup a non-zero <c>watchdogPeriod</c>, you should
    '''   call this function on a regular basis to prevent the application
    '''   inactivity error to be triggered.
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
    Public Overridable Function resetWatchdog() As Integer
      Return Me.set_watchdogPeriod(-1)
    End Function

    '''*
    ''' <summary>
    '''   Returns a string with last events observed on the digital input chain.
    ''' <para>
    '''   This method return only events that are still buffered in the device memory.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with last events observed (one per line).
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns  <c>YAPI.INVALID_STRING</c>.
    ''' </para>
    '''/
    Public Overridable Function get_lastEvents() As String
      Dim content As Byte() = New Byte(){}

      content = Me._download("events.txt")
      Return YAPI.DefaultEncoding.GetString(content)
    End Function

    '''*
    ''' <summary>
    '''   Registers a callback function to be called each time that an event is detected on the
    '''   input chain.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer.
    '''   The callback function should take four arguments:
    '''   the <c>YInputChain</c> object that emitted the event, the
    '''   UTC timestamp of the event, a character string describing
    '''   the type of event and a character string with the event data.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </param>
    '''/
    Public Overridable Function registerEventCallback(callback As YEventCallback) As Integer
      If (Not (callback Is Nothing)) Then
        Me.registerValueCallback(AddressOf yInternalEventCallback)
      Else
        Me.registerValueCallback(CType(Nothing, YInputChainValueCallback))
      End If
      REM // register user callback AFTER the internal pseudo-event,
      REM // to make sure we start with future events only
      Me._eventCallback = callback
      Return 0
    End Function

    Public Overridable Function _internalEventHandler(cbpos As String) As Integer
      Dim newPos As Integer = 0
      Dim url As String
      Dim content As Byte() = New Byte(){}
      Dim contentStr As String
      Dim eventArr As List(Of String) = New List(Of String)()
      Dim arrLen As Integer = 0
      Dim lenStr As String
      Dim arrPos As Integer = 0
      Dim eventStr As String
      Dim eventLen As Integer = 0
      Dim hexStamp As String
      Dim typePos As Integer = 0
      Dim dataPos As Integer = 0
      Dim evtStamp As Integer = 0
      Dim evtType As String
      Dim evtData As String
      Dim evtChange As String
      Dim chainIdx As Integer = 0
      newPos = YAPI._atoi(cbpos)
      If (newPos < Me._prevPos) Then
        Me._eventPos = 0
      End If
      Me._prevPos = newPos
      If (newPos < Me._eventPos) Then
        Return YAPI.SUCCESS
      End If
      If (Not (Not (Me._eventCallback Is Nothing))) Then
        REM // first simulated event, use it to initialize reference values
        Me._eventPos = newPos
        Me._eventChains.Clear()
        Me._eventChains.Add(Me.get_bitChain1())
        Me._eventChains.Add(Me.get_bitChain2())
        Me._eventChains.Add(Me.get_bitChain3())
        Me._eventChains.Add(Me.get_bitChain4())
        Me._eventChains.Add(Me.get_bitChain5())
        Me._eventChains.Add(Me.get_bitChain6())
        Me._eventChains.Add(Me.get_bitChain7())
        Return YAPI.SUCCESS
      End If
      url = "events.txt?pos=" + Convert.ToString(Me._eventPos)

      content = Me._download(url)
      contentStr = YAPI.DefaultEncoding.GetString(content)
      eventArr = New List(Of String)(contentStr.Split(vbLf.ToCharArray()))
      arrLen = eventArr.Count
      If Not(arrLen > 0) Then
        me._throw( YAPI.IO_ERROR,  "fail to download events")
        return YAPI.IO_ERROR
      end if
      REM // last element of array is the new position preceeded by '@'
      arrLen = arrLen - 1
      lenStr = eventArr(arrLen)
      lenStr = (lenStr).Substring( 1, (lenStr).Length-1)
      REM // update processed event position pointer
      Me._eventPos = YAPI._atoi(lenStr)
      REM // now generate callbacks for each event received
      arrPos = 0
      While (arrPos < arrLen)
        eventStr = eventArr(arrPos)
        eventLen = (eventStr).Length
        If (eventLen >= 1) Then
          hexStamp = (eventStr).Substring( 0, 8)
          evtStamp = Convert.ToInt32(hexStamp, 16)
          typePos = eventStr.IndexOf(":")+1
          If ((evtStamp >= Me._eventStamp) AndAlso (typePos > 8)) Then
            Me._eventStamp = evtStamp
            dataPos = eventStr.IndexOf("=")+1
            evtType = (eventStr).Substring( typePos, 1)
            evtData = ""
            evtChange = ""
            If (dataPos > 10) Then
              evtData = (eventStr).Substring( dataPos, (eventStr).Length-dataPos)
              If ("1234567".IndexOf(evtType) >= 0) Then
                chainIdx = YAPI._atoi(evtType) - 1
                evtChange = Me._strXor(evtData, Me._eventChains(chainIdx))
                Me._eventChains( chainIdx) = evtData
              End If
            End If
            Me._eventCallback(Me, evtStamp, evtType, evtData, evtChange)
          End If
        End If
        arrPos = arrPos + 1
      End While
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function _strXor(a As String, b As String) As String
      Dim lenA As Integer = 0
      Dim lenB As Integer = 0
      Dim res As String
      Dim idx As Integer = 0
      Dim digitA As Integer = 0
      Dim digitB As Integer = 0
      REM // make sure the result has the same length as first argument
      lenA = (a).Length
      lenB = (b).Length
      If (lenA > lenB) Then
        res = (a).Substring( 0, lenA-lenB)
        a = (a).Substring( lenA-lenB, lenB)
        lenA = lenB
      Else
        res = ""
        b = (b).Substring( lenA-lenB, lenA)
      End If
      REM // scan strings and compare digit by digit
      idx = 0
      While (idx < lenA)
        digitA = Convert.ToInt32((a).Substring( idx, 1), 16)
        digitB = Convert.ToInt32((b).Substring( idx, 1), 16)
        res = "" +  res + "" + (((digitA) Xor (digitB))).ToString("x")
        idx = idx + 1
      End While
      Return res
    End Function

    Public Overridable Function hex2array(hexstr As String) As List(Of Integer)
      Dim hexlen As Integer = 0
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim idx As Integer = 0
      Dim digit As Integer = 0
      hexlen = (hexstr).Length
      res.Clear()

      idx = hexlen
      While (idx > 0)
        idx = idx - 1
        digit = Convert.ToInt32((hexstr).Substring( idx, 1), 16)
        res.Add(((digit) And (1)))
        res.Add(((((digit) >> (1))) And (1)))
        res.Add(((((digit) >> (2))) And (1)))
        res.Add(((((digit) >> (3))) And (1)))
      End While

      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of digital input chains started using <c>yFirstInputChain()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned digital input chains order.
    '''   If you want to find a specific a digital input chain, use <c>InputChain.findInputChain()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YInputChain</c> object, corresponding to
    '''   a digital input chain currently online, or a <c>Nothing</c> pointer
    '''   if there are no more digital input chains to enumerate.
    ''' </returns>
    '''/
    Public Function nextInputChain() As YInputChain
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YInputChain.FindInputChain(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of digital input chains currently accessible.
    ''' <para>
    '''   Use the method <c>YInputChain.nextInputChain()</c> to iterate on
    '''   next digital input chains.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YInputChain</c> object, corresponding to
    '''   the first digital input chain currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstInputChain() As YInputChain
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("InputChain", 0, p, size, neededsize, errmsg)
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
      Return YInputChain.FindInputChain(serial + "." + funcId)
    End Function

    REM --- (end of YInputChain public methods declaration)

  End Class

  REM --- (YInputChain functions)

  '''*
  ''' <summary>
  '''   Retrieves a digital input chain for a given identifier.
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
  '''   This function does not require that the digital input chain is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YInputChain.isOnline()</c> to test if the digital input chain is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a digital input chain by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the digital input chain, for instance
  '''   <c>MyDevice.inputChain</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YInputChain</c> object allowing you to drive the digital input chain.
  ''' </returns>
  '''/
  Public Function yFindInputChain(ByVal func As String) As YInputChain
    Return YInputChain.FindInputChain(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of digital input chains currently accessible.
  ''' <para>
  '''   Use the method <c>YInputChain.nextInputChain()</c> to iterate on
  '''   next digital input chains.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YInputChain</c> object, corresponding to
  '''   the first digital input chain currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstInputChain() As YInputChain
    Return YInputChain.FirstInputChain()
  End Function


  REM --- (end of YInputChain functions)

End Module
