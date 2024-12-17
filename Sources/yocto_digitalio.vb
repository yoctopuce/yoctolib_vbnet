' ********************************************************************
'
'  $Id: yocto_digitalio.vb 63328 2024-11-13 09:35:22Z seb $
'
'  Implements yFindDigitalIO(), the high-level API for DigitalIO functions
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

Module yocto_digitalio

    REM --- (YDigitalIO return codes)
    REM --- (end of YDigitalIO return codes)
    REM --- (YDigitalIO dlldef)
    REM --- (end of YDigitalIO dlldef)
   REM --- (YDigitalIO yapiwrapper)
   REM --- (end of YDigitalIO yapiwrapper)
  REM --- (YDigitalIO globals)

  Public Const Y_PORTSTATE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PORTDIRECTION_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PORTOPENDRAIN_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PORTPOLARITY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PORTDIAGS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PORTSIZE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_OUTPUTVOLTAGE_USB_5V As Integer = 0
  Public Const Y_OUTPUTVOLTAGE_USB_3V As Integer = 1
  Public Const Y_OUTPUTVOLTAGE_EXT_V As Integer = 2
  Public Const Y_OUTPUTVOLTAGE_INVALID As Integer = -1
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YDigitalIOValueCallback(ByVal func As YDigitalIO, ByVal value As String)
  Public Delegate Sub YDigitalIOTimedReportCallback(ByVal func As YDigitalIO, ByVal measure As YMeasure)
  REM --- (end of YDigitalIO globals)

  REM --- (YDigitalIO class start)

  '''*
  ''' <summary>
  '''   The <c>YDigitalIO</c> class allows you drive a Yoctopuce digital input/output port.
  ''' <para>
  '''   It can be used to set up the direction of each channel, to read the state of each channel
  '''   and to switch the state of each channel configures as an output.
  '''   You can work on all channels at once, or one by one. Most functions
  '''   use a binary representation for channels where bit 0 matches channel #0 , bit 1 matches channel
  '''   #1 and so on. If you are not familiar with numbers binary representation, you will find more
  '''   information here: <c>https://en.wikipedia.org/wiki/Binary_number#Representation</c>. It is also possible
  '''   to automatically generate short pulses of a determined duration. Electrical behavior
  '''   of each I/O can be modified (open drain and reverse polarity).
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDigitalIO
    Inherits YFunction
    REM --- (end of YDigitalIO class start)

    REM --- (YDigitalIO definitions)
    Public Const PORTSTATE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PORTDIRECTION_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PORTOPENDRAIN_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PORTPOLARITY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PORTDIAGS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PORTSIZE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const OUTPUTVOLTAGE_USB_5V As Integer = 0
    Public Const OUTPUTVOLTAGE_USB_3V As Integer = 1
    Public Const OUTPUTVOLTAGE_EXT_V As Integer = 2
    Public Const OUTPUTVOLTAGE_INVALID As Integer = -1
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YDigitalIO definitions)

    REM --- (YDigitalIO attributes declaration)
    Protected _portState As Integer
    Protected _portDirection As Integer
    Protected _portOpenDrain As Integer
    Protected _portPolarity As Integer
    Protected _portDiags As Integer
    Protected _portSize As Integer
    Protected _outputVoltage As Integer
    Protected _command As String
    Protected _valueCallbackDigitalIO As YDigitalIOValueCallback
    REM --- (end of YDigitalIO attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "DigitalIO"
      REM --- (YDigitalIO attributes initialization)
      _portState = PORTSTATE_INVALID
      _portDirection = PORTDIRECTION_INVALID
      _portOpenDrain = PORTOPENDRAIN_INVALID
      _portPolarity = PORTPOLARITY_INVALID
      _portDiags = PORTDIAGS_INVALID
      _portSize = PORTSIZE_INVALID
      _outputVoltage = OUTPUTVOLTAGE_INVALID
      _command = COMMAND_INVALID
      _valueCallbackDigitalIO = Nothing
      REM --- (end of YDigitalIO attributes initialization)
    End Sub

    REM --- (YDigitalIO private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("portState") Then
        _portState = CInt(json_val.getLong("portState"))
      End If
      If json_val.has("portDirection") Then
        _portDirection = CInt(json_val.getLong("portDirection"))
      End If
      If json_val.has("portOpenDrain") Then
        _portOpenDrain = CInt(json_val.getLong("portOpenDrain"))
      End If
      If json_val.has("portPolarity") Then
        _portPolarity = CInt(json_val.getLong("portPolarity"))
      End If
      If json_val.has("portDiags") Then
        _portDiags = CInt(json_val.getLong("portDiags"))
      End If
      If json_val.has("portSize") Then
        _portSize = CInt(json_val.getLong("portSize"))
      End If
      If json_val.has("outputVoltage") Then
        _outputVoltage = CInt(json_val.getLong("outputVoltage"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YDigitalIO private methods declaration)

    REM --- (YDigitalIO public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the digital IO port state as an integer with each bit
    '''   representing a channel.
    ''' <para>
    '''   value 0 = <c>0b00000000</c> -> all channels are OFF
    '''   value 1 = <c>0b00000001</c> -> channel #0 is ON
    '''   value 2 = <c>0b00000010</c> -> channel #1 is ON
    '''   value 3 = <c>0b00000011</c> -> channels #0 and #1 are ON
    '''   value 4 = <c>0b00000100</c> -> channel #2 is ON
    '''   and so on...
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the digital IO port state as an integer with each bit
    '''   representing a channel
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDigitalIO.PORTSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portState() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PORTSTATE_INVALID
        End If
      End If
      res = Me._portState
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the state of all digital IO port's channels at once: the parameter
    '''   is an integer where each bit represents a channel, with bit 0 matching channel #0.
    ''' <para>
    '''   To set all channels to  0 -> <c>0b00000000</c> -> parameter = 0
    '''   To set channel #0 to 1 -> <c>0b00000001</c> -> parameter =  1
    '''   To set channel #1 to  1 -> <c>0b00000010</c> -> parameter = 2
    '''   To set channel #0 and #1 -> <c>0b00000011</c> -> parameter =  3
    '''   To set channel #2 to 1 -> <c>0b00000100</c> -> parameter =  4
    '''   an so on....
    '''   Only channels configured as outputs will be affecter, according to the value
    '''   configured using <c>set_portDirection</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the state of all digital IO port's channels at once: the parameter
    '''   is an integer where each bit represents a channel, with bit 0 matching channel #0
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
    Public Function set_portState(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("portState", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the I/O direction of all channels of the port (bitmap): 0 makes a bit an input, 1 makes it an output.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the I/O direction of all channels of the port (bitmap): 0 makes a bit
    '''   an input, 1 makes it an output
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDigitalIO.PORTDIRECTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portDirection() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PORTDIRECTION_INVALID
        End If
      End If
      res = Me._portDirection
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the I/O direction of all channels of the port (bitmap): 0 makes a bit an input, 1 makes it an output.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method  to make sure the setting is kept after a reboot.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the I/O direction of all channels of the port (bitmap): 0 makes a bit
    '''   an input, 1 makes it an output
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
    Public Function set_portDirection(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("portDirection", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the electrical interface for each bit of the port.
    ''' <para>
    '''   For each bit set to 0  the matching I/O works in the regular,
    '''   intuitive way, for each bit set to 1, the I/O works in reverse mode.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the electrical interface for each bit of the port
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDigitalIO.PORTOPENDRAIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portOpenDrain() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PORTOPENDRAIN_INVALID
        End If
      End If
      res = Me._portOpenDrain
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the electrical interface for each bit of the port.
    ''' <para>
    '''   0 makes a bit a regular input/output, 1 makes
    '''   it an open-drain (open-collector) input/output. Remember to call the
    '''   <c>saveToFlash()</c> method  to make sure the setting is kept after a reboot.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the electrical interface for each bit of the port
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
    Public Function set_portOpenDrain(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("portOpenDrain", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the polarity of all the bits of the port.
    ''' <para>
    '''   For each bit set to 0, the matching I/O works the regular,
    '''   intuitive way; for each bit set to 1, the I/O works in reverse mode.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the polarity of all the bits of the port
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDigitalIO.PORTPOLARITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portPolarity() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PORTPOLARITY_INVALID
        End If
      End If
      res = Me._portPolarity
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the polarity of all the bits of the port: For each bit set to 0, the matching I/O works the regular,
    '''   intuitive way; for each bit set to 1, the I/O works in reverse mode.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method  to make sure the setting will be kept after a reboot.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the polarity of all the bits of the port: For each bit set to 0, the
    '''   matching I/O works the regular,
    '''   intuitive way; for each bit set to 1, the I/O works in reverse mode
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
    Public Function set_portPolarity(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("portPolarity", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the port state diagnostics.
    ''' <para>
    '''   Bit 0 indicates a shortcut on output 0, etc.
    '''   Bit 8 indicates a power failure, and bit 9 signals overheating (overcurrent).
    '''   During normal use, all diagnostic bits should stay clear.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the port state diagnostics
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDigitalIO.PORTDIAGS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portDiags() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PORTDIAGS_INVALID
        End If
      End If
      res = Me._portDiags
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of bits (i.e.
    ''' <para>
    '''   channels)implemented in the I/O port.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of bits (i.e
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDigitalIO.PORTSIZE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portSize() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PORTSIZE_INVALID
        End If
      End If
      res = Me._portSize
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the voltage source used to drive output bits.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YDigitalIO.OUTPUTVOLTAGE_USB_5V</c>, <c>YDigitalIO.OUTPUTVOLTAGE_USB_3V</c> and
    '''   <c>YDigitalIO.OUTPUTVOLTAGE_EXT_V</c> corresponding to the voltage source used to drive output bits
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDigitalIO.OUTPUTVOLTAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_outputVoltage() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return OUTPUTVOLTAGE_INVALID
        End If
      End If
      res = Me._outputVoltage
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the voltage source used to drive output bits.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method  to make sure the setting is kept after a reboot.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YDigitalIO.OUTPUTVOLTAGE_USB_5V</c>, <c>YDigitalIO.OUTPUTVOLTAGE_USB_3V</c> and
    '''   <c>YDigitalIO.OUTPUTVOLTAGE_EXT_V</c> corresponding to the voltage source used to drive output bits
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
    Public Function set_outputVoltage(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("outputVoltage", rest_val)
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
    '''   Retrieves a digital IO port for a given identifier.
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
    '''   This function does not require that the digital IO port is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YDigitalIO.isOnline()</c> to test if the digital IO port is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a digital IO port by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the digital IO port, for instance
    '''   <c>YMINIIO0.digitalIO</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YDigitalIO</c> object allowing you to drive the digital IO port.
    ''' </returns>
    '''/
    Public Shared Function FindDigitalIO(func As String) As YDigitalIO
      Dim obj As YDigitalIO
      obj = CType(YFunction._FindFromCache("DigitalIO", func), YDigitalIO)
      If ((obj Is Nothing)) Then
        obj = New YDigitalIO(func)
        YFunction._AddToCache("DigitalIO", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YDigitalIOValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackDigitalIO = callback
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
      If (Not (Me._valueCallbackDigitalIO Is Nothing)) Then
        Me._valueCallbackDigitalIO(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Sets a single bit (i.e.
    ''' <para>
    '''   channel) of the I/O port.
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <param name="bitstate">
    '''   the state of the bit (1 or 0)
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_bitState(bitno As Integer, bitstate As Integer) As Integer
      If Not(bitstate >= 0) Then
        me._throw(YAPI.INVALID_ARGUMENT, "invalid bit state")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not(bitstate <= 1) Then
        me._throw(YAPI.INVALID_ARGUMENT, "invalid bit state")
        return YAPI.INVALID_ARGUMENT
      end if
      Return Me.set_command("" + Chr(82+bitstate) + "" + Convert.ToString(bitno))
    End Function

    '''*
    ''' <summary>
    '''   Returns the state of a single bit (i.e.
    ''' <para>
    '''   channel)  of the I/O port.
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <returns>
    '''   the bit state (0 or 1)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_bitState(bitno As Integer) As Integer
      Dim portVal As Integer = 0
      portVal = Me.get_portState()
      Return (((portVal >> bitno)) And (1))
    End Function

    '''*
    ''' <summary>
    '''   Reverts a single bit (i.e.
    ''' <para>
    '''   channel) of the I/O port.
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function toggle_bitState(bitno As Integer) As Integer
      Return Me.set_command("T" + Convert.ToString(bitno))
    End Function

    '''*
    ''' <summary>
    '''   Changes  the direction of a single bit (i.e.
    ''' <para>
    '''   channel) from the I/O port.
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <param name="bitdirection">
    '''   direction to set, 0 makes the bit an input, 1 makes it an output.
    '''   Remember to call the   <c>saveToFlash()</c> method to make sure the setting is kept after a reboot.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_bitDirection(bitno As Integer, bitdirection As Integer) As Integer
      If Not(bitdirection >= 0) Then
        me._throw(YAPI.INVALID_ARGUMENT, "invalid direction")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not(bitdirection <= 1) Then
        me._throw(YAPI.INVALID_ARGUMENT, "invalid direction")
        return YAPI.INVALID_ARGUMENT
      end if
      Return Me.set_command("" + Chr(73+6*bitdirection) + "" + Convert.ToString(bitno))
    End Function

    '''*
    ''' <summary>
    '''   Returns the direction of a single bit (i.e.
    ''' <para>
    '''   channel) from the I/O port (0 means the bit is an input, 1  an output).
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_bitDirection(bitno As Integer) As Integer
      Dim portDir As Integer = 0
      portDir = Me.get_portDirection()
      Return (((portDir >> bitno)) And (1))
    End Function

    '''*
    ''' <summary>
    '''   Changes the polarity of a single bit from the I/O port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0.
    ''' </param>
    ''' <param name="bitpolarity">
    '''   polarity to set, 0 makes the I/O work in regular mode, 1 makes the I/O  works in reverse mode.
    '''   Remember to call the   <c>saveToFlash()</c> method to make sure the setting is kept after a reboot.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_bitPolarity(bitno As Integer, bitpolarity As Integer) As Integer
      If Not(bitpolarity >= 0) Then
        me._throw(YAPI.INVALID_ARGUMENT, "invalid bit polarity")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not(bitpolarity <= 1) Then
        me._throw(YAPI.INVALID_ARGUMENT, "invalid bit polarity")
        return YAPI.INVALID_ARGUMENT
      end if
      Return Me.set_command("" + Chr(110+4*bitpolarity) + "" + Convert.ToString(bitno))
    End Function

    '''*
    ''' <summary>
    '''   Returns the polarity of a single bit from the I/O port (0 means the I/O works in regular mode, 1 means the I/O  works in reverse mode).
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_bitPolarity(bitno As Integer) As Integer
      Dim portPol As Integer = 0
      portPol = Me.get_portPolarity()
      Return (((portPol >> bitno)) And (1))
    End Function

    '''*
    ''' <summary>
    '''   Changes  the electrical interface of a single bit from the I/O port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <param name="opendrain">
    '''   0 makes a bit a regular input/output, 1 makes
    '''   it an open-drain (open-collector) input/output. Remember to call the
    '''   <c>saveToFlash()</c> method to make sure the setting is kept after a reboot.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_bitOpenDrain(bitno As Integer, opendrain As Integer) As Integer
      If Not(opendrain >= 0) Then
        me._throw(YAPI.INVALID_ARGUMENT, "invalid state")
        return YAPI.INVALID_ARGUMENT
      end if
      If Not(opendrain <= 1) Then
        me._throw(YAPI.INVALID_ARGUMENT, "invalid state")
        return YAPI.INVALID_ARGUMENT
      end if
      Return Me.set_command("" + Chr(100-32*opendrain) + "" + Convert.ToString(bitno))
    End Function

    '''*
    ''' <summary>
    '''   Returns the type of electrical interface of a single bit from the I/O port.
    ''' <para>
    '''   (0 means the bit is an input, 1  an output).
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <returns>
    '''   0 means the a bit is a regular input/output, 1 means the bit is an open-drain
    '''   (open-collector) input/output.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_bitOpenDrain(bitno As Integer) As Integer
      Dim portOpenDrain As Integer = 0
      portOpenDrain = Me.get_portOpenDrain()
      Return (((portOpenDrain >> bitno)) And (1))
    End Function

    '''*
    ''' <summary>
    '''   Triggers a pulse on a single bit for a specified duration.
    ''' <para>
    '''   The specified bit
    '''   will be turned to 1, and then back to 0 after the given duration.
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <param name="ms_duration">
    '''   desired pulse duration in milliseconds. Be aware that the device time
    '''   resolution is not guaranteed up to the millisecond.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function pulse(bitno As Integer, ms_duration As Integer) As Integer
      Return Me.set_command("Z" + Convert.ToString(bitno) + ",0," + Convert.ToString(ms_duration))
    End Function

    '''*
    ''' <summary>
    '''   Schedules a pulse on a single bit for a specified duration.
    ''' <para>
    '''   The specified bit
    '''   will be turned to 1, and then back to 0 after the given duration.
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit has index 0
    ''' </param>
    ''' <param name="ms_delay">
    '''   waiting time before the pulse, in milliseconds
    ''' </param>
    ''' <param name="ms_duration">
    '''   desired pulse duration in milliseconds. Be aware that the device time
    '''   resolution is not guaranteed up to the millisecond.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function delayedPulse(bitno As Integer, ms_delay As Integer, ms_duration As Integer) As Integer
      Return Me.set_command("Z" + Convert.ToString(bitno) + "," + Convert.ToString(ms_delay) + "," + Convert.ToString(ms_duration))
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of digital IO ports started using <c>yFirstDigitalIO()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned digital IO ports order.
    '''   If you want to find a specific a digital IO port, use <c>DigitalIO.findDigitalIO()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDigitalIO</c> object, corresponding to
    '''   a digital IO port currently online, or a <c>Nothing</c> pointer
    '''   if there are no more digital IO ports to enumerate.
    ''' </returns>
    '''/
    Public Function nextDigitalIO() As YDigitalIO
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YDigitalIO.FindDigitalIO(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of digital IO ports currently accessible.
    ''' <para>
    '''   Use the method <c>YDigitalIO.nextDigitalIO()</c> to iterate on
    '''   next digital IO ports.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDigitalIO</c> object, corresponding to
    '''   the first digital IO port currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstDigitalIO() As YDigitalIO
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("DigitalIO", 0, p, size, neededsize, errmsg)
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
      Return YDigitalIO.FindDigitalIO(serial + "." + funcId)
    End Function

    REM --- (end of YDigitalIO public methods declaration)

  End Class

  REM --- (YDigitalIO functions)

  '''*
  ''' <summary>
  '''   Retrieves a digital IO port for a given identifier.
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
  '''   This function does not require that the digital IO port is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YDigitalIO.isOnline()</c> to test if the digital IO port is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a digital IO port by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the digital IO port, for instance
  '''   <c>YMINIIO0.digitalIO</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YDigitalIO</c> object allowing you to drive the digital IO port.
  ''' </returns>
  '''/
  Public Function yFindDigitalIO(ByVal func As String) As YDigitalIO
    Return YDigitalIO.FindDigitalIO(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of digital IO ports currently accessible.
  ''' <para>
  '''   Use the method <c>YDigitalIO.nextDigitalIO()</c> to iterate on
  '''   next digital IO ports.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YDigitalIO</c> object, corresponding to
  '''   the first digital IO port currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstDigitalIO() As YDigitalIO
    Return YDigitalIO.FirstDigitalIO()
  End Function


  REM --- (end of YDigitalIO functions)

End Module
