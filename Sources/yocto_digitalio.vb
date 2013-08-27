'*********************************************************************
'*
'* $Id: pic24config.php 12323 2013-08-13 15:09:18Z mvuilleu $
'*
'* Implements yFindDigitalIO(), the high-level API for DigitalIO functions
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

Module yocto_digitalio

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YDigitalIO definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YDigitalIO, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PORTSTATE_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_PORTDIRECTION_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_PORTOPENDRAIN_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_PORTSIZE_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_OUTPUTVOLTAGE_USB_5V = 0
  Public Const Y_OUTPUTVOLTAGE_USB_3V3 = 1
  Public Const Y_OUTPUTVOLTAGE_EXT_V = 2
  Public Const Y_OUTPUTVOLTAGE_INVALID = -1

  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING


  REM --- (end of YDigitalIO definitions)

  REM --- (YDigitalIO implementation)

  Private _DigitalIOCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   .
  ''' <para>
  '''   ...
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDigitalIO
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const PORTSTATE_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const PORTDIRECTION_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const PORTOPENDRAIN_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const PORTSIZE_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const OUTPUTVOLTAGE_USB_5V = 0
    Public Const OUTPUTVOLTAGE_USB_3V3 = 1
    Public Const OUTPUTVOLTAGE_EXT_V = 2
    Public Const OUTPUTVOLTAGE_INVALID = -1

    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _portState As Long
    Protected _portDirection As Long
    Protected _portOpenDrain As Long
    Protected _portSize As Long
    Protected _outputVoltage As Long
    Protected _command As String

    Public Sub New(ByVal func As String)
      MyBase.new("DigitalIO", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _portState = Y_PORTSTATE_INVALID
      _portDirection = Y_PORTDIRECTION_INVALID
      _portOpenDrain = Y_PORTOPENDRAIN_INVALID
      _portSize = Y_PORTSIZE_INVALID
      _outputVoltage = Y_OUTPUTVOLTAGE_INVALID
      _command = Y_COMMAND_INVALID
    End Sub

    Protected Overrides Function _parse(ByRef j As TJSONRECORD) As Integer
      Dim member As TJSONRECORD
      Dim i As Integer
      If (j.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then
        Return -1
      End If
      For i = 0 To j.membercount - 1
        member = j.members(i)
        If (member.name = "logicalName") Then
          _logicalName = member.svalue
        ElseIf (member.name = "advertisedValue") Then
          _advertisedValue = member.svalue
        ElseIf (member.name = "portState") Then
          _portState = CLng(member.ivalue)
        ElseIf (member.name = "portDirection") Then
          _portDirection = CLng(member.ivalue)
        ElseIf (member.name = "portOpenDrain") Then
          _portOpenDrain = CLng(member.ivalue)
        ElseIf (member.name = "portSize") Then
          _portSize = CLng(member.ivalue)
        ElseIf (member.name = "outputVoltage") Then
          _outputVoltage = CLng(member.ivalue)
        ElseIf (member.name = "command") Then
          _command = member.svalue
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the digital IO port.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the digital IO port
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LOGICALNAME_INVALID
        End If
      End If
      Return _logicalName
    End Function

    '''*
    ''' <summary>
    '''   Changes the logical name of the digital IO port.
    ''' <para>
    '''   You can use <c>yCheckLogicalName()</c>
    '''   prior to this call to make sure that your parameter is valid.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the logical name of the digital IO port
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
    Public Function set_logicalName(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("logicalName", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of the digital IO port (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the digital IO port (no more than 6 characters)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ADVERTISEDVALUE_INVALID
        End If
      End If
      Return _advertisedValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the digital IO port state: bit 0 represents input 0, and so on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the digital IO port state: bit 0 represents input 0, and so on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PORTSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portState() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PORTSTATE_INVALID
        End If
      End If
      Return CType(_portState,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the digital IO port state: bit 0 represents input 0, and so on.
    ''' <para>
    '''   This function has no effect
    '''   on bits configured as input in <c>portDirection</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the digital IO port state: bit 0 represents input 0, and so on
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
    Public Function set_portState(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("portState", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the IO direction of all bits of the port: 0 makes a bit an input, 1 makes it an output.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the IO direction of all bits of the port: 0 makes a bit an input, 1
    '''   makes it an output
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PORTDIRECTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portDirection() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PORTDIRECTION_INVALID
        End If
      End If
      Return CType(_portDirection,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the IO direction of all bits of the port: 0 makes a bit an input, 1 makes it an output.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method  to make sure the setting will be kept after a reboot.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the IO direction of all bits of the port: 0 makes a bit an input, 1
    '''   makes it an output
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
    Public Function set_portDirection(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("portDirection", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the electrical interface for each bit of the port.
    ''' <para>
    '''   0 makes a bit a regular input/output, 1 makes
    '''   it an open-drain (open-collector) input/output.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the electrical interface for each bit of the port
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PORTOPENDRAIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portOpenDrain() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PORTOPENDRAIN_INVALID
        End If
      End If
      Return CType(_portOpenDrain,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the electrical interface for each bit of the port.
    ''' <para>
    '''   0 makes a bit a regular input/output, 1 makes
    '''   it an open-drain (open-collector) input/output. Remember to call the
    '''   <c>saveToFlash()</c> method  to make sure the setting will be kept after a reboot.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   Returns the number of bits implemented in the I/O port.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of bits implemented in the I/O port
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PORTSIZE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_portSize() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PORTSIZE_INVALID
        End If
      End If
      Return CType(_portSize,Integer)
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
    '''   a value among <c>Y_OUTPUTVOLTAGE_USB_5V</c>, <c>Y_OUTPUTVOLTAGE_USB_3V3</c> and
    '''   <c>Y_OUTPUTVOLTAGE_EXT_V</c> corresponding to the voltage source used to drive output bits
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OUTPUTVOLTAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_outputVoltage() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_OUTPUTVOLTAGE_INVALID
        End If
      End If
      Return CType(_outputVoltage,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the voltage source used to drive output bits.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method  to make sure the setting will be kept after a reboot.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_OUTPUTVOLTAGE_USB_5V</c>, <c>Y_OUTPUTVOLTAGE_USB_3V3</c> and
    '''   <c>Y_OUTPUTVOLTAGE_EXT_V</c> corresponding to the voltage source used to drive output bits
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
    Public Function set_outputVoltage(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("outputVoltage", rest_val)
    End Function

    Public Function get_command() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_COMMAND_INVALID
        End If
      End If
      Return _command
    End Function

    Public Function set_command(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("command", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Set a single bit of the I/O port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit is index 0
    ''' </param>
    ''' <param name="bitval">
    '''   the value of the bit (1 or 0)
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function set_bitState(bitno as integer, bitval as integer) as integer
        if not(bitval >= 0) then
me._throw( YAPI.INVALID_ARGUMENT, "invalid bitval")
 return  YAPI.INVALID_ARGUMENT
end if

        if not(bitval <= 1) then
me._throw( YAPI.INVALID_ARGUMENT, "invalid bitval")
 return  YAPI.INVALID_ARGUMENT
end if

        Return Me.set_command(""+Chr(82+bitval)+""+ Convert.ToString( bitno))
        
     end function

    '''*
    ''' <summary>
    '''   Returns the value of a single bit of the I/O port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit is index 0
    ''' </param>
    ''' <returns>
    '''   the bit value (0 or 1)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function get_bitState(bitno as integer) as integer
        dim  portVal as integer
        portVal = Me.get_portState()
        Return ((((portVal) >> (bitno))) and (1))
        
     end function

    '''*
    ''' <summary>
    '''   Revert a single bit of the I/O port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit is index 0
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function toggle_bitState(bitno as integer) as integer
        Return Me.set_command("T"+ Convert.ToString( bitno))
        
     end function

    '''*
    ''' <summary>
    '''   Change  the direction of a single bit from the I/O port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit is index 0
    ''' </param>
    ''' <param name="bitdirection">
    '''   direction to set, 0 makes the bit an input, 1 makes it an output.
    '''   Remember to call the   <c>saveToFlash()</c> method to make sure the setting will be kept after a reboot.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function set_bitDirection(bitno as integer, bitdirection as integer) as integer
        if not(bitdirection >= 0) then
me._throw( YAPI.INVALID_ARGUMENT, "invalid direction")
 return  YAPI.INVALID_ARGUMENT
end if

        if not(bitdirection <= 1) then
me._throw( YAPI.INVALID_ARGUMENT, "invalid direction")
 return  YAPI.INVALID_ARGUMENT
end if

        Return Me.set_command(""+Chr(73+6*bitdirection)+""+ Convert.ToString( bitno))
        
     end function

    '''*
    ''' <summary>
    '''   Change  the direction of a single bit from the I/O port (0 means the bit is an input, 1  an output).
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit is index 0
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function get_bitDirection(bitno as integer) as integer
        dim  portDir as integer
        portDir = Me.get_portDirection()
        Return ((((portDir) >> (bitno))) and (1))
        
     end function

    '''*
    ''' <summary>
    '''   Change  the electrical interface of a single bit from the I/O port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit is index 0
    ''' </param>
    ''' <param name="opendrain">
    '''   value to set, 0 makes a bit a regular input/output, 1 makes
    '''   it an open-drain (open-collector) input/output. Remember to call the
    '''   <c>saveToFlash()</c> method to make sure the setting will be kept after a reboot.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function set_bitOpenDrain(bitno as integer, opendrain as integer) as integer
        if not(opendrain >= 0) then
me._throw( YAPI.INVALID_ARGUMENT, "invalid state")
 return  YAPI.INVALID_ARGUMENT
end if

        if not(opendrain <= 1) then
me._throw( YAPI.INVALID_ARGUMENT, "invalid state")
 return  YAPI.INVALID_ARGUMENT
end if

        Return Me.set_command(""+Chr(100-32*opendrain)+""+ Convert.ToString( bitno))
        
     end function

    '''*
    ''' <summary>
    '''   Returns the type of electrical interface of a single bit from the I/O port.
    ''' <para>
    '''   (0 means the bit is an input, 1  an output).
    ''' </para>
    ''' </summary>
    ''' <param name="bitno">
    '''   the bit number; lowest bit is index 0
    ''' </param>
    ''' <returns>
    '''   0 means the a bit is a regular input/output, 1means the b it an open-drain (open-collector) input/output.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function get_bitOpenDrain(bitno as integer) as integer
        dim  portOpenDrain as integer
        portOpenDrain = Me.get_portOpenDrain()
        Return ((((portOpenDrain) >> (bitno))) and (1))
        
     end function


    '''*
    ''' <summary>
    '''   Continues the enumeration of digital IO port started using <c>yFirstDigitalIO()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDigitalIO</c> object, corresponding to
    '''   a digital IO port currently online, or a <c>null</c> pointer
    '''   if there are no more digital IO port to enumerate.
    ''' </returns>
    '''/
    Public Function nextDigitalIO() as YDigitalIO
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindDigitalIO(hwid)
    End Function

    '''*
    ''' <summary>
    '''   comment from .
    ''' <para>
    '''   yc definition
    ''' </para>
    ''' </summary>
    '''/
  Public Overloads Sub registerValueCallback(ByVal callback As UpdateCallback)
   If (callback IsNot Nothing) Then
     registerFuncCallback(Me)
   Else
     unregisterFuncCallback(Me)
   End If
   _callback = callback
  End Sub

  Public Sub set_callback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Sub setCallback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Overrides Sub advertiseValue(ByVal value As String)
    If (_callback IsNot Nothing) Then _callback(Me, value)
  End Sub


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
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the digital IO port
    ''' </param>
    ''' <returns>
    '''   a <c>YDigitalIO</c> object allowing you to drive the digital IO port.
    ''' </returns>
    '''/
    Public Shared Function FindDigitalIO(ByVal func As String) As YDigitalIO
      Dim res As YDigitalIO
      If (_DigitalIOCache.ContainsKey(func)) Then
        Return CType(_DigitalIOCache(func), YDigitalIO)
      End If
      res = New YDigitalIO(func)
      _DigitalIOCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of digital IO port currently accessible.
    ''' <para>
    '''   Use the method <c>YDigitalIO.nextDigitalIO()</c> to iterate on
    '''   next digital IO port.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDigitalIO</c> object, corresponding to
    '''   the first digital IO port currently online, or a <c>null</c> pointer
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

    REM --- (end of YDigitalIO implementation)

  End Class

  REM --- (DigitalIO functions)

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
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the digital IO port
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
  '''   Starts the enumeration of digital IO port currently accessible.
  ''' <para>
  '''   Use the method <c>YDigitalIO.nextDigitalIO()</c> to iterate on
  '''   next digital IO port.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YDigitalIO</c> object, corresponding to
  '''   the first digital IO port currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstDigitalIO() As YDigitalIO
    Return YDigitalIO.FirstDigitalIO()
  End Function

  Private Sub _DigitalIOCleanup()
  End Sub


  REM --- (end of DigitalIO functions)

End Module
