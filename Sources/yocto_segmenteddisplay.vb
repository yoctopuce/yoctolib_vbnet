'*********************************************************************
'*
'* $Id: yocto_segmenteddisplay.vb 18762 2014-12-16 16:00:39Z seb $
'*
'* Implements yFindSegmentedDisplay(), the high-level API for SegmentedDisplay functions
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

Module yocto_segmenteddisplay

    REM --- (YSegmentedDisplay return codes)
    REM --- (end of YSegmentedDisplay return codes)
    REM --- (YSegmentedDisplay dlldef)
    REM --- (end of YSegmentedDisplay dlldef)
  REM --- (YSegmentedDisplay globals)

  Public Const Y_DISPLAYEDTEXT_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_DISPLAYMODE_DISCONNECTED As Integer = 0
  Public Const Y_DISPLAYMODE_MANUAL As Integer = 1
  Public Const Y_DISPLAYMODE_AUTO1 As Integer = 2
  Public Const Y_DISPLAYMODE_AUTO60 As Integer = 3
  Public Const Y_DISPLAYMODE_INVALID As Integer = -1
  Public Delegate Sub YSegmentedDisplayValueCallback(ByVal func As YSegmentedDisplay, ByVal value As String)
  Public Delegate Sub YSegmentedDisplayTimedReportCallback(ByVal func As YSegmentedDisplay, ByVal measure As YMeasure)
  REM --- (end of YSegmentedDisplay globals)

  REM --- (YSegmentedDisplay class start)

  '''*
  ''' <summary>
  '''   The SegmentedDisplay class allows you to drive segmented displays.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSegmentedDisplay
    Inherits YFunction
    REM --- (end of YSegmentedDisplay class start)

    REM --- (YSegmentedDisplay definitions)
    Public Const DISPLAYEDTEXT_INVALID As String = YAPI.INVALID_STRING
    Public Const DISPLAYMODE_DISCONNECTED As Integer = 0
    Public Const DISPLAYMODE_MANUAL As Integer = 1
    Public Const DISPLAYMODE_AUTO1 As Integer = 2
    Public Const DISPLAYMODE_AUTO60 As Integer = 3
    Public Const DISPLAYMODE_INVALID As Integer = -1
    REM --- (end of YSegmentedDisplay definitions)

    REM --- (YSegmentedDisplay attributes declaration)
    Protected _displayedText As String
    Protected _displayMode As Integer
    Protected _valueCallbackSegmentedDisplay As YSegmentedDisplayValueCallback
    REM --- (end of YSegmentedDisplay attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "SegmentedDisplay"
      REM --- (YSegmentedDisplay attributes initialization)
      _displayedText = DISPLAYEDTEXT_INVALID
      _displayMode = DISPLAYMODE_INVALID
      _valueCallbackSegmentedDisplay = Nothing
      REM --- (end of YSegmentedDisplay attributes initialization)
    End Sub

    REM --- (YSegmentedDisplay private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "displayedText") Then
        _displayedText = member.svalue
        Return 1
      End If
      If (member.name = "displayMode") Then
        _displayMode = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YSegmentedDisplay private methods declaration)

    REM --- (YSegmentedDisplay public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the text currently displayed on the screen.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the text currently displayed on the screen
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DISPLAYEDTEXT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_displayedText() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return DISPLAYEDTEXT_INVALID
        End If
      End If
      Return Me._displayedText
    End Function


    '''*
    ''' <summary>
    '''   Changes the text currently displayed on the screen.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the text currently displayed on the screen
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
    Public Function set_displayedText(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("displayedText", rest_val)
    End Function
    Public Function get_displayMode() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return DISPLAYMODE_INVALID
        End If
      End If
      Return Me._displayMode
    End Function


    Public Function set_displayMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("displayMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a segmented display for a given identifier.
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
    '''   This function does not require that the segmented displays is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSegmentedDisplay.isOnline()</c> to test if the segmented displays is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a segmented display by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the segmented displays
    ''' </param>
    ''' <returns>
    '''   a <c>YSegmentedDisplay</c> object allowing you to drive the segmented displays.
    ''' </returns>
    '''/
    Public Shared Function FindSegmentedDisplay(func As String) As YSegmentedDisplay
      Dim obj As YSegmentedDisplay
      obj = CType(YFunction._FindFromCache("SegmentedDisplay", func), YSegmentedDisplay)
      If ((obj Is Nothing)) Then
        obj = New YSegmentedDisplay(func)
        YFunction._AddToCache("SegmentedDisplay", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YSegmentedDisplayValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSegmentedDisplay = callback
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
      If (Not (Me._valueCallbackSegmentedDisplay Is Nothing)) Then
        Me._valueCallbackSegmentedDisplay(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of segmented displays started using <c>yFirstSegmentedDisplay()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSegmentedDisplay</c> object, corresponding to
    '''   a segmented display currently online, or a <c>null</c> pointer
    '''   if there are no more segmented displays to enumerate.
    ''' </returns>
    '''/
    Public Function nextSegmentedDisplay() As YSegmentedDisplay
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSegmentedDisplay.FindSegmentedDisplay(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of segmented displays currently accessible.
    ''' <para>
    '''   Use the method <c>YSegmentedDisplay.nextSegmentedDisplay()</c> to iterate on
    '''   next segmented displays.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSegmentedDisplay</c> object, corresponding to
    '''   the first segmented displays currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstSegmentedDisplay() As YSegmentedDisplay
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("SegmentedDisplay", 0, p, size, neededsize, errmsg)
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
      Return YSegmentedDisplay.FindSegmentedDisplay(serial + "." + funcId)
    End Function

    REM --- (end of YSegmentedDisplay public methods declaration)

  End Class

  REM --- (SegmentedDisplay functions)

  '''*
  ''' <summary>
  '''   Retrieves a segmented display for a given identifier.
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
  '''   This function does not require that the segmented displays is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSegmentedDisplay.isOnline()</c> to test if the segmented displays is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a segmented display by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the segmented displays
  ''' </param>
  ''' <returns>
  '''   a <c>YSegmentedDisplay</c> object allowing you to drive the segmented displays.
  ''' </returns>
  '''/
  Public Function yFindSegmentedDisplay(ByVal func As String) As YSegmentedDisplay
    Return YSegmentedDisplay.FindSegmentedDisplay(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of segmented displays currently accessible.
  ''' <para>
  '''   Use the method <c>YSegmentedDisplay.nextSegmentedDisplay()</c> to iterate on
  '''   next segmented displays.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YSegmentedDisplay</c> object, corresponding to
  '''   the first segmented displays currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSegmentedDisplay() As YSegmentedDisplay
    Return YSegmentedDisplay.FirstSegmentedDisplay()
  End Function


  REM --- (end of SegmentedDisplay functions)

End Module
