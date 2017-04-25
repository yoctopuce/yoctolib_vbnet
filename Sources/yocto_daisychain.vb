'*********************************************************************
'*
'* $Id: yocto_daisychain.vb 27237 2017-04-21 16:36:03Z seb $
'*
'* Implements yFindDaisyChain(), the high-level API for DaisyChain functions
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

Module yocto_daisychain

    REM --- (YDaisyChain return codes)
    REM --- (end of YDaisyChain return codes)
    REM --- (YDaisyChain dlldef)
    REM --- (end of YDaisyChain dlldef)
  REM --- (YDaisyChain globals)

  Public Const Y_DAISYSTATE_READY As Integer = 0
  Public Const Y_DAISYSTATE_IS_CHILD As Integer = 1
  Public Const Y_DAISYSTATE_FIRMWARE_MISMATCH As Integer = 2
  Public Const Y_DAISYSTATE_CHILD_MISSING As Integer = 3
  Public Const Y_DAISYSTATE_CHILD_LOST As Integer = 4
  Public Const Y_DAISYSTATE_INVALID As Integer = -1
  Public Const Y_CHILDCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_REQUIREDCHILDCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YDaisyChainValueCallback(ByVal func As YDaisyChain, ByVal value As String)
  Public Delegate Sub YDaisyChainTimedReportCallback(ByVal func As YDaisyChain, ByVal measure As YMeasure)
  REM --- (end of YDaisyChain globals)

  REM --- (YDaisyChain class start)

  '''*
  ''' <summary>
  '''   The YDaisyChain interface can be used to verify that devices that
  '''   are daisy-chained directly from device to device, without a hub,
  '''   are detected properly.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDaisyChain
    Inherits YFunction
    REM --- (end of YDaisyChain class start)

    REM --- (YDaisyChain definitions)
    Public Const DAISYSTATE_READY As Integer = 0
    Public Const DAISYSTATE_IS_CHILD As Integer = 1
    Public Const DAISYSTATE_FIRMWARE_MISMATCH As Integer = 2
    Public Const DAISYSTATE_CHILD_MISSING As Integer = 3
    Public Const DAISYSTATE_CHILD_LOST As Integer = 4
    Public Const DAISYSTATE_INVALID As Integer = -1
    Public Const CHILDCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const REQUIREDCHILDCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YDaisyChain definitions)

    REM --- (YDaisyChain attributes declaration)
    Protected _daisyState As Integer
    Protected _childCount As Integer
    Protected _requiredChildCount As Integer
    Protected _valueCallbackDaisyChain As YDaisyChainValueCallback
    REM --- (end of YDaisyChain attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "DaisyChain"
      REM --- (YDaisyChain attributes initialization)
      _daisyState = DAISYSTATE_INVALID
      _childCount = CHILDCOUNT_INVALID
      _requiredChildCount = REQUIREDCHILDCOUNT_INVALID
      _valueCallbackDaisyChain = Nothing
      REM --- (end of YDaisyChain attributes initialization)
    End Sub

    REM --- (YDaisyChain private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("daisyState") Then
        _daisyState = CInt(json_val.getLong("daisyState"))
      End If
      If json_val.has("childCount") Then
        _childCount = CInt(json_val.getLong("childCount"))
      End If
      If json_val.has("requiredChildCount") Then
        _requiredChildCount = CInt(json_val.getLong("requiredChildCount"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YDaisyChain private methods declaration)

    REM --- (YDaisyChain public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the state of the daisy-link between modules.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_DAISYSTATE_READY</c>, <c>Y_DAISYSTATE_IS_CHILD</c>,
    '''   <c>Y_DAISYSTATE_FIRMWARE_MISMATCH</c>, <c>Y_DAISYSTATE_CHILD_MISSING</c> and
    '''   <c>Y_DAISYSTATE_CHILD_LOST</c> corresponding to the state of the daisy-link between modules
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DAISYSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_daisyState() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return DAISYSTATE_INVALID
        End If
      End If
      res = Me._daisyState
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of child nodes currently detected.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of child nodes currently detected
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CHILDCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_childCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CHILDCOUNT_INVALID
        End If
      End If
      res = Me._childCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of child nodes expected in normal conditions.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of child nodes expected in normal conditions
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_REQUIREDCHILDCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_requiredChildCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return REQUIREDCHILDCOUNT_INVALID
        End If
      End If
      res = Me._requiredChildCount
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of child nodes expected in normal conditions.
    ''' <para>
    '''   If the value is zero, no check is performed. If it is non-zero, the number
    '''   child nodes is checked on startup and the status will change to error if
    '''   the count does not match.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the number of child nodes expected in normal conditions
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
    Public Function set_requiredChildCount(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("requiredChildCount", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a module chain for a given identifier.
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
    '''   This function does not require that the module chain is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YDaisyChain.isOnline()</c> to test if the module chain is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a module chain by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the module chain
    ''' </param>
    ''' <returns>
    '''   a <c>YDaisyChain</c> object allowing you to drive the module chain.
    ''' </returns>
    '''/
    Public Shared Function FindDaisyChain(func As String) As YDaisyChain
      Dim obj As YDaisyChain
      obj = CType(YFunction._FindFromCache("DaisyChain", func), YDaisyChain)
      If ((obj Is Nothing)) Then
        obj = New YDaisyChain(func)
        YFunction._AddToCache("DaisyChain", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YDaisyChainValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackDaisyChain = callback
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
      If (Not (Me._valueCallbackDaisyChain Is Nothing)) Then
        Me._valueCallbackDaisyChain(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of module chains started using <c>yFirstDaisyChain()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDaisyChain</c> object, corresponding to
    '''   a module chain currently online, or a <c>Nothing</c> pointer
    '''   if there are no more module chains to enumerate.
    ''' </returns>
    '''/
    Public Function nextDaisyChain() As YDaisyChain
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YDaisyChain.FindDaisyChain(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of module chains currently accessible.
    ''' <para>
    '''   Use the method <c>YDaisyChain.nextDaisyChain()</c> to iterate on
    '''   next module chains.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDaisyChain</c> object, corresponding to
    '''   the first module chain currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstDaisyChain() As YDaisyChain
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("DaisyChain", 0, p, size, neededsize, errmsg)
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
      Return YDaisyChain.FindDaisyChain(serial + "." + funcId)
    End Function

    REM --- (end of YDaisyChain public methods declaration)

  End Class

  REM --- (DaisyChain functions)

  '''*
  ''' <summary>
  '''   Retrieves a module chain for a given identifier.
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
  '''   This function does not require that the module chain is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YDaisyChain.isOnline()</c> to test if the module chain is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a module chain by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the module chain
  ''' </param>
  ''' <returns>
  '''   a <c>YDaisyChain</c> object allowing you to drive the module chain.
  ''' </returns>
  '''/
  Public Function yFindDaisyChain(ByVal func As String) As YDaisyChain
    Return YDaisyChain.FindDaisyChain(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of module chains currently accessible.
  ''' <para>
  '''   Use the method <c>YDaisyChain.nextDaisyChain()</c> to iterate on
  '''   next module chains.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YDaisyChain</c> object, corresponding to
  '''   the first module chain currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstDaisyChain() As YDaisyChain
    Return YDaisyChain.FirstDaisyChain()
  End Function


  REM --- (end of DaisyChain functions)

End Module
