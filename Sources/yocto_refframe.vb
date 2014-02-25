'*********************************************************************
'*
'* $Id: yocto_refframe.vb 15039 2014-02-24 11:22:11Z seb $
'*
'* Implements yFindRefFrame(), the high-level API for RefFrame functions
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

Module yocto_refframe

    REM --- (YRefFrame return codes)
    REM --- (end of YRefFrame return codes)
  REM --- (YRefFrame globals)

 Public Enum   Y_MOUNTPOSITION
  BOTTOM = 0
  TOP = 1
  FRONT = 2
  RIGHT = 3
  REAR = 4
  LEFT = 5
end enum

 Public Enum   Y_MOUNTORIENTATION
  TWELVE = 0
  THREE = 1
  SIX = 2
  NINE = 3
end enum

  Public Const Y_MOUNTPOS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BEARING_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YRefFrameValueCallback(ByVal func As YRefFrame, ByVal value As String)
  Public Delegate Sub YRefFrameTimedReportCallback(ByVal func As YRefFrame, ByVal measure As YMeasure)
  REM --- (end of YRefFrame globals)

  REM --- (YRefFrame class start)

  '''*
  ''' <summary>
  '''   This class is used to setup the base orientation of the device, so that
  '''   the orientation functions, relative to the earth surface plane, use
  '''   the proper reference frame.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YRefFrame
    Inherits YFunction
    REM --- (end of YRefFrame class start)

    REM --- (YRefFrame definitions)
    Public Const MOUNTPOS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BEARING_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YRefFrame definitions)

    REM --- (YRefFrame attributes declaration)
    Protected _mountPos As Integer
    Protected _bearing As Double
    Protected _calibrationParam As String
    Protected _valueCallbackRefFrame As YRefFrameValueCallback
    REM --- (end of YRefFrame attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "RefFrame"
      REM --- (YRefFrame attributes initialization)
      _mountPos = MOUNTPOS_INVALID
      _bearing = BEARING_INVALID
      _calibrationParam = CALIBRATIONPARAM_INVALID
      _valueCallbackRefFrame = Nothing
      REM --- (end of YRefFrame attributes initialization)
    End Sub

  REM --- (YRefFrame private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "mountPos") Then
        _mountPos = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "bearing") Then
        _bearing = member.ivalue / 65536.0
        Return 1
      End If
      If (member.name = "calibrationParam") Then
        _calibrationParam = member.svalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YRefFrame private methods declaration)

    REM --- (YRefFrame public methods declaration)
    Public Function get_mountPos() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MOUNTPOS_INVALID
        End If
      End If
      Return Me._mountPos
    End Function


    Public Function set_mountPos(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("mountPos", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the reference bearing used by the compass.
    ''' <para>
    '''   The relative bearing
    '''   indicated by the compass is the difference between the measured magnetic
    '''   heading and the reference bearing indicated here.
    ''' </para>
    ''' <para>
    '''   For instance, if you setup as reference bearing the value of the earth
    '''   magnetic declination, the compass will provide the orientation relative
    '''   to the geographic North.
    ''' </para>
    ''' <para>
    '''   Similarly, when the sensor is not mounted along the standard directions
    '''   because it has an additional yaw angle, you can set this angle in the reference
    '''   bearing so that the compass provides the expected natural direction.
    ''' </para>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the reference bearing used by the compass
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
    Public Function set_bearing(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("bearing", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the reference bearing used by the compass.
    ''' <para>
    '''   The relative bearing
    '''   indicated by the compass is the difference between the measured magnetic
    '''   heading and the reference bearing indicated here.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the reference bearing used by the compass
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BEARING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bearing() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BEARING_INVALID
        End If
      End If
      Return Me._bearing
    End Function

    Public Function get_calibrationParam() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALIBRATIONPARAM_INVALID
        End If
      End If
      Return Me._calibrationParam
    End Function


    Public Function set_calibrationParam(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("calibrationParam", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a reference frame for a given identifier.
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
    '''   This function does not require that the reference frame is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YRefFrame.isOnline()</c> to test if the reference frame is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a reference frame by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the reference frame
    ''' </param>
    ''' <returns>
    '''   a <c>YRefFrame</c> object allowing you to drive the reference frame.
    ''' </returns>
    '''/
    Public Shared Function FindRefFrame(func As String) As YRefFrame
      Dim obj As YRefFrame
      obj = CType(YFunction._FindFromCache("RefFrame", func), YRefFrame)
      If ((obj Is Nothing)) Then
        obj = New YRefFrame(func)
        YFunction._AddToCache("RefFrame", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YRefFrameValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me , True)
      Else
        YFunction._UpdateValueCallbackList(Me , False)
      End If
      Me._valueCallbackRefFrame = callback
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
      If (Not (Me._valueCallbackRefFrame Is Nothing)) Then
        Me._valueCallbackRefFrame(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the installation position of the device, as configured
    '''   in order to define the reference frame for the compass and the
    '''   pitch/roll tilt sensors.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among the <c>Y_MOUNTPOSITION</c> enumeration
    '''   (<c>Y_MOUNTPOSITION_BOTTOM</c>,   <c>Y_MOUNTPOSITION_TOP</c>,
    '''   <c>Y_MOUNTPOSITION_FRONT</c>,    <c>Y_MOUNTPOSITION_RIGHT</c>,
    '''   <c>Y_MOUNTPOSITION_REAR</c>,     <c>Y_MOUNTPOSITION_LEFT</c>),
    '''   corresponding to the installation in a box, on one of the six faces.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_mountPosition() As Y_MOUNTPOSITION
      Dim pos As Integer = 0
      pos = Me.get_mountPos()
      return CType(((pos) >> (2)), Y_MOUNTPOSITION)
    End Function

    '''*
    ''' <summary>
    '''   Returns the installation orientation of the device, as configured
    '''   in order to define the reference frame for the compass and the
    '''   pitch/roll tilt sensors.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among the enumeration <c>Y_MOUNTORIENTATION</c>
    '''   (<c>Y_MOUNTORIENTATION_TWELVE</c>, <c>Y_MOUNTORIENTATION_THREE</c>,
    '''   <c>Y_MOUNTORIENTATION_SIX</c>,     <c>Y_MOUNTORIENTATION_NINE</c>)
    '''   corresponding to the orientation of the "X" arrow on the device,
    '''   as on a clock dial seen from an observer in the center of the box.
    '''   On the bottom face, the 12H orientation points to the front, while
    '''   on the top face, the 12H orientation points to the rear.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_mountOrientation() As Y_MOUNTORIENTATION
      Dim pos As Integer = 0
      pos = Me.get_mountPos()
      return CType(((pos) And (3)), Y_MOUNTORIENTATION)
    End Function

    '''*
    ''' <summary>
    '''   Changes the compass and tilt sensor frame of reference.
    ''' <para>
    '''   The magnetic compass
    '''   and the tilt sensors (pitch and roll) naturally work in the plane
    '''   parallel to the earth surface. In case the device is not installed upright
    '''   and horizontally, you must select its reference orientation (parallel to
    '''   the earth surface) so that the measures are made relative to this position.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="position">
    '''   a value among the <c>Y_MOUNTPOSITION</c> enumeration
    '''   (<c>Y_MOUNTPOSITION_BOTTOM</c>,   <c>Y_MOUNTPOSITION_TOP</c>,
    '''   <c>Y_MOUNTPOSITION_FRONT</c>,    <c>Y_MOUNTPOSITION_RIGHT</c>,
    '''   <c>Y_MOUNTPOSITION_REAR</c>,     <c>Y_MOUNTPOSITION_LEFT</c>),
    '''   corresponding to the installation in a box, on one of the six faces.
    ''' </param>
    ''' <param name="orientation">
    '''   a value among the enumeration <c>Y_MOUNTORIENTATION</c>
    '''   (<c>Y_MOUNTORIENTATION_TWELVE</c>, <c>Y_MOUNTORIENTATION_THREE</c>,
    '''   <c>Y_MOUNTORIENTATION_SIX</c>,     <c>Y_MOUNTORIENTATION_NINE</c>)
    '''   corresponding to the orientation of the "X" arrow on the device,
    '''   as on a clock dial seen from an observer in the center of the box.
    '''   On the bottom face, the 12H orientation points to the front, while
    '''   on the top face, the 12H orientation points to the rear.
    ''' </param>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_mountPosition(position As Y_MOUNTPOSITION, orientation As Y_MOUNTORIENTATION) As Integer
      Dim pos As Integer = 0
      pos = ((position) << (2)) + orientation
      return Me.set_mountPos(pos)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of reference frames started using <c>yFirstRefFrame()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRefFrame</c> object, corresponding to
    '''   a reference frame currently online, or a <c>null</c> pointer
    '''   if there are no more reference frames to enumerate.
    ''' </returns>
    '''/
    Public Function nextRefFrame() As YRefFrame
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YRefFrame.FindRefFrame(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of reference frames currently accessible.
    ''' <para>
    '''   Use the method <c>YRefFrame.nextRefFrame()</c> to iterate on
    '''   next reference frames.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRefFrame</c> object, corresponding to
    '''   the first reference frame currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstRefFrame() As YRefFrame
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("RefFrame", 0, p, size, neededsize, errmsg)
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
      Return YRefFrame.FindRefFrame(serial + "." + funcId)
    End Function

    REM --- (end of YRefFrame public methods declaration)

  End Class

  REM --- (RefFrame functions)

  '''*
  ''' <summary>
  '''   Retrieves a reference frame for a given identifier.
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
  '''   This function does not require that the reference frame is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YRefFrame.isOnline()</c> to test if the reference frame is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a reference frame by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the reference frame
  ''' </param>
  ''' <returns>
  '''   a <c>YRefFrame</c> object allowing you to drive the reference frame.
  ''' </returns>
  '''/
  Public Function yFindRefFrame(ByVal func As String) As YRefFrame
    Return YRefFrame.FindRefFrame(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of reference frames currently accessible.
  ''' <para>
  '''   Use the method <c>YRefFrame.nextRefFrame()</c> to iterate on
  '''   next reference frames.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YRefFrame</c> object, corresponding to
  '''   the first reference frame currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstRefFrame() As YRefFrame
    Return YRefFrame.FirstRefFrame()
  End Function


  REM --- (end of RefFrame functions)

End Module
