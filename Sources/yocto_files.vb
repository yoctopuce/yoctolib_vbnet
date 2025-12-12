'*********************************************************************
'*
'* $Id: yocto_files.vb 70518 2025-11-26 16:18:50Z mvuilleu $
'*
'* Implements yFindFiles(), the high-level API for Files functions
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

Module yocto_files
  REM --- (generated code: YFileRecord globals)

  REM --- (end of generated code: YFileRecord globals)
  REM --- (generated code: YFiles globals)

  Public Const Y_FILESCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_FREESPACE_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YFilesValueCallback(ByVal func As YFiles, ByVal value As String)
  Public Delegate Sub YFilesTimedReportCallback(ByVal func As YFiles, ByVal measure As YMeasure)
  REM --- (end of generated code: YFiles globals)


  REM --- (generated code: YFileRecord class start)

  '''*
  ''' <c>YFileRecord</c> objects are used to describe a file that is stored on a Yoctopuce device.
  ''' These objects are used in particular in conjunction with the <c>YFiles</c> class.
  '''/
  Public Class YFileRecord
    REM --- (end of generated code: YFileRecord class start)

    REM --- (generated code: YFileRecord definitions)
    REM --- (end of generated code: YFileRecord definitions)

    REM --- (generated code: YFileRecord attributes declaration)
    Protected _name As String
    Protected _size As Integer
    Protected _crc As Integer
    REM --- (end of generated code: YFileRecord attributes declaration)

    REM --- (generated code: YFileRecord private methods declaration)

    REM --- (end of generated code: YFileRecord private methods declaration)

    REM --- (generated code: YFileRecord public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the name of the file.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with the name of the file.
    ''' </returns>
    '''/
    Public Overridable Function get_name() As String
      Return Me._name
    End Function

    '''*
    ''' <summary>
    '''   Returns the size of the file in bytes.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the size of the file.
    ''' </returns>
    '''/
    Public Overridable Function get_size() As Integer
      Return Me._size
    End Function

    '''*
    ''' <summary>
    '''   Returns the 32-bit CRC of the file content.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the 32-bit CRC of the file content.
    ''' </returns>
    '''/
    Public Overridable Function get_crc() As Integer
      Return Me._crc
    End Function



    REM --- (end of generated code: YFileRecord public methods declaration)

   Public Sub New(ByVal data As String)
      Dim obj As YJSONObject  = New YJSONObject(data)
      obj.parse()
      Me._name = obj.getString("name")
      Me._size = CInt(obj.getInt("size"))
      Me._crc = CInt(obj.getInt("crc"))
    End Sub

  End Class

  REM --- (generated code: YFileRecord functions)


  REM --- (end of generated code: YFileRecord functions)



  REM --- (generated code: YFiles class start)

  '''*
  ''' <summary>
  '''   The YFiles class is used to access the filesystem embedded on
  '''   some Yoctopuce devices.
  ''' <para>
  '''   This filesystem makes it
  '''   possible for instance to design a custom web UI
  '''   (for networked devices) or to add fonts (on display devices).
  ''' </para>
  ''' </summary>
  '''/
  Public Class YFiles
    Inherits YFunction
    REM --- (end of generated code: YFiles class start)

    REM --- (generated code: YFiles definitions)
    Public Const FILESCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const FREESPACE_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of generated code: YFiles definitions)

    REM --- (generated code: YFiles attributes declaration)
    Protected _filesCount As Integer
    Protected _freeSpace As Integer
    Protected _valueCallbackFiles As YFilesValueCallback
    Protected _ver As Integer
    REM --- (end of generated code: YFiles attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.new(func)
      _className = "Files"
      REM --- (generated code: YFiles attributes initialization)
      _filesCount = FILESCOUNT_INVALID
      _freeSpace = FREESPACE_INVALID
      _valueCallbackFiles = Nothing
      _ver = 0
      REM --- (end of generated code: YFiles attributes initialization)
    End Sub

    REM --- (generated code: YFiles private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("filesCount") Then
        _filesCount = CInt(json_val.getLong("filesCount"))
      End If
      If json_val.has("freeSpace") Then
        _freeSpace = CInt(json_val.getLong("freeSpace"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YFiles private methods declaration)

    REM --- (generated code: YFiles public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the number of files currently loaded in the filesystem.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of files currently loaded in the filesystem
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YFiles.FILESCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_filesCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return FILESCOUNT_INVALID
        End If
      End If
      res = Me._filesCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the free space for uploading new files to the filesystem, in bytes.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the free space for uploading new files to the filesystem, in bytes
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YFiles.FREESPACE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_freeSpace() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return FREESPACE_INVALID
        End If
      End If
      res = Me._freeSpace
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a filesystem for a given identifier.
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
    '''   This function does not require that the filesystem is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YFiles.isOnline()</c> to test if the filesystem is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a filesystem by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the filesystem, for instance
    '''   <c>YRGBLED2.files</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YFiles</c> object allowing you to drive the filesystem.
    ''' </returns>
    '''/
    Public Shared Function FindFiles(func As String) As YFiles
      Dim obj As YFiles
      obj = CType(YFunction._FindFromCache("Files", func), YFiles)
      If ((obj Is Nothing)) Then
        obj = New YFiles(func)
        YFunction._AddToCache("Files", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YFilesValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackFiles = callback
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
      If (Not (Me._valueCallbackFiles Is Nothing)) Then
        Me._valueCallbackFiles(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function sendCommand(command As String) As Byte()
      Dim url As String
      url = "files.json?a=" + command

      Return Me._download(url)
    End Function

    Public Overridable Function _getVersion() As Integer
      Dim json As Byte() = New Byte(){}
      If (Me._ver > 0) Then
        Return Me._ver
      End If
      REM //may throw an exception
      json = Me.sendCommand("info")
      If (json(0) <> 123) Then
        REM // ascii code for '{'
        Me._ver = 30
      Else
        Me._ver = YAPI._atoi(Me._json_get_key(json, "ver"))
      End If
      Return Me._ver
    End Function

    '''*
    ''' <summary>
    '''   Reinitialize the filesystem to its clean, unfragmented, empty state.
    ''' <para>
    '''   All files previously uploaded are permanently lost.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function format_fs() As Integer
      Dim json As Byte() = New Byte(){}
      Dim res As String
      json = Me.sendCommand("format")
      res = Me._json_get_key(json, "res")
      If Not(res = "ok") Then
        me._throw(YAPI.IO_ERROR, "format failed")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of YFileRecord objects that describe files currently loaded
    '''   in the filesystem.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="pattern">
    '''   an optional filter pattern, using star and question marks
    '''   as wild cards. When an empty pattern is provided, all file records
    '''   are returned.
    ''' </param>
    ''' <returns>
    '''   a list of <c>YFileRecord</c> objects, containing the file path
    '''   and name, byte size and 32-bit CRC of the file content.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list.
    ''' </para>
    '''/
    Public Overridable Function get_list(pattern As String) As List(Of YFileRecord)
      Dim json As Byte() = New Byte(){}
      Dim filelist As List(Of Byte()) = New List(Of Byte())()
      Dim res As List(Of YFileRecord) = New List(Of YFileRecord)()
      json = Me.sendCommand("dir&f=" + pattern)
      filelist = Me._json_get_array(json)
      res.Clear()
      Dim ii_0 As Integer
      For ii_0 = 0 To filelist.Count - 1
        res.Add(New YFileRecord(YAPI.DefaultEncoding.GetString(filelist(ii_0))))
      Next ii_0
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Tests if a file exists on the filesystem of the module.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="filename">
    '''   the filename to test.
    ''' </param>
    ''' <returns>
    '''   true if the file exists, false otherwise.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception.
    ''' </para>
    '''/
    Public Overridable Function fileExist(filename As String) As Boolean
      Dim json As Byte() = New Byte(){}
      Dim filelist As List(Of Byte()) = New List(Of Byte())()
      If ((filename).Length = 0) Then
        Return False
      End If
      json = Me.sendCommand("dir&f=" + filename)
      filelist = Me._json_get_array(json)
      If (filelist.Count > 0) Then
        Return True
      End If
      Return False
    End Function

    '''*
    ''' <summary>
    '''   Downloads the requested file and returns a binary buffer with its content.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="pathname">
    '''   path and name of the file to download
    ''' </param>
    ''' <returns>
    '''   a binary buffer with the file content
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty content.
    ''' </para>
    '''/
    Public Overridable Function download(pathname As String) As Byte()
      Return Me._download(pathname)
    End Function

    '''*
    ''' <summary>
    '''   Uploads a file to the filesystem, to the specified full path name.
    ''' <para>
    '''   If a file already exists with the same path name, its content is overwritten.
    ''' </para>
    ''' </summary>
    ''' <param name="pathname">
    '''   path and name of the new file to create
    ''' </param>
    ''' <param name="content">
    '''   binary buffer with the content to set
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function upload(pathname As String, content As Byte()) As Integer
      Return Me._upload(pathname, content)
    End Function

    '''*
    ''' <summary>
    '''   Deletes a file, given by its full path name, from the filesystem.
    ''' <para>
    '''   Because of filesystem fragmentation, deleting a file may not always
    '''   free up the whole space used by the file. However, rewriting a file
    '''   with the same path name will always reuse any space not freed previously.
    '''   If you need to ensure that no space is taken by previously deleted files,
    '''   you can use <c>format_fs</c> to fully reinitialize the filesystem.
    ''' </para>
    ''' </summary>
    ''' <param name="pathname">
    '''   path and name of the file to remove.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function remove(pathname As String) As Integer
      Dim json As Byte() = New Byte(){}
      Dim res As String
      json = Me.sendCommand("del&f=" + pathname)
      res  = Me._json_get_key(json, "res")
      If Not(res = "ok") Then
        me._throw(YAPI.IO_ERROR, "unable to remove file")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns the expected file CRC for a given content.
    ''' <para>
    '''   Note that the CRC value may vary depending on the version
    '''   of the filesystem used by the hub, so it is important to
    '''   use this method if a reference value needs to be computed.
    ''' </para>
    ''' </summary>
    ''' <param name="content">
    '''   a buffer representing a file content
    ''' </param>
    ''' <returns>
    '''   the 32-bit CRC summarizing the file content, as it would
    '''   be returned by the <c>get_crc()</c> method of
    '''   <c>YFileRecord</c> objects returned by <c>get_list()</c>.
    ''' </returns>
    '''/
    Public Overridable Function get_content_crc(content As Byte()) As Integer
      Dim fsver As Integer = 0
      Dim sz As Integer = 0
      Dim blkcnt As Integer = 0
      Dim meta As Byte() = New Byte(){}
      Dim blkidx As Integer = 0
      Dim blksz As Integer = 0
      Dim part As Integer = 0
      Dim res As Integer = 0
      sz = (content).Length

      fsver = Me._getVersion()
      If (fsver < 40) Then
        res = YAPI._bincrc(content, 0, sz)
        res = (((res) And (&H7fffffff)) - 2 * (((res >> 1)) And (&H40000000)))
        Return res
      End If
      blkcnt = ((sz + 255) \ 256)
      ReDim meta(4 * blkcnt-1)
      blkidx = 0
      While (blkidx < blkcnt)
        blksz = sz - blkidx * 256
        If (blksz > 256) Then
          blksz = 256
        End If
        part = ((YAPI._bincrc(content, blkidx * 256, blksz)) Xor (CType(&Hffffffff, Integer)))
        meta(4 * blkidx) = Convert.ToByte(((part) And (255)) And &HFF)
        meta(4 * blkidx + 1) = Convert.ToByte((((part >> 8)) And (255)) And &HFF)
        meta(4 * blkidx + 2) = Convert.ToByte((((part >> 16)) And (255)) And &HFF)
        meta(4 * blkidx + 3) = Convert.ToByte((((part >> 24)) And (255)) And &HFF)
        blkidx = blkidx + 1
      End While
      res = ((YAPI._bincrc(meta, 0, 4 * blkcnt)) Xor (CType(&Hffffffff, Integer)))
      res = (((res) And (&H7fffffff)) - 2 * (((res >> 1)) And (&H40000000)))
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of filesystems started using <c>yFirstFiles()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned filesystems order.
    '''   If you want to find a specific a filesystem, use <c>Files.findFiles()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YFiles</c> object, corresponding to
    '''   a filesystem currently online, or a <c>Nothing</c> pointer
    '''   if there are no more filesystems to enumerate.
    ''' </returns>
    '''/
    Public Function nextFiles() As YFiles
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YFiles.FindFiles(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of filesystems currently accessible.
    ''' <para>
    '''   Use the method <c>YFiles.nextFiles()</c> to iterate on
    '''   next filesystems.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YFiles</c> object, corresponding to
    '''   the first filesystem currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstFiles() As YFiles
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Files", 0, p, size, neededsize, errmsg)
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
      Return YFiles.FindFiles(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YFiles public methods declaration)

  End Class

  REM --- (generated code: YFiles functions)

  '''*
  ''' <summary>
  '''   Retrieves a filesystem for a given identifier.
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
  '''   This function does not require that the filesystem is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YFiles.isOnline()</c> to test if the filesystem is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a filesystem by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the filesystem, for instance
  '''   <c>YRGBLED2.files</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YFiles</c> object allowing you to drive the filesystem.
  ''' </returns>
  '''/
  Public Function yFindFiles(ByVal func As String) As YFiles
    Return YFiles.FindFiles(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of filesystems currently accessible.
  ''' <para>
  '''   Use the method <c>YFiles.nextFiles()</c> to iterate on
  '''   next filesystems.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YFiles</c> object, corresponding to
  '''   the first filesystem currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstFiles() As YFiles
    Return YFiles.FirstFiles()
  End Function


  REM --- (end of generated code: YFiles functions)

End Module
