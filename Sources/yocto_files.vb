'*********************************************************************
'*
'* $Id: yocto_files.vb 17674 2014-09-16 16:18:58Z seb $
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
    Public Overridable Function get_name() As String
      Return Me._name
    End Function

    Public Overridable Function get_size() As Integer
      Return Me._size
    End Function

    Public Overridable Function get_crc() As Integer
      Return Me._crc
    End Function



    REM --- (end of generated code: YFileRecord public methods declaration)

   Public Sub New(ByVal data As String)
      Dim p As TJsonParser
      Dim node As Nullable(Of TJSONRECORD)
      p = New TJsonParser(data, False)
      node = p.GetChildNode(Nothing, "name")
      Me._name = node.Value.svalue
      node = p.GetChildNode(Nothing, "size")
      Me._size = CInt(node.Value.ivalue)
      node = p.GetChildNode(Nothing, "crc")
      Me._crc = CInt(node.Value.ivalue)
    End Sub  
    
  End Class
  
  REM --- (generated code: FileRecord functions)


  REM --- (end of generated code: FileRecord functions)



  REM --- (generated code: YFiles class start)

  '''*
  ''' <summary>
  '''   The filesystem interface makes it possible to store files
  '''   on some devices, for instance to design a custom web UI
  '''   (for networked devices) or to add fonts (on display
  '''   devices).
  ''' <para>
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
    REM --- (end of generated code: YFiles attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.new(func)
      _className = "Files"
      REM --- (generated code: YFiles attributes initialization)
      _filesCount = FILESCOUNT_INVALID
      _freeSpace = FREESPACE_INVALID
      _valueCallbackFiles = Nothing
      REM --- (end of generated code: YFiles attributes initialization)
    End Sub
    
    REM --- (generated code: YFiles private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "filesCount") Then
        _filesCount = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "freeSpace") Then
        _freeSpace = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
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
    '''   On failure, throws an exception or returns <c>Y_FILESCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_filesCount() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return FILESCOUNT_INVALID
        End If
      End If
      Return Me._filesCount
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
    '''   On failure, throws an exception or returns <c>Y_FREESPACE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_freeSpace() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return FREESPACE_INVALID
        End If
      End If
      Return Me._freeSpace
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
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the filesystem
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
    Public Overloads Function registerValueCallback(callback As YFilesValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackFiles = callback
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
      REM // may throw an exception
      Return Me._download(url)
    End Function

    '''*
    ''' <summary>
    '''   Reinitialize the filesystem to its clean, unfragmented, empty state.
    ''' <para>
    '''   All files previously uploaded are permanently lost.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function format_fs() As Integer
      Dim json As Byte()
      Dim res As String
      json = Me.sendCommand("format")
      res = Me._json_get_key(json, "res")
      If Not(res = "ok") Then
        me._throw( YAPI.IO_ERROR,  "format failed")
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
    '''   as wildcards. When an empty pattern is provided, all file records
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
      Dim i_i As Integer
      Dim json As Byte()
      Dim filelist As List(Of String) = New List(Of String)()
      Dim res As List(Of YFileRecord) = New List(Of YFileRecord)()
      json = Me.sendCommand("dir&f=" + pattern)
      filelist = Me._json_get_array(json)
      res.Clear()
      For i_i = 0 To filelist.Count - 1
        res.Add(New YFileRecord(filelist(i_i)))
      Next i_i
      Return res
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function remove(pathname As String) As Integer
      Dim json As Byte()
      Dim res As String
      json = Me.sendCommand("del&f=" + pathname)
      res  = Me._json_get_key(json, "res")
      If Not(res = "ok") Then
        me._throw( YAPI.IO_ERROR,  "unable to remove file")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of filesystems started using <c>yFirstFiles()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YFiles</c> object, corresponding to
    '''   a filesystem currently online, or a <c>null</c> pointer
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
    '''   the first filesystem currently online, or a <c>null</c> pointer
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

  REM --- (generated code: Files functions)

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
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the filesystem
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
  '''   the first filesystem currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstFiles() As YFiles
    Return YFiles.FirstFiles()
  End Function


  REM --- (end of generated code: Files functions)

End Module
