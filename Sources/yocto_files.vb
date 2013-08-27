'*********************************************************************
'*
'* $Id: yocto_files.vb 12326 2013-08-13 15:52:20Z mvuilleu $
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

  REM --- (definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YFiles, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_FILESCOUNT_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_FREESPACE_INVALID As Integer = YAPI.INVALID_UNSIGNED

  REM --- (end of generated code: definitions)



  REM--- (generated code: YFileRecord implementation)


  Public Class YFileRecord



    public function get_name() as string
        Return Me._name
     end function

    public function get_size() as integer
        Return Me._size
     end function

    public function get_crc() as integer
        Return Me._crc
     end function

    public function name() as string
        Return Me._name
     end function

    public function size() as integer
        Return Me._size
     end function

    public function crc() as integer
        Return Me._crc
     end function



    REM --- (end of generated code: YFileRecord implementation)

    Protected _name As String
    Protected _crc As Integer
    Protected _size As Integer

    Public Sub New(ByVal data As String)
      Dim p As TJsonParser
      Dim node As Nullable(Of TJSONRECORD)
      p = New TJsonParser(data, False)
      node = p.GetChildNode(Nothing, "name")
      Me._name = node.Value.svalue
      node = p.GetChildNode(Nothing, "size")
      Me._size = node.Value.ivalue
      node = p.GetChildNode(Nothing, "crc")
      Me._crc = node.Value.ivalue
    End Sub


  End Class





  REM --- (generated code: YFiles implementation)

  Private _FilesCache As New Hashtable()
  Private _callback As UpdateCallback

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
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const FILESCOUNT_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const FREESPACE_INVALID As Integer = YAPI.INVALID_UNSIGNED

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _filesCount As Long
    Protected _freeSpace As Long

    Public Sub New(ByVal func As String)
      MyBase.new("Files", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _filesCount = Y_FILESCOUNT_INVALID
      _freeSpace = Y_FREESPACE_INVALID
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
        ElseIf (member.name = "filesCount") Then
          _filesCount = CLng(member.ivalue)
        ElseIf (member.name = "freeSpace") Then
          _freeSpace = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the filesystem.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the filesystem
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
    '''   Changes the logical name of the filesystem.
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
    '''   a string corresponding to the logical name of the filesystem
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
    '''   Returns the current value of the filesystem (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the filesystem (no more than 6 characters)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_FILESCOUNT_INVALID
        End If
      End If
      Return CType(_filesCount,Integer)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_FREESPACE_INVALID
        End If
      End If
      Return CType(_freeSpace,Integer)
    End Function
    public function sendCommand(command as string) as byte()
        dim  url as string
        url =  "files.json?a="+command
        Return Me._download(url)
        
     end function

    '''*
    ''' <summary>
    '''   Reinitializes the filesystem to its clean, unfragmented, empty state.
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
    public function format_fs() as integer
        dim  json as byte()
        dim  res as string
        json = Me.sendCommand("format")
        res  = Me._json_get_key(json, "res")
        if not(res = "ok") then
me._throw( YAPI.IO_ERROR, "format failed")
 return  YAPI.IO_ERROR
end if

        Return YAPI.SUCCESS
     end function

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
    public function get_list(pattern as string) as List(of YFileRecord)
        dim  json as byte()
        dim  list as string()
        dim  res as List(of YFileRecord) = new List(of YFileRecord)()
        json = Me.sendCommand("dir&f="+pattern)
        list = Me._json_get_array(json)
        dim i_i as integer
for i_i=0 to list.Count-1
res.Add(new YFileRecord(list(i_i)))
next i_i

        Return res
     end function

    '''*
    ''' <summary>
    '''   Downloads the requested file and returns a binary buffer with its content.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="pathname">
    '''   path and name of the new file to load
    ''' </param>
    ''' <returns>
    '''   a binary buffer with the file content
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty content.
    ''' </para>
    '''/
    public function download(pathname as string) as byte()
        Return Me._download(pathname)
        
     end function

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
    public function upload(pathname as string, content as byte()) as integer
        Return Me._upload(pathname,content)
        
     end function

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
    public function remove(pathname as string) as integer
        dim  json as byte()
        dim  res as string
        json = Me.sendCommand("del&f="+pathname)
        res  = Me._json_get_key(json, "res")
        if not(res = "ok") then
me._throw( YAPI.IO_ERROR, "unable to remove file")
 return  YAPI.IO_ERROR
end if

        Return YAPI.SUCCESS
     end function


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
    Public Function nextFiles() as YFiles
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindFiles(hwid)
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
    Public Shared Function FindFiles(ByVal func As String) As YFiles
      Dim res As YFiles
      If (_FilesCache.ContainsKey(func)) Then
        Return CType(_FilesCache(func), YFiles)
      End If
      res = New YFiles(func)
      _FilesCache.Add(func, res)
      Return res
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

    REM --- (end of generated code: YFiles implementation)

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

  Private Sub _FilesCleanup()
  End Sub


  REM --- (end of generated code: Files functions)

End Module
