'*********************************************************************
'*
'* $Id: yocto_messagebox.vb 43580 2021-01-26 17:46:01Z mvuilleu $
'*
'* Implements yFindMessageBox(), the high-level API for MessageBox functions
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

Module yocto_messagebox


    REM --- (generated code: YSms return codes)
    REM --- (end of generated code: YSms return codes)
    REM --- (generated code: YSms dlldef)
    REM --- (end of generated code: YSms dlldef)
  REM --- (generated code: YSms globals)

  REM --- (end of generated code: YSms globals)

  REM --- (generated code: YSms class start)

  '''*
  ''' <c>YSms</c> objects are used to describe an SMS message, received or to be sent.
  ''' These objects are used in particular in conjunction with the <c>YMessageBox</c> class.
  '''/
  Public Class YSms
    REM --- (end of generated code: YSms class start)

    REM --- (generated code: YSms definitions)
    REM --- (end of generated code: YSms definitions)

    REM --- (generated code: YSms attributes declaration)
    Protected _mbox As YMessageBox
    Protected _slot As Integer
    Protected _deliv As Boolean
    Protected _smsc As String
    Protected _mref As Integer
    Protected _orig As String
    Protected _dest As String
    Protected _pid As Integer
    Protected _alphab As Integer
    Protected _mclass As Integer
    Protected _stamp As String
    Protected _udh As Byte()
    Protected _udata As Byte()
    Protected _npdu As Integer
    Protected _pdu As Byte()
    Protected _parts As List(Of YSms)
    Protected _aggSig As String
    Protected _aggIdx As Integer
    Protected _aggCnt As Integer
    REM --- (end of generated code: YSms attributes declaration)

    Public Sub New(ByVal mbox As YMessageBox)
      REM --- (generated code: YSms attributes initialization)
      _slot = 0
      _mref = 0
      _pid = 0
      _alphab = 0
      _mclass = 0
      _npdu = 0
      _parts = New List(Of YSms)()
      _aggIdx = 0
      _aggCnt = 0
      REM --- (end of generated code: YSms attributes initialization)
      _mbox = mbox
    End Sub

    REM --- (generated code: YSms private methods declaration)

    REM --- (end of generated code: YSms private methods declaration)

    REM --- (generated code: YSms public methods declaration)
    Public Overridable Function get_slot() As Integer
      Return Me._slot
    End Function

    Public Overridable Function get_smsc() As String
      Return Me._smsc
    End Function

    Public Overridable Function get_msgRef() As Integer
      Return Me._mref
    End Function

    Public Overridable Function get_sender() As String
      Return Me._orig
    End Function

    Public Overridable Function get_recipient() As String
      Return Me._dest
    End Function

    Public Overridable Function get_protocolId() As Integer
      Return Me._pid
    End Function

    Public Overridable Function isReceived() As Boolean
      Return Me._deliv
    End Function

    Public Overridable Function get_alphabet() As Integer
      Return Me._alphab
    End Function

    Public Overridable Function get_msgClass() As Integer
      If (((Me._mclass) And (16)) = 0) Then
        Return -1
      End If
      Return ((Me._mclass) And (3))
    End Function

    Public Overridable Function get_dcs() As Integer
      Return ((Me._mclass) Or ((((Me._alphab) << (2)))))
    End Function

    Public Overridable Function get_timestamp() As String
      Return Me._stamp
    End Function

    Public Overridable Function get_userDataHeader() As Byte()
      Return Me._udh
    End Function

    Public Overridable Function get_userData() As Byte()
      Return Me._udata
    End Function

    '''*
    ''' <summary>
    '''   Returns the content of the message.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with the content of the message.
    ''' </returns>
    '''/
    Public Overridable Function get_textData() As String
      Dim isolatin As Byte()
      Dim isosize As Integer = 0
      Dim i As Integer = 0
      If (Me._alphab = 0) Then
        REM // using GSM standard 7-bit alphabet
        Return Me._mbox.gsm2str(Me._udata)
      End If
      If (Me._alphab = 2) Then
        REM // using UCS-2 alphabet
        isosize = (((Me._udata).Length) >> (1))
        ReDim isolatin(isosize-1)
        i = 0
        While (i < isosize)
          isolatin( i) = Convert.ToByte(Me._udata(2*i+1) And &HFF)
          i = i + 1
        End While
        Return YAPI.DefaultEncoding.GetString(isolatin)
      End If
      REM // default: convert 8 bit to string as-is
      Return YAPI.DefaultEncoding.GetString(Me._udata)
    End Function

    Public Overridable Function get_unicodeData() As List(Of Integer)
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim unisize As Integer = 0
      Dim unival As Integer = 0
      Dim i As Integer = 0
      If (Me._alphab = 0) Then
        REM // using GSM standard 7-bit alphabet
        Return Me._mbox.gsm2unicode(Me._udata)
      End If
      If (Me._alphab = 2) Then
        REM // using UCS-2 alphabet
        unisize = (((Me._udata).Length) >> (1))
        res.Clear()
        i = 0
        While (i < unisize)
          unival = 256*Me._udata(2*i)+Me._udata(2*i+1)
          res.Add(unival)
          i = i + 1
        End While
      Else
        REM // return straight 8-bit values
        unisize = (Me._udata).Length
        res.Clear()
        i = 0
        While (i < unisize)
          res.Add(Me._udata(i)+0)
          i = i + 1
        End While
      End If

      Return res
    End Function

    Public Overridable Function get_partCount() As Integer
      If (Me._npdu = 0) Then
        Me.generatePdu()
      End If
      Return Me._npdu
    End Function

    Public Overridable Function get_pdu() As Byte()
      If (Me._npdu = 0) Then
        Me.generatePdu()
      End If
      Return Me._pdu
    End Function

    Public Overridable Function get_parts() As List(Of YSms)
      If (Me._npdu = 0) Then
        Me.generatePdu()
      End If
      Return Me._parts
    End Function

    Public Overridable Function get_concatSignature() As String
      If (Me._npdu = 0) Then
        Me.generatePdu()
      End If
      Return Me._aggSig
    End Function

    Public Overridable Function get_concatIndex() As Integer
      If (Me._npdu = 0) Then
        Me.generatePdu()
      End If
      Return Me._aggIdx
    End Function

    Public Overridable Function get_concatCount() As Integer
      If (Me._npdu = 0) Then
        Me.generatePdu()
      End If
      Return Me._aggCnt
    End Function

    Public Overridable Function set_slot(val As Integer) As Integer
      Me._slot = val
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_received(val As Boolean) As Integer
      Me._deliv = val
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_smsc(val As String) As Integer
      Me._smsc = val
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_msgRef(val As Integer) As Integer
      Me._mref = val
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_sender(val As String) As Integer
      Me._orig = val
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_recipient(val As String) As Integer
      Me._dest = val
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_protocolId(val As Integer) As Integer
      Me._pid = val
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_alphabet(val As Integer) As Integer
      Me._alphab = val
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_msgClass(val As Integer) As Integer
      If (val = -1) Then
        Me._mclass = 0
      Else
        Me._mclass = 16+val
      End If
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_dcs(val As Integer) As Integer
      Me._alphab = (((((val) >> (2)))) And (3))
      Me._mclass = ((val) And (16+3))
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_timestamp(val As String) As Integer
      Me._stamp = val
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_userDataHeader(val As Byte()) As Integer
      Me._udh = val
      Me._npdu = 0
      Me.parseUserDataHeader()
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_userData(val As Byte()) As Integer
      Me._udata = val
      Me._npdu = 0
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function convertToUnicode() As Integer
      Dim ucs2 As List(Of Integer) = New List(Of Integer)()
      Dim udatalen As Integer = 0
      Dim i As Integer = 0
      Dim uni As Integer = 0
      If (Me._alphab = 2) Then
        Return YAPI.SUCCESS
      End If
      If (Me._alphab = 0) Then
        ucs2 = Me._mbox.gsm2unicode(Me._udata)
      Else
        udatalen = (Me._udata).Length
        ucs2.Clear()
        i = 0
        While (i < udatalen)
          uni = Me._udata(i)
          ucs2.Add(uni)
          i = i + 1
        End While
      End If
      Me._alphab = 2
      ReDim Me._udata(0-1)
      Me.addUnicodeData(ucs2)
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Add a regular text to the SMS.
    ''' <para>
    '''   This function support messages
    '''   of more than 160 characters. ISO-latin accented characters
    '''   are supported. For messages with special unicode characters such as asian
    '''   characters and emoticons, use the  <c>addUnicodeData</c> method.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="val">
    '''   the text to be sent in the message
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    '''/
    Public Overridable Function addText(val As String) As Integer
      Dim udata As Byte()
      Dim udatalen As Integer = 0
      Dim newdata As Byte()
      Dim newdatalen As Integer = 0
      Dim i As Integer = 0
      If ((val).Length = 0) Then
        Return YAPI.SUCCESS
      End If
      If (Me._alphab = 0) Then
        REM // Try to append using GSM 7-bit alphabet
        newdata = Me._mbox.str2gsm(val)
        newdatalen = (newdata).Length
        If (newdatalen = 0) Then
          REM // 7-bit not possible, switch to unicode
          Me.convertToUnicode()
          newdata = YAPI.DefaultEncoding.GetBytes(val)
          newdatalen = (newdata).Length
        End If
      Else
        newdata = YAPI.DefaultEncoding.GetBytes(val)
        newdatalen = (newdata).Length
      End If
      udatalen = (Me._udata).Length
      If (Me._alphab = 2) Then
        REM // Append in unicode directly
        ReDim udata(udatalen + 2*newdatalen-1)
        i = 0
        While (i < udatalen)
          udata( i) = Convert.ToByte(Me._udata(i) And &HFF)
          i = i + 1
        End While
        i = 0
        While (i < newdatalen)
          udata( udatalen+1) = Convert.ToByte(newdata(i) And &HFF)
          udatalen = udatalen + 2
          i = i + 1
        End While
      Else
        REM // Append binary buffers
        ReDim udata(udatalen+newdatalen-1)
        i = 0
        While (i < udatalen)
          udata( i) = Convert.ToByte(Me._udata(i) And &HFF)
          i = i + 1
        End While
        i = 0
        While (i < newdatalen)
          udata( udatalen) = Convert.ToByte(newdata(i) And &HFF)
          udatalen = udatalen + 1
          i = i + 1
        End While
      End If
      Return Me.set_userData(udata)
    End Function

    '''*
    ''' <summary>
    '''   Add a unicode text to the SMS.
    ''' <para>
    '''   This function support messages
    '''   of more than 160 characters, using SMS concatenation.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="val">
    '''   an array of special unicode characters
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    '''/
    Public Overridable Function addUnicodeData(val As List(Of Integer)) As Integer
      Dim arrlen As Integer = 0
      Dim newdatalen As Integer = 0
      Dim i As Integer = 0
      Dim uni As Integer = 0
      Dim udata As Byte()
      Dim udatalen As Integer = 0
      Dim surrogate As Integer = 0
      If (Me._alphab <> 2) Then
        Me.convertToUnicode()
      End If
      REM // compute number of 16-bit code units
      arrlen = val.Count
      newdatalen = arrlen
      i = 0
      While (i < arrlen)
        uni = val(i)
        If (uni > 65535) Then
          newdatalen = newdatalen + 1
        End If
        i = i + 1
      End While
      REM // now build utf-16 buffer
      udatalen = (Me._udata).Length
      ReDim udata(udatalen+2*newdatalen-1)
      i = 0
      While (i < udatalen)
        udata( i) = Convert.ToByte(Me._udata(i) And &HFF)
        i = i + 1
      End While
      i = 0
      While (i < arrlen)
        uni = val(i)
        If (uni >= 65536) Then
          surrogate = uni - 65536
          uni = (((((surrogate) >> (10))) And (1023))) + 55296
          udata( udatalen) = Convert.ToByte(((uni) >> (8)) And &HFF)
          udata( udatalen+1) = Convert.ToByte(((uni) And (255)) And &HFF)
          udatalen = udatalen + 2
          uni = (((surrogate) And (1023))) + 56320
        End If
        udata( udatalen) = Convert.ToByte(((uni) >> (8)) And &HFF)
        udata( udatalen+1) = Convert.ToByte(((uni) And (255)) And &HFF)
        udatalen = udatalen + 2
        i = i + 1
      End While
      Return Me.set_userData(udata)
    End Function

    Public Overridable Function set_pdu(pdu As Byte()) As Integer
      Me._pdu = pdu
      Me._npdu = 1
      Return Me.parsePdu(pdu)
    End Function

    Public Overridable Function set_parts(parts As List(Of YSms)) As Integer
      Dim sorted As List(Of YSms) = New List(Of YSms)()
      Dim partno As Integer = 0
      Dim initpartno As Integer = 0
      Dim i As Integer = 0
      Dim retcode As Integer = 0
      Dim totsize As Integer = 0
      Dim subsms As YSms
      Dim subdata As Byte()
      Dim res As Byte()
      Me._npdu = parts.Count
      If (Me._npdu = 0) Then
        Return YAPI.INVALID_ARGUMENT
      End If
      sorted.Clear()
      partno = 0
      While (partno < Me._npdu)
        initpartno = partno
        i = 0
        While (i < Me._npdu)
          subsms = parts(i)
          If (subsms.get_concatIndex() = partno) Then
            sorted.Add(subsms)
            partno = partno + 1
          End If
          i = i + 1
        End While
        If (initpartno = partno) Then
          partno = partno + 1
        End If
      End While

      Me._parts = sorted
      Me._npdu = sorted.Count
      REM // inherit header fields from first part
      subsms = Me._parts(0)
      retcode = Me.parsePdu(subsms.get_pdu())
      If (retcode <> YAPI.SUCCESS) Then
        Return retcode
      End If
      REM // concatenate user data from all parts
      totsize = 0
      partno = 0
      While (partno < Me._parts.Count)
        subsms = Me._parts(partno)
        subdata = subsms.get_userData()
        totsize = totsize + (subdata).Length
        partno = partno + 1
      End While
      ReDim res(totsize-1)
      totsize = 0
      partno = 0
      While (partno < Me._parts.Count)
        subsms = Me._parts(partno)
        subdata = subsms.get_userData()
        i = 0
        While (i < (subdata).Length)
          res( totsize) = Convert.ToByte(subdata(i) And &HFF)
          totsize = totsize + 1
          i = i + 1
        End While
        partno = partno + 1
      End While
      Me._udata = res
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function encodeAddress(addr As String) As Byte()
      Dim bytes As Byte()
      Dim srclen As Integer = 0
      Dim numlen As Integer = 0
      Dim i As Integer = 0
      Dim val As Integer = 0
      Dim digit As Integer = 0
      Dim res As Byte()
      bytes = YAPI.DefaultEncoding.GetBytes(addr)
      srclen = (bytes).Length
      numlen = 0
      i = 0
      While (i < srclen)
        val = bytes(i)
        If ((val >= 48) AndAlso (val < 58)) Then
          numlen = numlen + 1
        End If
        i = i + 1
      End While
      If (numlen = 0) Then
        ReDim res(1-1)
        res( 0) = Convert.ToByte(0 And &HFF)
        Return res
      End If
      ReDim res(2+((numlen+1) >> (1))-1)
      res( 0) = Convert.ToByte(numlen And &HFF)
      If (bytes(0) = 43) Then
        res( 1) = Convert.ToByte(145 And &HFF)
      Else
        res( 1) = Convert.ToByte(129 And &HFF)
      End If
      numlen = 4
      digit = 0
      i = 0
      While (i < srclen)
        val = bytes(i)
        If ((val >= 48) AndAlso (val < 58)) Then
          If (((numlen) And (1)) = 0) Then
            digit = val - 48
          Else
            res( ((numlen) >> (1))) = Convert.ToByte(digit + 16*(val-48) And &HFF)
          End If
          numlen = numlen + 1
        End If
        i = i + 1
      End While
      REM // pad with F if needed
      If (((numlen) And (1)) <> 0) Then
        res( ((numlen) >> (1))) = Convert.ToByte(digit + 240 And &HFF)
      End If
      Return res
    End Function

    Public Overridable Function decodeAddress(addr As Byte(), ofs As Integer, siz As Integer) As String
      Dim addrType As Integer = 0
      Dim gsm7 As Byte()
      Dim res As String
      Dim i As Integer = 0
      Dim rpos As Integer = 0
      Dim carry As Integer = 0
      Dim nbits As Integer = 0
      Dim byt As Integer = 0
      If (siz = 0) Then
        Return ""
      End If
      res = ""
      addrType = ((addr(ofs)) And (112))
      If (addrType = 80) Then
        REM // alphanumeric number
        siz = (4*siz \ 7)
        ReDim gsm7(siz-1)
        rpos = 1
        carry = 0
        nbits = 0
        i = 0
        While (i < siz)
          If (nbits = 7) Then
            gsm7( i) = Convert.ToByte(carry And &HFF)
            carry = 0
            nbits = 0
          Else
            byt = addr(ofs+rpos)
            rpos = rpos + 1
            gsm7( i) = Convert.ToByte(((carry) Or ((((((byt) << (nbits)))) And (127)))) And &HFF)
            carry = ((byt) >> ((7 - nbits)))
            nbits = nbits + 1
          End If
          i = i + 1
        End While
        Return Me._mbox.gsm2str(gsm7)
      Else
        REM // standard phone number
        If (addrType = 16) Then
          res = "+"
        End If
        siz = (((siz+1)) >> (1))
        i = 0
        While (i < siz)
          byt = addr(ofs+i+1)
          res = "" +  res + "" + ( ((byt) And (15))).ToString("x") + "" + (((byt) >> (4))).ToString("x")
          i = i + 1
        End While
        REM // remove padding digit if needed
        If (((addr(ofs+siz)) >> (4)) = 15) Then
          res = (res).Substring( 0, (res).Length-1)
        End If
        Return res
      End If
    End Function

    Public Overridable Function encodeTimeStamp(exp As String) As Byte()
      Dim explen As Integer = 0
      Dim i As Integer = 0
      Dim res As Byte()
      Dim n As Integer = 0
      Dim expasc As Byte()
      Dim v1 As Integer = 0
      Dim v2 As Integer = 0
      explen = (exp).Length
      If (explen = 0) Then
        ReDim res(0-1)
        Return res
      End If
      If ((exp).Substring(0, 1) = "+") Then
        n = YAPI._atoi((exp).Substring(1, explen-1))
        ReDim res(1-1)
        If (n > 30*86400) Then
          n = 192+((n+6*86400) \ (7*86400))
        Else
          If (n > 86400) Then
            n = 166+((n+86399) \ 86400)
          Else
            If (n > 43200) Then
              n = 143+((n-43200+1799) \ 1800)
            Else
              n = -1+((n+299) \ 300)
            End If
          End If
        End If
        If (n < 0) Then
          n = 0
        End If
        res(0) = Convert.ToByte(n And &HFF)
        Return res
      End If
      If ((exp).Substring(4, 1) = "-" OrElse (exp).Substring(4, 1) = "/") Then
        REM // ignore century
        exp = (exp).Substring( 2, explen-2)
        explen = (exp).Length
      End If
      expasc = YAPI.DefaultEncoding.GetBytes(exp)
      ReDim res(7-1)
      n = 0
      i = 0
      While ((i+1 < explen) AndAlso (n < 7))
        v1 = expasc(i)
        If ((v1 >= 48) AndAlso (v1 < 58)) Then
          v2 = expasc(i+1)
          If ((v2 >= 48) AndAlso (v2 < 58)) Then
            v1 = v1 - 48
            v2 = v2 - 48
            res( n) = Convert.ToByte((((v2) << (4))) + v1 And &HFF)
            n = n + 1
            i = i + 1
          End If
        End If
        i = i + 1
      End While
      While (n < 7)
        res( n) = Convert.ToByte(0 And &HFF)
        n = n + 1
      End While
      If (i+2 < explen) Then
        REM // convert for timezone in cleartext ISO format +/-nn:nn
        v1 = expasc(i-3)
        v2 = expasc(i)
        If (((v1 = 43) OrElse (v1 = 45)) AndAlso (v2 = 58)) Then
          v1 = expasc(i+1)
          v2 = expasc(i+2)
          If ((v1 >= 48) AndAlso (v1 < 58) AndAlso (v1 >= 48) AndAlso (v1 < 58)) Then
            v1 = ((10*(v1 - 48)+(v2 - 48)) \ 15)
            n = n - 1
            v2 = 4 * res(n) + v1
            If (expasc(i-3) = 45) Then
              v2 += 128
            End If
            res( n) = Convert.ToByte(v2 And &HFF)
          End If
        End If
      End If
      Return res
    End Function

    Public Overridable Function decodeTimeStamp(exp As Byte(), ofs As Integer, siz As Integer) As String
      Dim n As Integer = 0
      Dim res As String
      Dim i As Integer = 0
      Dim byt As Integer = 0
      Dim sign As String
      Dim hh As String
      Dim ss As String
      If (siz < 1) Then
        Return ""
      End If
      If (siz = 1) Then
        n = exp(ofs)
        If (n < 144) Then
          n = n * 300
        Else
          If (n < 168) Then
            n = (n-143) * 1800
          Else
            If (n < 197) Then
              n = (n-166) * 86400
            Else
              n = (n-192) * 7 * 86400
            End If
          End If
        End If
        Return "+" + Convert.ToString(n)
      End If
      res = "20"
      i = 0
      While ((i < siz) AndAlso (i < 6))
        byt = exp(ofs+i)
        res = "" +  res + "" + ( ((byt) And (15))).ToString("x") + "" + (((byt) >> (4))).ToString("x")
        If (i < 3) Then
          If (i < 2) Then
            res = "" + res + "-"
          Else
            res = "" + res + " "
          End If
        Else
          If (i < 5) Then
            res = "" + res + ":"
          End If
        End If
        i = i + 1
      End While
      If (siz = 7) Then
        byt = exp(ofs+i)
        sign = "+"
        If (((byt) And (8)) <> 0) Then
          byt = byt - 8
          sign = "-"
        End If
        byt = (10*(((byt) And (15)))) + (((byt) >> (4)))
        hh = "" + Convert.ToString(((byt) >> (2)))
        ss = "" + Convert.ToString(15*(((byt) And (3))))
        If ((hh).Length<2) Then
          hh = "0" + hh
        End If
        If ((ss).Length<2) Then
          ss = "0" + ss
        End If
        res = "" +  res + "" +  sign + "" +  hh + ":" + ss
      End If
      Return res
    End Function

    Public Overridable Function udataSize() As Integer
      Dim res As Integer = 0
      Dim udhsize As Integer = 0
      udhsize = (Me._udh).Length
      res = (Me._udata).Length
      If (Me._alphab = 0) Then
        If (udhsize > 0) Then
          res = res + ((8 + 8*udhsize + 6) \ 7)
        End If
        res = ((res * 7 + 7) \ 8)
      Else
        If (udhsize > 0) Then
          res = res + 1 + udhsize
        End If
      End If
      Return res
    End Function

    Public Overridable Function encodeUserData() As Byte()
      Dim udsize As Integer = 0
      Dim udlen As Integer = 0
      Dim udhsize As Integer = 0
      Dim udhlen As Integer = 0
      Dim res As Byte()
      Dim i As Integer = 0
      Dim wpos As Integer = 0
      Dim carry As Integer = 0
      Dim nbits As Integer = 0
      Dim thi_b As Integer = 0
      REM // nbits = number of bits in carry
      udsize = Me.udataSize()
      udhsize = (Me._udh).Length
      udlen = (Me._udata).Length
      ReDim res(1+udsize-1)
      udhlen = 0
      nbits = 0
      carry = 0
      REM // 1. Encode UDL
      If (Me._alphab = 0) Then
        REM // 7-bit encoding
        If (udhsize > 0) Then
          udhlen = ((8 + 8*udhsize + 6) \ 7)
          nbits = 7*udhlen - 8 - 8*udhsize
        End If
        res( 0) = Convert.ToByte(udhlen+udlen And &HFF)
      Else
        REM // 8-bit encoding
        res( 0) = Convert.ToByte(udsize And &HFF)
      End If
      REM // 2. Encode UDHL and UDL
      wpos = 1
      If (udhsize > 0) Then
        res( wpos) = Convert.ToByte(udhsize And &HFF)
        wpos = wpos + 1
        i = 0
        While (i < udhsize)
          res( wpos) = Convert.ToByte(Me._udh(i) And &HFF)
          wpos = wpos + 1
          i = i + 1
        End While
      End If
      REM // 3. Encode UD
      If (Me._alphab = 0) Then
        REM // 7-bit encoding
        i = 0
        While (i < udlen)
          If (nbits = 0) Then
            carry = Me._udata(i)
            nbits = 7
          Else
            thi_b = Me._udata(i)
            res( wpos) = Convert.ToByte(((carry) Or ((((((thi_b) << (nbits)))) And (255)))) And &HFF)
            wpos = wpos + 1
            nbits = nbits - 1
            carry = ((thi_b) >> ((7 - nbits)))
          End If
          i = i + 1
        End While
        If (nbits > 0) Then
          res( wpos) = Convert.ToByte(carry And &HFF)
        End If
      Else
        REM // 8-bit encoding
        i = 0
        While (i < udlen)
          res( wpos) = Convert.ToByte(Me._udata(i) And &HFF)
          wpos = wpos + 1
          i = i + 1
        End While
      End If
      Return res
    End Function

    Public Overridable Function generateParts() As Integer
      Dim udhsize As Integer = 0
      Dim udlen As Integer = 0
      Dim mss As Integer = 0
      Dim partno As Integer = 0
      Dim partlen As Integer = 0
      Dim newud As Byte()
      Dim newudh As Byte()
      Dim newpdu As YSms
      Dim i As Integer = 0
      Dim wpos As Integer = 0
      udhsize = (Me._udh).Length
      udlen = (Me._udata).Length
      mss = 140 - 1 - 5 - udhsize
      If (Me._alphab = 0) Then
        mss = ((mss * 8 - 6) \ 7)
      End If
      Me._npdu = ((udlen+mss-1) \ mss)
      Me._parts.Clear()
      partno = 0
      wpos = 0
      While (wpos < udlen)
        partno = partno + 1
        ReDim newudh(5+udhsize-1)
        newudh( 0) = Convert.ToByte(0 And &HFF)
        REM // IEI: concatenated message
        newudh( 1) = Convert.ToByte(3 And &HFF)
        REM // IEDL: 3 bytes
        newudh( 2) = Convert.ToByte(Me._mref And &HFF)
        newudh( 3) = Convert.ToByte(Me._npdu And &HFF)
        newudh( 4) = Convert.ToByte(partno And &HFF)
        i = 0
        While (i < udhsize)
          newudh( 5+i) = Convert.ToByte(Me._udh(i) And &HFF)
          i = i + 1
        End While
        If (wpos+mss < udlen) Then
          partlen = mss
        Else
          partlen = udlen-wpos
        End If
        ReDim newud(partlen-1)
        i = 0
        While (i < partlen)
          newud( i) = Convert.ToByte(Me._udata(wpos) And &HFF)
          wpos = wpos + 1
          i = i + 1
        End While
        newpdu = New YSms(Me._mbox)
        newpdu.set_received(Me.isReceived())
        newpdu.set_smsc(Me.get_smsc())
        newpdu.set_msgRef(Me.get_msgRef())
        newpdu.set_sender(Me.get_sender())
        newpdu.set_recipient(Me.get_recipient())
        newpdu.set_protocolId(Me.get_protocolId())
        newpdu.set_dcs(Me.get_dcs())
        newpdu.set_timestamp(Me.get_timestamp())
        newpdu.set_userDataHeader(newudh)
        newpdu.set_userData(newud)
        Me._parts.Add(newpdu)
      End While
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function generatePdu() As Integer
      Dim sca As Byte()
      Dim hdr As Byte()
      Dim addr As Byte()
      Dim stamp As Byte()
      Dim udata As Byte()
      Dim pdutyp As Integer = 0
      Dim pdulen As Integer = 0
      Dim i As Integer = 0
      REM // Determine if the message can fit within a single PDU
      Me._parts.Clear()
      If (Me.udataSize() > 140) Then
        REM // multiple PDU are needed
        ReDim Me._pdu(0-1)
        Return Me.generateParts()
      End If
      sca = Me.encodeAddress(Me._smsc)
      If ((sca).Length > 0) Then
        sca( 0) = Convert.ToByte((sca).Length-1 And &HFF)
      End If
      stamp = Me.encodeTimeStamp(Me._stamp)
      udata = Me.encodeUserData()
      If (Me._deliv) Then
        addr = Me.encodeAddress(Me._orig)
        ReDim hdr(1-1)
        pdutyp = 0
      Else
        addr = Me.encodeAddress(Me._dest)
        Me._mref = Me._mbox.nextMsgRef()
        ReDim hdr(2-1)
        hdr(1) = Convert.ToByte(Me._mref And &HFF)
        pdutyp = 1
        If ((stamp).Length > 0) Then
          pdutyp = pdutyp + 16
        End If
        If ((stamp).Length = 7) Then
          pdutyp = pdutyp + 8
        End If
      End If
      If ((Me._udh).Length > 0) Then
        pdutyp = pdutyp + 64
      End If
      hdr(0) = Convert.ToByte(pdutyp And &HFF)
      pdulen = (sca).Length+(hdr).Length+(addr).Length+2+(stamp).Length+(udata).Length
      ReDim Me._pdu(pdulen-1)
      pdulen = 0
      i = 0
      While (i < (sca).Length)
        Me._pdu( pdulen) = Convert.ToByte(sca(i) And &HFF)
        pdulen = pdulen + 1
        i = i + 1
      End While
      i = 0
      While (i < (hdr).Length)
        Me._pdu( pdulen) = Convert.ToByte(hdr(i) And &HFF)
        pdulen = pdulen + 1
        i = i + 1
      End While
      i = 0
      While (i < (addr).Length)
        Me._pdu( pdulen) = Convert.ToByte(addr(i) And &HFF)
        pdulen = pdulen + 1
        i = i + 1
      End While
      Me._pdu( pdulen) = Convert.ToByte(Me._pid And &HFF)
      pdulen = pdulen + 1
      Me._pdu( pdulen) = Convert.ToByte(Me.get_dcs() And &HFF)
      pdulen = pdulen + 1
      i = 0
      While (i < (stamp).Length)
        Me._pdu( pdulen) = Convert.ToByte(stamp(i) And &HFF)
        pdulen = pdulen + 1
        i = i + 1
      End While
      i = 0
      While (i < (udata).Length)
        Me._pdu( pdulen) = Convert.ToByte(udata(i) And &HFF)
        pdulen = pdulen + 1
        i = i + 1
      End While
      Me._npdu = 1
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function parseUserDataHeader() As Integer
      Dim udhlen As Integer = 0
      Dim i As Integer = 0
      Dim iei As Integer = 0
      Dim ielen As Integer = 0
      Dim sig As String
      Me._aggSig = ""
      Me._aggIdx = 0
      Me._aggCnt = 0
      udhlen = (Me._udh).Length
      i = 0
      While (i+1 < udhlen)
        iei = Me._udh(i)
        ielen = Me._udh(i+1)
        i = i + 2
        If (i + ielen <= udhlen) Then
          If ((iei = 0) AndAlso (ielen = 3)) Then
            REM // concatenated SMS, 8-bit ref
            sig = "" +  Me._orig + "-" +  Me._dest + "-" + (
            Me._mref).ToString("x02") + "-" + (Me._udh(i)).ToString("x02")
            Me._aggSig = sig
            Me._aggCnt = Me._udh(i+1)
            Me._aggIdx = Me._udh(i+2)
          End If
          If ((iei = 8) AndAlso (ielen = 4)) Then
            REM // concatenated SMS, 16-bit ref
            sig = "" +  Me._orig + "-" +  Me._dest + "-" + (
            Me._mref).ToString("x02") + "-" + ( Me._udh(i)).ToString("x02") + "" + (Me._udh(i+1)).ToString("x02")
            Me._aggSig = sig
            Me._aggCnt = Me._udh(i+2)
            Me._aggIdx = Me._udh(i+3)
          End If
        End If
        i = i + ielen
      End While
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function parsePdu(pdu As Byte()) As Integer
      Dim rpos As Integer = 0
      Dim addrlen As Integer = 0
      Dim pdutyp As Integer = 0
      Dim tslen As Integer = 0
      Dim dcs As Integer = 0
      Dim udlen As Integer = 0
      Dim udhsize As Integer = 0
      Dim udhlen As Integer = 0
      Dim i As Integer = 0
      Dim carry As Integer = 0
      Dim nbits As Integer = 0
      Dim thi_b As Integer = 0
      Me._pdu = pdu
      Me._npdu = 1
      REM // parse meta-data
      Me._smsc = Me.decodeAddress(pdu, 1, 2*(pdu(0)-1))
      rpos = 1+pdu(0)
      pdutyp = pdu(rpos)
      rpos = rpos + 1
      Me._deliv = (((pdutyp) And (3)) = 0)
      If (Me._deliv) Then
        addrlen = pdu(rpos)
        rpos = rpos + 1
        Me._orig = Me.decodeAddress(pdu, rpos, addrlen)
        Me._dest = ""
        tslen = 7
      Else
        Me._mref = pdu(rpos)
        rpos = rpos + 1
        addrlen = pdu(rpos)
        rpos = rpos + 1
        Me._dest = Me.decodeAddress(pdu, rpos, addrlen)
        Me._orig = ""
        If ((((pdutyp) And (16))) <> 0) Then
          If ((((pdutyp) And (8))) <> 0) Then
            tslen = 7
          Else
            tslen= 1
          End If
        Else
          tslen = 0
        End If
      End If
      rpos = rpos + ((((addrlen+3)) >> (1)))
      Me._pid = pdu(rpos)
      rpos = rpos + 1
      dcs = pdu(rpos)
      rpos = rpos + 1
      Me._alphab = (((((dcs) >> (2)))) And (3))
      Me._mclass = ((dcs) And (16+3))
      Me._stamp = Me.decodeTimeStamp(pdu, rpos, tslen)
      rpos = rpos + tslen
      REM // parse user data (including udh)
      nbits = 0
      carry = 0
      udlen = pdu(rpos)
      rpos = rpos + 1
      If (((pdutyp) And (64)) <> 0) Then
        udhsize = pdu(rpos)
        rpos = rpos + 1
        ReDim Me._udh(udhsize-1)
        i = 0
        While (i < udhsize)
          Me._udh( i) = Convert.ToByte(pdu(rpos) And &HFF)
          rpos = rpos + 1
          i = i + 1
        End While
        If (Me._alphab = 0) Then
          REM // 7-bit encoding
          udhlen = ((8 + 8*udhsize + 6) \ 7)
          nbits = 7*udhlen - 8 - 8*udhsize
          If (nbits > 0) Then
            thi_b = pdu(rpos)
            rpos = rpos + 1
            carry = ((thi_b) >> (nbits))
            nbits = 8 - nbits
          End If
        Else
          REM // byte encoding
          udhlen = 1+udhsize
        End If
        udlen = udlen - udhlen
      Else
        udhsize = 0
        ReDim Me._udh(0-1)
      End If
      ReDim Me._udata(udlen-1)
      If (Me._alphab = 0) Then
        REM // 7-bit encoding
        i = 0
        While (i < udlen)
          If (nbits = 7) Then
            Me._udata( i) = Convert.ToByte(carry And &HFF)
            carry = 0
            nbits = 0
          Else
            thi_b = pdu(rpos)
            rpos = rpos + 1
            Me._udata( i) = Convert.ToByte(((carry) Or ((((((thi_b) << (nbits)))) And (127)))) And &HFF)
            carry = ((thi_b) >> ((7 - nbits)))
            nbits = nbits + 1
          End If
          i = i + 1
        End While
      Else
        REM // 8-bit encoding
        i = 0
        While (i < udlen)
          Me._udata( i) = Convert.ToByte(pdu(rpos) And &HFF)
          rpos = rpos + 1
          i = i + 1
        End While
      End If
      Me.parseUserDataHeader()
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Sends the SMS to the recipient.
    ''' <para>
    '''   Messages of more than 160 characters are supported
    '''   using SMS concatenation.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function send() As Integer
      Dim i As Integer = 0
      Dim retcode As Integer = 0
      Dim pdu As YSms

      If (Me._npdu = 0) Then
        Me.generatePdu()
      End If
      If (Me._npdu = 1) Then
        Return Me._mbox._upload("sendSMS", Me._pdu)
      End If
      retcode = YAPI.SUCCESS
      i = 0
      While ((i < Me._npdu) AndAlso (retcode = YAPI.SUCCESS))
        pdu = Me._parts(i)
        retcode= pdu.send()
        i = i + 1
      End While
      Return retcode
    End Function

    Public Overridable Function deleteFromSIM() As Integer
      Dim i As Integer = 0
      Dim retcode As Integer = 0
      Dim pdu As YSms

      If (Me._slot > 0) Then
        Return Me._mbox.clearSIMSlot(Me._slot)
      End If
      retcode = YAPI.SUCCESS
      i = 0
      While ((i < Me._npdu) AndAlso (retcode = YAPI.SUCCESS))
        pdu = Me._parts(i)
        retcode= pdu.deleteFromSIM()
        i = i + 1
      End While
      Return retcode
    End Function



    REM --- (end of generated code: YSms public methods declaration)

  End Class

  REM --- (generated code: YSms functions)


  REM --- (end of generated code: YSms functions)



    REM --- (generated code: YMessageBox return codes)
    REM --- (end of generated code: YMessageBox return codes)
    REM --- (generated code: YMessageBox dlldef)
    REM --- (end of generated code: YMessageBox dlldef)
  REM --- (generated code: YMessageBox globals)

  Public Const Y_SLOTSINUSE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_SLOTSCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_SLOTSBITMAP_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PDUSENT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PDURECEIVED_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YMessageBoxValueCallback(ByVal func As YMessageBox, ByVal value As String)
  Public Delegate Sub YMessageBoxTimedReportCallback(ByVal func As YMessageBox, ByVal measure As YMeasure)
  REM --- (end of generated code: YMessageBox globals)

  REM --- (generated code: YMessageBox class start)

  '''*
  ''' <summary>
  '''   The <c>YMessageBox</c> class provides SMS sending and receiving capability for
  '''   GSM-enabled Yoctopuce devices.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMessageBox
    Inherits YFunction
    REM --- (end of generated code: YMessageBox class start)

    REM --- (generated code: YMessageBox definitions)
    Public Const SLOTSINUSE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const SLOTSCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const SLOTSBITMAP_INVALID As String = YAPI.INVALID_STRING
    Public Const PDUSENT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PDURECEIVED_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of generated code: YMessageBox definitions)

    REM --- (generated code: YMessageBox attributes declaration)
    Protected _slotsInUse As Integer
    Protected _slotsCount As Integer
    Protected _slotsBitmap As String
    Protected _pduSent As Integer
    Protected _pduReceived As Integer
    Protected _command As String
    Protected _valueCallbackMessageBox As YMessageBoxValueCallback
    Protected _nextMsgRef As Integer
    Protected _prevBitmapStr As String
    Protected _pdus As List(Of YSms)
    Protected _messages As List(Of YSms)
    Protected _gsm2unicodeReady As Boolean
    Protected _gsm2unicode As List(Of Integer)
    Protected _iso2gsm As Byte()
    REM --- (end of generated code: YMessageBox attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "MessageBox"
      REM --- (generated code: YMessageBox attributes initialization)
      _slotsInUse = SLOTSINUSE_INVALID
      _slotsCount = SLOTSCOUNT_INVALID
      _slotsBitmap = SLOTSBITMAP_INVALID
      _pduSent = PDUSENT_INVALID
      _pduReceived = PDURECEIVED_INVALID
      _command = COMMAND_INVALID
      _valueCallbackMessageBox = Nothing
      _nextMsgRef = 0
      _pdus = New List(Of YSms)()
      _messages = New List(Of YSms)()
      _gsm2unicode = New List(Of Integer)()
      REM --- (end of generated code: YMessageBox attributes initialization)
    End Sub

    REM --- (generated code: YMessageBox private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("slotsInUse") Then
        _slotsInUse = CInt(json_val.getLong("slotsInUse"))
      End If
      If json_val.has("slotsCount") Then
        _slotsCount = CInt(json_val.getLong("slotsCount"))
      End If
      If json_val.has("slotsBitmap") Then
        _slotsBitmap = json_val.getString("slotsBitmap")
      End If
      If json_val.has("pduSent") Then
        _pduSent = CInt(json_val.getLong("pduSent"))
      End If
      If json_val.has("pduReceived") Then
        _pduReceived = CInt(json_val.getLong("pduReceived"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YMessageBox private methods declaration)

    REM --- (generated code: YMessageBox public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the number of message storage slots currently in use.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of message storage slots currently in use
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMessageBox.SLOTSINUSE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_slotsInUse() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SLOTSINUSE_INVALID
        End If
      End If
      res = Me._slotsInUse
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the total number of message storage slots on the SIM card.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the total number of message storage slots on the SIM card
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMessageBox.SLOTSCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_slotsCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SLOTSCOUNT_INVALID
        End If
      End If
      res = Me._slotsCount
      Return res
    End Function

    Public Function get_slotsBitmap() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SLOTSBITMAP_INVALID
        End If
      End If
      res = Me._slotsBitmap
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of SMS units sent so far.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of SMS units sent so far
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMessageBox.PDUSENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pduSent() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PDUSENT_INVALID
        End If
      End If
      res = Me._pduSent
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the value of the outgoing SMS units counter.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the value of the outgoing SMS units counter
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
    Public Function set_pduSent(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pduSent", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of SMS units received so far.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of SMS units received so far
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMessageBox.PDURECEIVED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pduReceived() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PDURECEIVED_INVALID
        End If
      End If
      res = Me._pduReceived
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the value of the incoming SMS units counter.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the value of the incoming SMS units counter
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
    Public Function set_pduReceived(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pduReceived", rest_val)
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
    '''   Retrieves a SMS message box interface for a given identifier.
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
    '''   This function does not require that the SMS message box interface is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YMessageBox.isOnline()</c> to test if the SMS message box interface is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a SMS message box interface by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the SMS message box interface, for instance
    '''   <c>YHUBGSM1.messageBox</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YMessageBox</c> object allowing you to drive the SMS message box interface.
    ''' </returns>
    '''/
    Public Shared Function FindMessageBox(func As String) As YMessageBox
      Dim obj As YMessageBox
      obj = CType(YFunction._FindFromCache("MessageBox", func), YMessageBox)
      If ((obj Is Nothing)) Then
        obj = New YMessageBox(func)
        YFunction._AddToCache("MessageBox", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YMessageBoxValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackMessageBox = callback
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
      If (Not (Me._valueCallbackMessageBox Is Nothing)) Then
        Me._valueCallbackMessageBox(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function nextMsgRef() As Integer
      Me._nextMsgRef = Me._nextMsgRef + 1
      Return Me._nextMsgRef
    End Function

    Public Overridable Function clearSIMSlot(slot As Integer) As Integer
      Me._prevBitmapStr = ""
      Return Me.set_command("DS" + Convert.ToString(slot))
    End Function

    Public Overridable Function fetchPdu(slot As Integer) As YSms
      Dim binPdu As Byte()
      Dim arrPdu As List(Of String) = New List(Of String)()
      Dim hexPdu As String
      Dim sms As YSms

      binPdu = Me._download("sms.json?pos=" + Convert.ToString(slot) + "&len=1")
      arrPdu = Me._json_get_array(binPdu)
      hexPdu = Me._decode_json_string(arrPdu(0))
      sms = New YSms(Me)
      sms.set_slot(slot)
      sms.parsePdu(YAPI._hexStrToBin(hexPdu))
      Return sms
    End Function

    Public Overridable Function initGsm2Unicode() As Integer
      Dim i As Integer = 0
      Dim uni As Integer = 0
      Me._gsm2unicode.Clear()
      REM // 00-07
      Me._gsm2unicode.Add(64)
      Me._gsm2unicode.Add(163)
      Me._gsm2unicode.Add(36)
      Me._gsm2unicode.Add(165)
      Me._gsm2unicode.Add(232)
      Me._gsm2unicode.Add(233)
      Me._gsm2unicode.Add(249)
      Me._gsm2unicode.Add(236)
      REM // 08-0F
      Me._gsm2unicode.Add(242)
      Me._gsm2unicode.Add(199)
      Me._gsm2unicode.Add(10)
      Me._gsm2unicode.Add(216)
      Me._gsm2unicode.Add(248)
      Me._gsm2unicode.Add(13)
      Me._gsm2unicode.Add(197)
      Me._gsm2unicode.Add(229)
      REM // 10-17
      Me._gsm2unicode.Add(916)
      Me._gsm2unicode.Add(95)
      Me._gsm2unicode.Add(934)
      Me._gsm2unicode.Add(915)
      Me._gsm2unicode.Add(923)
      Me._gsm2unicode.Add(937)
      Me._gsm2unicode.Add(928)
      Me._gsm2unicode.Add(936)
      REM // 18-1F
      Me._gsm2unicode.Add(931)
      Me._gsm2unicode.Add(920)
      Me._gsm2unicode.Add(926)
      Me._gsm2unicode.Add(27)
      Me._gsm2unicode.Add(198)
      Me._gsm2unicode.Add(230)
      Me._gsm2unicode.Add(223)
      Me._gsm2unicode.Add(201)
      REM // 20-7A
      i = 32
      While (i <= 122)
        Me._gsm2unicode.Add(i)
        i = i + 1
      End While
      REM // exceptions in range 20-7A
      Me._gsm2unicode( 36) = 164
      Me._gsm2unicode( 64) = 161
      Me._gsm2unicode( 91) = 196
      Me._gsm2unicode( 92) = 214
      Me._gsm2unicode( 93) = 209
      Me._gsm2unicode( 94) = 220
      Me._gsm2unicode( 95) = 167
      Me._gsm2unicode( 96) = 191
      REM // 7B-7F
      Me._gsm2unicode.Add(228)
      Me._gsm2unicode.Add(246)
      Me._gsm2unicode.Add(241)
      Me._gsm2unicode.Add(252)
      Me._gsm2unicode.Add(224)

      REM // Invert table as well wherever possible
      ReDim Me._iso2gsm(256-1)
      i = 0
      While (i <= 127)
        uni = Me._gsm2unicode(i)
        If (uni <= 255) Then
          Me._iso2gsm( uni) = Convert.ToByte(i And &HFF)
        End If
        i = i + 1
      End While
      i = 0
      While (i < 4)
        REM // mark escape sequences
        Me._iso2gsm( 91+i) = Convert.ToByte(27 And &HFF)
        Me._iso2gsm( 123+i) = Convert.ToByte(27 And &HFF)
        i = i + 1
      End While
      REM // Done
      Me._gsm2unicodeReady = True
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function gsm2unicode(gsm As Byte()) As List(Of Integer)
      Dim i As Integer = 0
      Dim gsmlen As Integer = 0
      Dim reslen As Integer = 0
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim uni As Integer = 0
      If (Not (Me._gsm2unicodeReady)) Then
        Me.initGsm2Unicode()
      End If
      gsmlen = (gsm).Length
      reslen = gsmlen
      i = 0
      While (i < gsmlen)
        If (gsm(i) = 27) Then
          reslen = reslen - 1
        End If
        i = i + 1
      End While
      res.Clear()
      i = 0
      While (i < gsmlen)
        uni = Me._gsm2unicode(gsm(i))
        If ((uni = 27) AndAlso (i+1 < gsmlen)) Then
          i = i + 1
          uni = gsm(i)
          If (uni < 60) Then
            If (uni < 41) Then
              If (uni=20) Then
                uni=94
              Else
                If (uni=40) Then
                  uni=123
                Else
                  uni=0
                End If
              End If
            Else
              If (uni=41) Then
                uni=125
              Else
                If (uni=47) Then
                  uni=92
                Else
                  uni=0
                End If
              End If
            End If
          Else
            If (uni < 62) Then
              If (uni=60) Then
                uni=91
              Else
                If (uni=61) Then
                  uni=126
                Else
                  uni=0
                End If
              End If
            Else
              If (uni=62) Then
                uni=93
              Else
                If (uni=64) Then
                  uni=124
                Else
                  If (uni=101) Then
                    uni=164
                  Else
                    uni=0
                  End If
                End If
              End If
            End If
          End If
        End If
        If (uni > 0) Then
          res.Add(uni)
        End If
        i = i + 1
      End While

      Return res
    End Function

    Public Overridable Function gsm2str(gsm As Byte()) As String
      Dim i As Integer = 0
      Dim gsmlen As Integer = 0
      Dim reslen As Integer = 0
      Dim resbin As Byte()
      Dim resstr As String
      Dim uni As Integer = 0
      If (Not (Me._gsm2unicodeReady)) Then
        Me.initGsm2Unicode()
      End If
      gsmlen = (gsm).Length
      reslen = gsmlen
      i = 0
      While (i < gsmlen)
        If (gsm(i) = 27) Then
          reslen = reslen - 1
        End If
        i = i + 1
      End While
      ReDim resbin(reslen-1)
      i = 0
      reslen = 0
      While (i < gsmlen)
        uni = Me._gsm2unicode(gsm(i))
        If ((uni = 27) AndAlso (i+1 < gsmlen)) Then
          i = i + 1
          uni = gsm(i)
          If (uni < 60) Then
            If (uni < 41) Then
              If (uni=20) Then
                uni=94
              Else
                If (uni=40) Then
                  uni=123
                Else
                  uni=0
                End If
              End If
            Else
              If (uni=41) Then
                uni=125
              Else
                If (uni=47) Then
                  uni=92
                Else
                  uni=0
                End If
              End If
            End If
          Else
            If (uni < 62) Then
              If (uni=60) Then
                uni=91
              Else
                If (uni=61) Then
                  uni=126
                Else
                  uni=0
                End If
              End If
            Else
              If (uni=62) Then
                uni=93
              Else
                If (uni=64) Then
                  uni=124
                Else
                  If (uni=101) Then
                    uni=164
                  Else
                    uni=0
                  End If
                End If
              End If
            End If
          End If
        End If
        If ((uni > 0) AndAlso (uni < 256)) Then
          resbin( reslen) = Convert.ToByte(uni And &HFF)
          reslen = reslen + 1
        End If
        i = i + 1
      End While
      resstr = YAPI.DefaultEncoding.GetString(resbin)
      If ((resstr).Length > reslen) Then
        resstr = (resstr).Substring(0, reslen)
      End If
      Return resstr
    End Function

    Public Overridable Function str2gsm(msg As String) As Byte()
      Dim asc As Byte()
      Dim asclen As Integer = 0
      Dim i As Integer = 0
      Dim ch As Integer = 0
      Dim gsm7 As Integer = 0
      Dim extra As Integer = 0
      Dim res As Byte()
      Dim wpos As Integer = 0
      If (Not (Me._gsm2unicodeReady)) Then
        Me.initGsm2Unicode()
      End If
      asc = YAPI.DefaultEncoding.GetBytes(msg)
      asclen = (asc).Length
      extra = 0
      i = 0
      While (i < asclen)
        ch = asc(i)
        gsm7 = Me._iso2gsm(ch)
        If (gsm7 = 27) Then
          extra = extra + 1
        End If
        If (gsm7 = 0) Then
          REM // cannot use standard GSM encoding
          ReDim res(0-1)
          Return res
        End If
        i = i + 1
      End While
      ReDim res(asclen+extra-1)
      wpos = 0
      i = 0
      While (i < asclen)
        ch = asc(i)
        gsm7 = Me._iso2gsm(ch)
        res( wpos) = Convert.ToByte(gsm7 And &HFF)
        wpos = wpos + 1
        If (gsm7 = 27) Then
          If (ch < 100) Then
            If (ch<93) Then
              If (ch<92) Then
                gsm7=60
              Else
                gsm7=47
              End If
            Else
              If (ch<94) Then
                gsm7=62
              Else
                gsm7=20
              End If
            End If
          Else
            If (ch<125) Then
              If (ch<124) Then
                gsm7=40
              Else
                gsm7=64
              End If
            Else
              If (ch<126) Then
                gsm7=41
              Else
                gsm7=61
              End If
            End If
          End If
          res( wpos) = Convert.ToByte(gsm7 And &HFF)
          wpos = wpos + 1
        End If
        i = i + 1
      End While
      Return res
    End Function

    Public Overridable Function checkNewMessages() As Integer
      Dim bitmapStr As String
      Dim prevBitmap As Byte()
      Dim newBitmap As Byte()
      Dim slot As Integer = 0
      Dim nslots As Integer = 0
      Dim pduIdx As Integer = 0
      Dim idx As Integer = 0
      Dim bitVal As Integer = 0
      Dim prevBit As Integer = 0
      Dim i As Integer = 0
      Dim nsig As Integer = 0
      Dim cnt As Integer = 0
      Dim sig As String
      Dim newArr As List(Of YSms) = New List(Of YSms)()
      Dim newMsg As List(Of YSms) = New List(Of YSms)()
      Dim newAgg As List(Of YSms) = New List(Of YSms)()
      Dim signatures As List(Of String) = New List(Of String)()
      Dim sms As YSms

      bitmapStr = Me.get_slotsBitmap()
      If (bitmapStr = Me._prevBitmapStr) Then
        Return YAPI.SUCCESS
      End If
      prevBitmap = YAPI._hexStrToBin(Me._prevBitmapStr)
      newBitmap = YAPI._hexStrToBin(bitmapStr)
      Me._prevBitmapStr = bitmapStr
      nslots = 8*(newBitmap).Length
      newArr.Clear()
      newMsg.Clear()
      signatures.Clear()
      nsig = 0
      REM // copy known messages
      pduIdx = 0
      While (pduIdx < Me._pdus.Count)
        sms = Me._pdus(pduIdx)
        slot = sms.get_slot()
        idx = ((slot) >> (3))
        If (idx < (newBitmap).Length) Then
          bitVal = ((1) << ((((slot) And (7)))))
          If ((((newBitmap(idx)) And (bitVal))) <> 0) Then
            newArr.Add(sms)
            If (sms.get_concatCount() = 0) Then
              newMsg.Add(sms)
            Else
              sig = sms.get_concatSignature()
              i = 0
              While ((i < nsig) AndAlso ((sig).Length > 0))
                If (signatures(i) = sig) Then
                  sig = ""
                End If
                i = i + 1
              End While
              If ((sig).Length > 0) Then
                signatures.Add(sig)
                nsig = nsig + 1
              End If
            End If
          End If
        End If
        pduIdx = pduIdx + 1
      End While
      REM // receive new messages
      slot = 0
      While (slot < nslots)
        idx = ((slot) >> (3))
        bitVal = ((1) << ((((slot) And (7)))))
        prevBit = 0
        If (idx < (prevBitmap).Length) Then
          prevBit = ((prevBitmap(idx)) And (bitVal))
        End If
        If ((((newBitmap(idx)) And (bitVal))) <> 0) Then
          If (prevBit = 0) Then
            sms = Me.fetchPdu(slot)
            newArr.Add(sms)
            If (sms.get_concatCount() = 0) Then
              newMsg.Add(sms)
            Else
              sig = sms.get_concatSignature()
              i = 0
              While ((i < nsig) AndAlso ((sig).Length > 0))
                If (signatures(i) = sig) Then
                  sig = ""
                End If
                i = i + 1
              End While
              If ((sig).Length > 0) Then
                signatures.Add(sig)
                nsig = nsig + 1
              End If
            End If
          End If
        End If
        slot = slot + 1
      End While

      Me._pdus = newArr
      REM // append complete concatenated messages
      i = 0
      While (i < nsig)
        sig = signatures(i)
        cnt = 0
        pduIdx = 0
        While (pduIdx < Me._pdus.Count)
          sms = Me._pdus(pduIdx)
          If (sms.get_concatCount() > 0) Then
            If (sms.get_concatSignature() = sig) Then
              If (cnt = 0) Then
                cnt = sms.get_concatCount()
                newAgg.Clear()
              End If
              newAgg.Add(sms)
            End If
          End If
          pduIdx = pduIdx + 1
        End While
        If ((cnt > 0) AndAlso (newAgg.Count = cnt)) Then
          sms = New YSms(Me)
          sms.set_parts(newAgg)
          newMsg.Add(sms)
        End If
        i = i + 1
      End While

      Me._messages = newMsg
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function get_pdus() As List(Of YSms)
      Me.checkNewMessages()
      Return Me._pdus
    End Function

    '''*
    ''' <summary>
    '''   Clear the SMS units counters.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function clearPduCounters() As Integer
      Dim retcode As Integer = 0

      retcode = Me.set_pduReceived(0)
      If (retcode <> YAPI.SUCCESS) Then
        Return retcode
      End If
      retcode = Me.set_pduSent(0)
      Return retcode
    End Function

    '''*
    ''' <summary>
    '''   Sends a regular text SMS, with standard parameters.
    ''' <para>
    '''   This function can send messages
    '''   of more than 160 characters, using SMS concatenation. ISO-latin accented characters
    '''   are supported. For sending messages with special unicode characters such as asian
    '''   characters and emoticons, use <c>newMessage</c> to create a new message and define
    '''   the content of using methods <c>addText</c> and <c>addUnicodeData</c>.
    ''' </para>
    ''' </summary>
    ''' <param name="recipient">
    '''   a text string with the recipient phone number, either as a
    '''   national number, or in international format starting with a plus sign
    ''' </param>
    ''' <param name="message">
    '''   the text to be sent in the message
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function sendTextMessage(recipient As String, message As String) As Integer
      Dim sms As YSms

      sms = New YSms(Me)
      sms.set_recipient(recipient)
      sms.addText(message)
      Return sms.send()
    End Function

    '''*
    ''' <summary>
    '''   Sends a Flash SMS (class 0 message).
    ''' <para>
    '''   Flash messages are displayed on the handset
    '''   immediately and are usually not saved on the SIM card. This function can send messages
    '''   of more than 160 characters, using SMS concatenation. ISO-latin accented characters
    '''   are supported. For sending messages with special unicode characters such as asian
    '''   characters and emoticons, use <c>newMessage</c> to create a new message and define
    '''   the content of using methods <c>addText</c> et <c>addUnicodeData</c>.
    ''' </para>
    ''' </summary>
    ''' <param name="recipient">
    '''   a text string with the recipient phone number, either as a
    '''   national number, or in international format starting with a plus sign
    ''' </param>
    ''' <param name="message">
    '''   the text to be sent in the message
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function sendFlashMessage(recipient As String, message As String) As Integer
      Dim sms As YSms

      sms = New YSms(Me)
      sms.set_recipient(recipient)
      sms.set_msgClass(0)
      sms.addText(message)
      Return sms.send()
    End Function

    '''*
    ''' <summary>
    '''   Creates a new empty SMS message, to be configured and sent later on.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="recipient">
    '''   a text string with the recipient phone number, either as a
    '''   national number, or in international format starting with a plus sign
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function newMessage(recipient As String) As YSms
      Dim sms As YSms
      sms = New YSms(Me)
      sms.set_recipient(recipient)
      Return sms
    End Function

    '''*
    ''' <summary>
    '''   Returns the list of messages received and not deleted.
    ''' <para>
    '''   This function
    '''   will automatically decode concatenated SMS.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an YSms object list.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list.
    ''' </para>
    '''/
    Public Overridable Function get_messages() As List(Of YSms)
      Me.checkNewMessages()
      Return Me._messages
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of SMS message box interfaces started using <c>yFirstMessageBox()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned SMS message box interfaces order.
    '''   If you want to find a specific a SMS message box interface, use <c>MessageBox.findMessageBox()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMessageBox</c> object, corresponding to
    '''   a SMS message box interface currently online, or a <c>Nothing</c> pointer
    '''   if there are no more SMS message box interfaces to enumerate.
    ''' </returns>
    '''/
    Public Function nextMessageBox() As YMessageBox
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YMessageBox.FindMessageBox(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of SMS message box interfaces currently accessible.
    ''' <para>
    '''   Use the method <c>YMessageBox.nextMessageBox()</c> to iterate on
    '''   next SMS message box interfaces.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMessageBox</c> object, corresponding to
    '''   the first SMS message box interface currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstMessageBox() As YMessageBox
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("MessageBox", 0, p, size, neededsize, errmsg)
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
      Return YMessageBox.FindMessageBox(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YMessageBox public methods declaration)

  End Class

  REM --- (generated code: YMessageBox functions)

  '''*
  ''' <summary>
  '''   Retrieves a SMS message box interface for a given identifier.
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
  '''   This function does not require that the SMS message box interface is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YMessageBox.isOnline()</c> to test if the SMS message box interface is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a SMS message box interface by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the SMS message box interface, for instance
  '''   <c>YHUBGSM1.messageBox</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YMessageBox</c> object allowing you to drive the SMS message box interface.
  ''' </returns>
  '''/
  Public Function yFindMessageBox(ByVal func As String) As YMessageBox
    Return YMessageBox.FindMessageBox(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of SMS message box interfaces currently accessible.
  ''' <para>
  '''   Use the method <c>YMessageBox.nextMessageBox()</c> to iterate on
  '''   next SMS message box interfaces.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YMessageBox</c> object, corresponding to
  '''   the first SMS message box interface currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMessageBox() As YMessageBox
    Return YMessageBox.FirstMessageBox()
  End Function


  REM --- (end of generated code: YMessageBox functions)

End Module
