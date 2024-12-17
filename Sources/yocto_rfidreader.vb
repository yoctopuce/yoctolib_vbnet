' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindRfidReader(), the high-level API for RfidReader functions
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

Module yocto_rfidreader

    REM --- (generated code: YRfidTagInfo return codes)
    REM --- (end of generated code: YRfidTagInfo return codes)
    REM --- (generated code: YRfidTagInfo dlldef)
    REM --- (end of generated code: YRfidTagInfo dlldef)
   REM --- (generated code: YRfidTagInfo yapiwrapper)
   REM --- (end of generated code: YRfidTagInfo yapiwrapper)
  REM --- (generated code: YRfidTagInfo globals)

  REM --- (end of generated code: YRfidTagInfo globals)

  REM --- (generated code: YRfidTagInfo class start)

  '''*
  ''' <c>YRfidTagInfo</c> objects are used to describe RFID tag attributes,
  ''' such as the tag type and its storage size. These objects are returned by
  ''' method <c>get_tagInfo()</c> of class <c>YRfidReader</c>.
  '''/
  Public Class YRfidTagInfo
    REM --- (end of generated code: YRfidTagInfo class start)

    REM --- (generated code: YRfidTagInfo definitions)
    Public Const IEC_15693 As Integer = 1
    Public Const IEC_14443 As Integer = 2
    Public Const IEC_14443_MIFARE_ULTRALIGHT As Integer = 3
    Public Const IEC_14443_MIFARE_CLASSIC1K As Integer = 4
    Public Const IEC_14443_MIFARE_CLASSIC4K As Integer = 5
    Public Const IEC_14443_MIFARE_DESFIRE As Integer = 6
    Public Const IEC_14443_NTAG_213 As Integer = 7
    Public Const IEC_14443_NTAG_215 As Integer = 8
    Public Const IEC_14443_NTAG_216 As Integer = 9
    Public Const IEC_14443_NTAG_424_DNA As Integer = 10
    REM --- (end of generated code: YRfidTagInfo definitions)

    REM --- (generated code: YRfidTagInfo attributes declaration)
    Protected _tagId As String
    Protected _tagType As Integer
    Protected _typeStr As String
    Protected _size As Integer
    Protected _usable As Integer
    Protected _blksize As Integer
    Protected _fblk As Integer
    Protected _lblk As Integer
    REM --- (end of generated code: YRfidTagInfo attributes declaration)

    Public Sub New()
      MyBase.New()
      REM --- (generated code: YRfidTagInfo attributes initialization)
      _tagType = 0
      _size = 0
      _usable = 0
      _blksize = 0
      _fblk = 0
      _lblk = 0
      REM --- (end of generated code: YRfidTagInfo attributes initialization)
    End Sub

    REM --- (generated code: YRfidTagInfo private methods declaration)

    REM --- (end of generated code: YRfidTagInfo private methods declaration)

    REM --- (generated code: YRfidTagInfo public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the RFID tag identifier.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with the RFID tag identifier.
    ''' </returns>
    '''/
    Public Overridable Function get_tagId() As String
      Return Me._tagId
    End Function

    '''*
    ''' <summary>
    '''   Returns the type of the RFID tag, as a numeric constant.
    ''' <para>
    '''   (<c>IEC_14443_MIFARE_CLASSIC1K</c>, ...).
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the RFID tag type
    ''' </returns>
    '''/
    Public Overridable Function get_tagType() As Integer
      Return Me._tagType
    End Function

    '''*
    ''' <summary>
    '''   Returns the type of the RFID tag, as a string.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the RFID tag type
    ''' </returns>
    '''/
    Public Overridable Function get_tagTypeStr() As String
      Return Me._typeStr
    End Function

    '''*
    ''' <summary>
    '''   Returns the total memory size of the RFID tag, in bytes.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the total memory size of the RFID tag
    ''' </returns>
    '''/
    Public Overridable Function get_tagMemorySize() As Integer
      Return Me._size
    End Function

    '''*
    ''' <summary>
    '''   Returns the usable storage size of the RFID tag, in bytes.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the usable storage size of the RFID tag
    ''' </returns>
    '''/
    Public Overridable Function get_tagUsableSize() As Integer
      Return Me._usable
    End Function

    '''*
    ''' <summary>
    '''   Returns the block size of the RFID tag, in bytes.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the block size of the RFID tag
    ''' </returns>
    '''/
    Public Overridable Function get_tagBlockSize() As Integer
      Return Me._blksize
    End Function

    '''*
    ''' <summary>
    '''   Returns the index of the block available for data storage on the RFID tag.
    ''' <para>
    '''   Some tags have special block used to configure the tag behavior, these
    '''   blocks must be handled with precaution. However, the  block return by
    '''   <c>get_tagFirstBlock()</c> can be locked, use <c>get_tagLockState()</c>
    '''   to find out  which block are locked.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the index of the first usable storage block on the RFID tag
    ''' </returns>
    '''/
    Public Overridable Function get_tagFirstBlock() As Integer
      Return Me._fblk
    End Function

    '''*
    ''' <summary>
    '''   Returns the index of the last last black available for data storage on the RFID tag,
    '''   However, this block can be locked, use <c>get_tagLockState()</c> to find out
    '''   which block are locked.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the index of the last usable storage block on the RFID tag
    ''' </returns>
    '''/
    Public Overridable Function get_tagLastBlock() As Integer
      Return Me._lblk
    End Function

    Public Overridable Sub imm_init(tagId As String, tagType As Integer, size As Integer, usable As Integer, blksize As Integer, fblk As Integer, lblk As Integer)
      Dim typeStr As String
      typeStr = "unknown"
      If (tagType = IEC_15693) Then
        typeStr = "IEC 15693"
      End If
      If (tagType = IEC_14443) Then
        typeStr = "IEC 14443"
      End If
      If (tagType = IEC_14443_MIFARE_ULTRALIGHT) Then
        typeStr = "MIFARE Ultralight"
      End If
      If (tagType = IEC_14443_MIFARE_CLASSIC1K) Then
        typeStr = "MIFARE Classic 1K"
      End If
      If (tagType = IEC_14443_MIFARE_CLASSIC4K) Then
        typeStr = "MIFARE Classic 4K"
      End If
      If (tagType = IEC_14443_MIFARE_DESFIRE) Then
        typeStr = "MIFARE DESFire"
      End If
      If (tagType = IEC_14443_NTAG_213) Then
        typeStr = "NTAG 213"
      End If
      If (tagType = IEC_14443_NTAG_215) Then
        typeStr = "NTAG 215"
      End If
      If (tagType = IEC_14443_NTAG_216) Then
        typeStr = "NTAG 216"
      End If
      If (tagType = IEC_14443_NTAG_424_DNA) Then
        typeStr = "NTAG 424 DNA"
      End If
      Me._tagId = tagId
      Me._tagType = tagType
      Me._typeStr = typeStr
      Me._size = size
      Me._usable = usable
      Me._blksize = blksize
      Me._fblk = fblk
      Me._lblk = lblk
    End Sub



    REM --- (end of generated code: YRfidTagInfo public methods declaration)

  End Class

  REM --- (generated code: YRfidTagInfo functions)


  REM --- (end of generated code: YRfidTagInfo functions)


    REM --- (generated code: YRfidOptions return codes)
    REM --- (end of generated code: YRfidOptions return codes)
    REM --- (generated code: YRfidOptions dlldef)
    REM --- (end of generated code: YRfidOptions dlldef)
   REM --- (generated code: YRfidOptions yapiwrapper)
   REM --- (end of generated code: YRfidOptions yapiwrapper)
  REM --- (generated code: YRfidOptions globals)

  REM --- (end of generated code: YRfidOptions globals)

  REM --- (generated code: YRfidOptions class start)

  '''*
  ''' <summary>
  '''   The <c>YRfidOptions</c> objects are used to specify additional
  '''   optional parameters to RFID commands that interact with tags,
  '''   including security keys.
  ''' <para>
  '''   When instantiated,the parameters of
  '''   this object are pre-initialized to a value  which corresponds
  '''   to the most common usage.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YRfidOptions
    REM --- (end of generated code: YRfidOptions class start)

    REM --- (generated code: YRfidOptions definitions)
    Public Const NO_RFID_KEY As Integer = 0
    Public Const MIFARE_KEY_A As Integer = 1
    Public Const MIFARE_KEY_B As Integer = 2
    REM --- (end of generated code: YRfidOptions definitions)

    REM --- (generated code: YRfidOptions attributes declaration)

    '''*
    ''' <summary>
    '''   Type of security key to be used to access the RFID tag.
    ''' <para>
    '''   For MIFARE Classic tags, allowed values are
    '''   <c>Y_MIFARE_KEY_A</c> or <c>Y_MIFARE_KEY_B</c>.
    '''   The default value is <c>Y_NO_RFID_KEY</c>, in that case
    '''   the reader will use the most common default key for the
    '''   tag type.
    '''   When a security key is required, it must be provided
    '''   using property <c>HexKey</c>.
    ''' </para>
    ''' </summary>
    '''/
    Public KeyType As Integer

    '''*
    ''' <summary>
    '''   Security key to be used to access the RFID tag, as an
    '''   hexadecimal string.
    ''' <para>
    '''   The key will only be used if you
    '''   also specify which type of key it is, using property
    '''   <c>KeyType</c>.
    ''' </para>
    ''' </summary>
    '''/
    Public HexKey As String

    '''*
    ''' <summary>
    '''   Forces the use of single-block commands to access RFID tag memory blocks.
    ''' <para>
    '''   By default, the Yoctopuce library uses the most efficient access strategy
    '''   generally available for each tag type, but you can force the use of
    '''   single-block commands if the RFID tags you are using do not support
    '''   multi-block commands. If opération speed is not a priority, choose
    '''   single-block mode as it will work with any mode.
    ''' </para>
    ''' </summary>
    '''/
    Public ForceSingleBlockAccess As Boolean

    '''*
    ''' <summary>
    '''   Forces the use of multi-block commands to access RFID tag memory blocks.
    ''' <para>
    '''   By default, the Yoctopuce library uses the most efficient access strategy
    '''   generally available for each tag type, but you can force the use of
    '''   multi-block commands if you know for sure that the RFID tags you are using
    '''   do support multi-block commands. Be  aware that even if a tag allows multi-block
    '''   operations, the maximum number of blocks that can be written or read at the same
    '''   time can be (very) limited. If the tag does not support multi-block mode
    '''   for the wanted opération, the option will be ignored.
    ''' </para>
    ''' </summary>
    '''/
    Public ForceMultiBlockAccess As Boolean

    '''*
    ''' <summary>
    '''   Enables direct access to RFID tag control blocks.
    ''' <para>
    '''   By default, Yoctopuce library read and write functions only work
    '''   on data blocks and automatically skip special blocks, as specific functions are provided
    '''   to configure security parameters found in control blocks.
    '''   If you need to access control blocks in your own way using
    '''   read/write functions, enable this option.  Use this option wisely,
    '''   as overwriting a special block migth very well irreversibly alter your
    '''   tag behavior.
    ''' </para>
    ''' </summary>
    '''/
    Public EnableRawAccess As Boolean

    '''*
    ''' <summary>
    '''   Disables the tag memory overflow test.
    ''' <para>
    '''   By default, the Yoctopuce
    '''   library's read/write functions detect overruns and do not run
    '''   commands that are likely to fail. If you nevertheless wish to
    '''   try to access more memory than the tag announces, you can try to use
    '''   this option.
    ''' </para>
    ''' </summary>
    '''/
    Public DisableBoundaryChecks As Boolean

    '''*
    ''' <summary>
    '''   Enables simulation mode to check the affected block range as well
    '''   as access rights.
    ''' <para>
    '''   When this option is active, the operation is
    '''   not fully applied to the RFID tag but the affected block range
    '''   is determined and the optional access key is tested on these blocks.
    '''   The access key rights are not tested though. This option applies to
    '''   write / configure operations only, it is ignored for read operations.
    ''' </para>
    ''' </summary>
    '''/
    Public EnableDryRun As Boolean
    REM --- (end of generated code: YRfidOptions attributes declaration)

    Public Sub New()
      MyBase.New()
      REM --- (generated code: YRfidOptions attributes initialization)
      KeyType = 0
      REM --- (end of generated code: YRfidOptions attributes initialization)
    End Sub

    REM --- (generated code: YRfidOptions private methods declaration)

    REM --- (end of generated code: YRfidOptions private methods declaration)

    REM --- (generated code: YRfidOptions public methods declaration)
    Public Overridable Function imm_getParams() As String
      Dim opt As Integer = 0
      Dim res As String
      If (Me.ForceSingleBlockAccess) Then
        opt = 1
      Else
        opt = 0
      End If
      If (Me.ForceMultiBlockAccess) Then
        opt = ((opt) Or (2))
      End If
      If (Me.EnableRawAccess) Then
        opt = ((opt) Or (4))
      End If
      If (Me.DisableBoundaryChecks) Then
        opt = ((opt) Or (8))
      End If
      If (Me.EnableDryRun) Then
        opt = ((opt) Or (16))
      End If
      res = "&o=" + Convert.ToString(opt)
      If (Me.KeyType <> 0) Then
        res = "" + res + "&k=" + (Me.KeyType).ToString("x02") + ":" + Me.HexKey
      End If
      Return res
    End Function



    REM --- (end of generated code: YRfidOptions public methods declaration)

  End Class

  REM --- (generated code: YRfidOptions functions)


  REM --- (end of generated code: YRfidOptions functions)


    REM --- (generated code: YRfidStatus return codes)
    REM --- (end of generated code: YRfidStatus return codes)
    REM --- (generated code: YRfidStatus dlldef)
    REM --- (end of generated code: YRfidStatus dlldef)
   REM --- (generated code: YRfidStatus yapiwrapper)
   REM --- (end of generated code: YRfidStatus yapiwrapper)
  REM --- (generated code: YRfidStatus globals)

  REM --- (end of generated code: YRfidStatus globals)

  REM --- (generated code: YRfidStatus class start)

  '''*
  ''' <c>YRfidStatus</c> objects provide additional information about
  ''' operations on RFID tags, including the range of blocks affected
  ''' by read/write operations and possible errors when communicating
  ''' with RFID tags.
  ''' This makes it possible, for example, to distinguish communication
  ''' errors that can be recovered by an additional attempt, from
  ''' security or other errors on the tag.
  ''' Combined with the <c>EnableDryRun</c> option in <c>RfidOptions</c>,
  ''' this structure can be used to predict which blocks will be affected
  ''' by a write operation.
  '''/
  Public Class YRfidStatus
    REM --- (end of generated code: YRfidStatus class start)

    REM --- (generated code: YRfidStatus definitions)
    Public Const SUCCESS As Integer = 0
    Public Const COMMAND_NOT_SUPPORTED As Integer = 1
    Public Const COMMAND_NOT_RECOGNIZED As Integer = 2
    Public Const COMMAND_OPTION_NOT_RECOGNIZED As Integer = 3
    Public Const COMMAND_CANNOT_BE_PROCESSED_IN_TIME As Integer = 4
    Public Const UNDOCUMENTED_ERROR As Integer = 15
    Public Const BLOCK_NOT_AVAILABLE As Integer = 16
    Public Const BLOCK_ALREADY_LOCKED As Integer = 17
    Public Const BLOCK_LOCKED As Integer = 18
    Public Const BLOCK_NOT_SUCESSFULLY_PROGRAMMED As Integer = 19
    Public Const BLOCK_NOT_SUCESSFULLY_LOCKED As Integer = 20
    Public Const BLOCK_IS_PROTECTED As Integer = 21
    Public Const CRYPTOGRAPHIC_ERROR As Integer = 64
    Public Const READER_BUSY As Integer = 1000
    Public Const TAG_NOTFOUND As Integer = 1001
    Public Const TAG_LEFT As Integer = 1002
    Public Const TAG_JUSTLEFT As Integer = 1003
    Public Const TAG_COMMUNICATION_ERROR As Integer = 1004
    Public Const TAG_NOT_RESPONDING As Integer = 1005
    Public Const TIMEOUT_ERROR As Integer = 1006
    Public Const COLLISION_DETECTED As Integer = 1007
    Public Const INVALID_CMD_ARGUMENTS As Integer = -66
    Public Const UNKNOWN_CAPABILITIES As Integer = -67
    Public Const MEMORY_NOT_SUPPORTED As Integer = -68
    Public Const INVALID_BLOCK_INDEX As Integer = -69
    Public Const MEM_SPACE_UNVERRUN_ATTEMPT As Integer = -70
    Public Const BROWNOUT_DETECTED As Integer = -71
    Public Const BUFFER_OVERFLOW As Integer = -72
    Public Const CRC_ERROR As Integer = -73
    Public Const COMMAND_RECEIVE_TIMEOUT As Integer = -75
    Public Const DID_NOT_SLEEP As Integer = -76
    Public Const ERROR_DECIMAL_EXPECTED As Integer = -77
    Public Const HARDWARE_FAILURE As Integer = -78
    Public Const ERROR_HEX_EXPECTED As Integer = -79
    Public Const FIFO_LENGTH_ERROR As Integer = -80
    Public Const FRAMING_ERROR As Integer = -81
    Public Const NOT_IN_CNR_MODE As Integer = -82
    Public Const NUMBER_OU_OF_RANGE As Integer = -83
    Public Const NOT_SUPPORTED As Integer = -84
    Public Const NO_RF_FIELD_ACTIVE As Integer = -85
    Public Const READ_DATA_LENGTH_ERROR As Integer = -86
    Public Const WATCHDOG_RESET As Integer = -87
    Public Const UNKNOW_COMMAND As Integer = -91
    Public Const UNKNOW_ERROR As Integer = -92
    Public Const UNKNOW_PARAMETER As Integer = -93
    Public Const UART_RECEIVE_ERROR As Integer = -94
    Public Const WRONG_DATA_LENGTH As Integer = -95
    Public Const WRONG_MODE As Integer = -96
    Public Const UNKNOWN_DWARFxx_ERROR_CODE As Integer = -97
    Public Const RESPONSE_SHORT As Integer = -98
    Public Const UNEXPECTED_TAG_ID_IN_RESPONSE As Integer = -99
    Public Const UNEXPECTED_TAG_INDEX As Integer = -100
    Public Const READ_EOF As Integer = -101
    Public Const READ_OK_SOFAR As Integer = -102
    Public Const WRITE_DATA_MISSING As Integer = -103
    Public Const WRITE_TOO_MUCH_DATA As Integer = -104
    Public Const TRANSFER_CLOSED As Integer = -105
    Public Const COULD_NOT_BUILD_REQUEST As Integer = -106
    Public Const INVALID_OPTIONS As Integer = -107
    Public Const UNEXPECTED_RESPONSE As Integer = -108
    Public Const AFI_NOT_AVAILABLE As Integer = -109
    Public Const DSFID_NOT_AVAILABLE As Integer = -110
    Public Const TAG_RESPONSE_TOO_SHORT As Integer = -111
    Public Const DEC_EXPECTED As Integer = -112
    Public Const HEX_EXPECTED As Integer = -113
    Public Const NOT_SAME_SECOR As Integer = -114
    Public Const MIFARE_AUTHENTICATED As Integer = -115
    Public Const NO_DATABLOCK As Integer = -116
    Public Const KEYB_IS_READABLE As Integer = -117
    Public Const OPERATION_NOT_EXECUTED As Integer = -118
    Public Const BLOK_MODE_ERROR As Integer = -119
    Public Const BLOCK_NOT_WRITABLE As Integer = -120
    Public Const BLOCK_ACCESS_ERROR As Integer = -121
    Public Const BLOCK_NOT_AUTHENTICATED As Integer = -122
    Public Const ACCESS_KEY_BIT_NOT_WRITABLE As Integer = -123
    Public Const USE_KEYA_FOR_AUTH As Integer = -124
    Public Const USE_KEYB_FOR_AUTH As Integer = -125
    Public Const KEY_NOT_CHANGEABLE As Integer = -126
    Public Const BLOCK_TOO_HIGH As Integer = -127
    Public Const AUTH_ERR As Integer = -128
    Public Const NOKEY_SELECT As Integer = -129
    Public Const CARD_NOT_SELECTED As Integer = -130
    Public Const BLOCK_TO_READ_NONE As Integer = -131
    Public Const NO_TAG As Integer = -132
    Public Const TOO_MUCH_DATA As Integer = -133
    Public Const CON_NOT_SATISFIED As Integer = -134
    Public Const BLOCK_IS_SPECIAL As Integer = -135
    Public Const READ_BEYOND_ANNOUNCED_SIZE As Integer = -136
    Public Const BLOCK_ZERO_IS_RESERVED As Integer = -137
    Public Const VALUE_BLOCK_BAD_FORMAT As Integer = -138
    Public Const ISO15693_ONLY_FEATURE As Integer = -139
    Public Const ISO14443_ONLY_FEATURE As Integer = -140
    Public Const MIFARE_CLASSIC_ONLY_FEATURE As Integer = -141
    Public Const BLOCK_MIGHT_BE_PROTECTED As Integer = -142
    Public Const NO_SUCH_BLOCK As Integer = -143
    Public Const COUNT_TOO_BIG As Integer = -144
    Public Const UNKNOWN_MEM_SIZE As Integer = -145
    Public Const MORE_THAN_2BLOCKS_MIGHT_NOT_WORK As Integer = -146
    Public Const READWRITE_NOT_SUPPORTED As Integer = -147
    Public Const UNEXPECTED_VICC_ID_IN_RESPONSE As Integer = -148
    Public Const LOCKBLOCK_NOT_SUPPORTED As Integer = -150
    Public Const INTERNAL_ERROR_SHOULD_NEVER_HAPPEN As Integer = -151
    Public Const INVLD_BLOCK_MODE_COMBINATION As Integer = -152
    Public Const INVLD_ACCESS_MODE_COMBINATION As Integer = -153
    Public Const INVALID_SIZE As Integer = -154
    Public Const BAD_PASSWORD_FORMAT As Integer = -155
    Public Const RADIO_IS_OFF As Integer = -156
    REM --- (end of generated code: YRfidStatus definitions)

    REM --- (generated code: YRfidStatus attributes declaration)
    Protected _tagId As String
    Protected _errCode As Integer
    Protected _errBlk As Integer
    Protected _errMsg As String
    Protected _yapierr As Integer
    Protected _fab As Integer
    Protected _lab As Integer
    REM --- (end of generated code: YRfidStatus attributes declaration)

    Public Sub New()
      MyBase.New()
      REM --- (generated code: YRfidStatus attributes initialization)
      _errCode = 0
      _errBlk = 0
      _yapierr = 0
      _fab = 0
      _lab = 0
      REM --- (end of generated code: YRfidStatus attributes initialization)
    End Sub

    REM --- (generated code: YRfidStatus private methods declaration)

    REM --- (end of generated code: YRfidStatus private methods declaration)

    REM --- (generated code: YRfidStatus public methods declaration)
    '''*
    ''' <summary>
    '''   Returns RFID tag identifier related to the status.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with the RFID tag identifier.
    ''' </returns>
    '''/
    Public Overridable Function get_tagId() As String
      Return Me._tagId
    End Function

    '''*
    ''' <summary>
    '''   Returns the detailled error code, or 0 if no error happened.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a numeric error code
    ''' </returns>
    '''/
    Public Overridable Function get_errorCode() As Integer
      Return Me._errCode
    End Function

    '''*
    ''' <summary>
    '''   Returns the RFID tag memory block number where the error was encountered, or -1 if the
    '''   error is not specific to a memory block.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an RFID tag block number
    ''' </returns>
    '''/
    Public Overridable Function get_errorBlock() As Integer
      Return Me._errBlk
    End Function

    '''*
    ''' <summary>
    '''   Returns a string describing precisely the RFID commande result.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an error message string
    ''' </returns>
    '''/
    Public Overridable Function get_errorMessage() As String
      Return Me._errMsg
    End Function

    Public Overridable Function get_yapiError() As Integer
      Return Me._yapierr
    End Function

    '''*
    ''' <summary>
    '''   Returns the block number of the first RFID tag memory block affected
    '''   by the operation.
    ''' <para>
    '''   Depending on the type of operation and on the tag
    '''   memory granularity, this number may be smaller than the requested
    '''   memory block index.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an RFID tag block number
    ''' </returns>
    '''/
    Public Overridable Function get_firstAffectedBlock() As Integer
      Return Me._fab
    End Function

    '''*
    ''' <summary>
    '''   Returns the block number of the last RFID tag memory block affected
    '''   by the operation.
    ''' <para>
    '''   Depending on the type of operation and on the tag
    '''   memory granularity, this number may be bigger than the requested
    '''   memory block index.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an RFID tag block number
    ''' </returns>
    '''/
    Public Overridable Function get_lastAffectedBlock() As Integer
      Return Me._lab
    End Function

    Public Overridable Sub imm_init(tagId As String, errCode As Integer, errBlk As Integer, fab As Integer, lab As Integer)
      Dim errMsg As String
      If (errCode = 0) Then
        Me._yapierr = YAPI.SUCCESS
        errMsg = "Success (no error)"
      Else
        If (errCode < 0) Then
          If (errCode > -50) Then
            Me._yapierr = errCode
            errMsg = "YoctoLib error " + Convert.ToString(errCode)
          Else
            Me._yapierr = YAPI.RFID_HARD_ERROR
            errMsg = "Non-recoverable RFID error " + Convert.ToString(errCode)
          End If
        Else
          If (errCode > 1000) Then
            Me._yapierr = YAPI.RFID_SOFT_ERROR
            errMsg = "Recoverable RFID error " + Convert.ToString(errCode)
          Else
            Me._yapierr = YAPI.RFID_HARD_ERROR
            errMsg = "Non-recoverable RFID error " + Convert.ToString(errCode)
          End If
        End If
        If (errCode = TAG_NOTFOUND) Then
          errMsg = "Tag not found"
        End If
        If (errCode = TAG_JUSTLEFT) Then
          errMsg = "Tag left during operation"
        End If
        If (errCode = TAG_LEFT) Then
          errMsg = "Tag not here anymore"
        End If
        If (errCode = READER_BUSY) Then
          errMsg = "Reader is busy"
        End If
        If (errCode = INVALID_CMD_ARGUMENTS) Then
          errMsg = "Invalid command arguments"
        End If
        If (errCode = UNKNOWN_CAPABILITIES) Then
          errMsg = "Unknown capabilities"
        End If
        If (errCode = MEMORY_NOT_SUPPORTED) Then
          errMsg = "Memory no present"
        End If
        If (errCode = INVALID_BLOCK_INDEX) Then
          errMsg = "Invalid block index"
        End If
        If (errCode = MEM_SPACE_UNVERRUN_ATTEMPT) Then
          errMsg = "Tag memory space overrun attempt"
        End If
        If (errCode = COMMAND_NOT_SUPPORTED) Then
          errMsg = "The command is not supported"
        End If
        If (errCode = COMMAND_NOT_RECOGNIZED) Then
          errMsg = "The command is not recognized"
        End If
        If (errCode = COMMAND_OPTION_NOT_RECOGNIZED) Then
          errMsg = "The command option is not supported."
        End If
        If (errCode = COMMAND_CANNOT_BE_PROCESSED_IN_TIME) Then
          errMsg = "The command cannot be processed in time"
        End If
        If (errCode = UNDOCUMENTED_ERROR) Then
          errMsg = "Error with no information given"
        End If
        If (errCode = BLOCK_NOT_AVAILABLE) Then
          errMsg = "Block is not available"
        End If
        If (errCode = BLOCK_ALREADY_LOCKED) Then
          errMsg = "Block / byte is already locked and thus cannot be locked again."
        End If
        If (errCode = BLOCK_LOCKED) Then
          errMsg = "Block / byte is locked and its content cannot be changed"
        End If
        If (errCode = BLOCK_NOT_SUCESSFULLY_PROGRAMMED) Then
          errMsg = "Block was not successfully programmed"
        End If
        If (errCode = BLOCK_NOT_SUCESSFULLY_LOCKED) Then
          errMsg = "Block was not successfully locked"
        End If
        If (errCode = BLOCK_IS_PROTECTED) Then
          errMsg = "Block is protected"
        End If
        If (errCode = CRYPTOGRAPHIC_ERROR) Then
          errMsg = "Generic cryptographic error"
        End If
        If (errCode = BROWNOUT_DETECTED) Then
          errMsg = "BrownOut detected (BOD)"
        End If
        If (errCode = BUFFER_OVERFLOW) Then
          errMsg = "Buffer Overflow (BOF)"
        End If
        If (errCode = CRC_ERROR) Then
          errMsg = "Communication CRC Error (CCE)"
        End If
        If (errCode = COLLISION_DETECTED) Then
          errMsg = "Collision Detected (CLD/CDT)"
        End If
        If (errCode = COMMAND_RECEIVE_TIMEOUT) Then
          errMsg = "Command Receive Timeout (CRT)"
        End If
        If (errCode = DID_NOT_SLEEP) Then
          errMsg = "Did Not Sleep (DNS)"
        End If
        If (errCode = ERROR_DECIMAL_EXPECTED) Then
          errMsg = "Error Decimal Expected (EDX)"
        End If
        If (errCode = HARDWARE_FAILURE) Then
          errMsg = "Error Hardware Failure (EHF)"
        End If
        If (errCode = ERROR_HEX_EXPECTED) Then
          errMsg = "Error Hex Expected (EHX)"
        End If
        If (errCode = FIFO_LENGTH_ERROR) Then
          errMsg = "FIFO length error (FLE)"
        End If
        If (errCode = FRAMING_ERROR) Then
          errMsg = "Framing error (FER)"
        End If
        If (errCode = NOT_IN_CNR_MODE) Then
          errMsg = "Not in CNR Mode (NCM)"
        End If
        If (errCode = NUMBER_OU_OF_RANGE) Then
          errMsg = "Number Out of Range (NOR)"
        End If
        If (errCode = NOT_SUPPORTED) Then
          errMsg = "Not Supported (NOS)"
        End If
        If (errCode = NO_RF_FIELD_ACTIVE) Then
          errMsg = "No RF field active (NRF)"
        End If
        If (errCode = READ_DATA_LENGTH_ERROR) Then
          errMsg = "Read data length error (RDL)"
        End If
        If (errCode = WATCHDOG_RESET) Then
          errMsg = "Watchdog reset (SRT)"
        End If
        If (errCode = TAG_COMMUNICATION_ERROR) Then
          errMsg = "Tag Communication Error (TCE)"
        End If
        If (errCode = TAG_NOT_RESPONDING) Then
          errMsg = "Tag Not Responding (TNR)"
        End If
        If (errCode = TIMEOUT_ERROR) Then
          errMsg = "TimeOut Error (TOE)"
        End If
        If (errCode = UNKNOW_COMMAND) Then
          errMsg = "Unknown Command (UCO)"
        End If
        If (errCode = UNKNOW_ERROR) Then
          errMsg = "Unknown error (UER)"
        End If
        If (errCode = UNKNOW_PARAMETER) Then
          errMsg = "Unknown Parameter (UPA)"
        End If
        If (errCode = UART_RECEIVE_ERROR) Then
          errMsg = "UART Receive Error (URE)"
        End If
        If (errCode = WRONG_DATA_LENGTH) Then
          errMsg = "Wrong Data Length (WDL)"
        End If
        If (errCode = WRONG_MODE) Then
          errMsg = "Wrong Mode (WMO)"
        End If
        If (errCode = UNKNOWN_DWARFxx_ERROR_CODE) Then
          errMsg = "Unknown DWARF15 error code"
        End If
        If (errCode = UNEXPECTED_TAG_ID_IN_RESPONSE) Then
          errMsg = "Unexpected Tag id in response"
        End If
        If (errCode = UNEXPECTED_TAG_INDEX) Then
          errMsg = "internal error : unexpected TAG index"
        End If
        If (errCode = TRANSFER_CLOSED) Then
          errMsg = "transfer closed"
        End If
        If (errCode = WRITE_DATA_MISSING) Then
          errMsg = "Missing write data"
        End If
        If (errCode = WRITE_TOO_MUCH_DATA) Then
          errMsg = "Attempt to write too much data"
        End If
        If (errCode = COULD_NOT_BUILD_REQUEST) Then
          errMsg = "Could not not request"
        End If
        If (errCode = INVALID_OPTIONS) Then
          errMsg = "Invalid transfer options"
        End If
        If (errCode = UNEXPECTED_RESPONSE) Then
          errMsg = "Unexpected Tag response"
        End If
        If (errCode = AFI_NOT_AVAILABLE) Then
          errMsg = "AFI not available"
        End If
        If (errCode = DSFID_NOT_AVAILABLE) Then
          errMsg = "DSFID not available"
        End If
        If (errCode = TAG_RESPONSE_TOO_SHORT) Then
          errMsg = "Tag's response too short"
        End If
        If (errCode = DEC_EXPECTED) Then
          errMsg = "Error Decimal value Expected, or is missing"
        End If
        If (errCode = HEX_EXPECTED) Then
          errMsg = "Error Hexadecimal value Expected, or is missing"
        End If
        If (errCode = NOT_SAME_SECOR) Then
          errMsg = "Input and Output block are not in the same Sector"
        End If
        If (errCode = MIFARE_AUTHENTICATED) Then
          errMsg = "No chip with MIFARE Classic technology Authenticated"
        End If
        If (errCode = NO_DATABLOCK) Then
          errMsg = "No Data Block"
        End If
        If (errCode = KEYB_IS_READABLE) Then
          errMsg = "Key B is Readable"
        End If
        If (errCode = OPERATION_NOT_EXECUTED) Then
          errMsg = "Operation Not Executed, would have caused an overflow"
        End If
        If (errCode = BLOK_MODE_ERROR) Then
          errMsg = "Block has not been initialized as a 'value block'"
        End If
        If (errCode = BLOCK_NOT_WRITABLE) Then
          errMsg = "Block Not Writable"
        End If
        If (errCode = BLOCK_ACCESS_ERROR) Then
          errMsg = "Block Access Error"
        End If
        If (errCode = BLOCK_NOT_AUTHENTICATED) Then
          errMsg = "Block Not Authenticated"
        End If
        If (errCode = ACCESS_KEY_BIT_NOT_WRITABLE) Then
          errMsg = "Access bits or Keys not Writable"
        End If
        If (errCode = USE_KEYA_FOR_AUTH) Then
          errMsg = "Use Key B for authentication"
        End If
        If (errCode = USE_KEYB_FOR_AUTH) Then
          errMsg = "Use Key A for authentication"
        End If
        If (errCode = KEY_NOT_CHANGEABLE) Then
          errMsg = "Key(s) not changeable"
        End If
        If (errCode = BLOCK_TOO_HIGH) Then
          errMsg = "Block index is too high"
        End If
        If (errCode = AUTH_ERR) Then
          errMsg = "Authentication Error (i.e. wrong key)"
        End If
        If (errCode = NOKEY_SELECT) Then
          errMsg = "No Key Select, select a temporary or a static key"
        End If
        If (errCode = CARD_NOT_SELECTED) Then
          errMsg = " Card is Not Selected"
        End If
        If (errCode = BLOCK_TO_READ_NONE) Then
          errMsg = "Number of Blocks to Read is 0"
        End If
        If (errCode = NO_TAG) Then
          errMsg = "No Tag detected"
        End If
        If (errCode = TOO_MUCH_DATA) Then
          errMsg = "Too Much Data (i.e. Uart input buffer overflow)"
        End If
        If (errCode = CON_NOT_SATISFIED) Then
          errMsg = "Conditions Not Satisfied"
        End If
        If (errCode = BLOCK_IS_SPECIAL) Then
          errMsg = "Bad parameter: block is a special block"
        End If
        If (errCode = READ_BEYOND_ANNOUNCED_SIZE) Then
          errMsg = "Attempt to read more than announced size."
        End If
        If (errCode = BLOCK_ZERO_IS_RESERVED) Then
          errMsg = "Block 0 is reserved and cannot be used"
        End If
        If (errCode = VALUE_BLOCK_BAD_FORMAT) Then
          errMsg = "One value block is not properly initialized"
        End If
        If (errCode = ISO15693_ONLY_FEATURE) Then
          errMsg = "Feature available on ISO 15693 only"
        End If
        If (errCode = ISO14443_ONLY_FEATURE) Then
          errMsg = "Feature available on ISO 14443 only"
        End If
        If (errCode = MIFARE_CLASSIC_ONLY_FEATURE) Then
          errMsg = "Feature available on ISO 14443 MIFARE Classic only"
        End If
        If (errCode = BLOCK_MIGHT_BE_PROTECTED) Then
          errMsg = "Block might be protected"
        End If
        If (errCode = NO_SUCH_BLOCK) Then
          errMsg = "No such block"
        End If
        If (errCode = COUNT_TOO_BIG) Then
          errMsg = "Count parameter is too large"
        End If
        If (errCode = UNKNOWN_MEM_SIZE) Then
          errMsg = "Tag memory size is unknown"
        End If
        If (errCode = MORE_THAN_2BLOCKS_MIGHT_NOT_WORK) Then
          errMsg = "Writing more than two blocks at once might not be supported by this tag"
        End If
        If (errCode = READWRITE_NOT_SUPPORTED) Then
          errMsg = "Read/write operation not supported for this tag"
        End If
        If (errCode = UNEXPECTED_VICC_ID_IN_RESPONSE) Then
          errMsg = "Unexpected VICC ID in response"
        End If
        If (errCode = LOCKBLOCK_NOT_SUPPORTED) Then
          errMsg = "This tag does not support the Lock block function"
        End If
        If (errCode = INTERNAL_ERROR_SHOULD_NEVER_HAPPEN) Then
          errMsg = "Yoctopuce RFID code ran into an unexpected state, please contact support"
        End If
        If (errCode = INVLD_BLOCK_MODE_COMBINATION) Then
          errMsg = "Invalid combination of block mode options"
        End If
        If (errCode = INVLD_ACCESS_MODE_COMBINATION) Then
          errMsg = "Invalid combination of access mode options"
        End If
        If (errCode = INVALID_SIZE) Then
          errMsg = "Invalid data size parameter"
        End If
        If (errCode = BAD_PASSWORD_FORMAT) Then
          errMsg = "Bad password format or type"
        End If
        If (errCode = RADIO_IS_OFF) Then
          errMsg = "Radio is OFF (refreshRate=0)."
        End If
        If (errBlk >= 0) Then
          errMsg = "" + errMsg + " (block " + Convert.ToString(errBlk) + ")"
        End If
      End If
      Me._tagId = tagId
      Me._errCode = errCode
      Me._errBlk = errBlk
      Me._errMsg = errMsg
      Me._fab = fab
      Me._lab = lab
    End Sub



    REM --- (end of generated code: YRfidStatus public methods declaration)

  End Class

  REM --- (generated code: YRfidStatus functions)


  REM --- (end of generated code: YRfidStatus functions)


    REM --- (generated code: YRfidReader return codes)
    REM --- (end of generated code: YRfidReader return codes)
    REM --- (generated code: YRfidReader dlldef)
    REM --- (end of generated code: YRfidReader dlldef)
   REM --- (generated code: YRfidReader yapiwrapper)
   REM --- (end of generated code: YRfidReader yapiwrapper)
  REM --- (generated code: YRfidReader globals)

  Public Const Y_NTAGS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_REFRESHRATE_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YRfidReaderValueCallback(ByVal func As YRfidReader, ByVal value As String)
  Public Delegate Sub YRfidReaderTimedReportCallback(ByVal func As YRfidReader, ByVal measure As YMeasure)
  Public Delegate Sub YEventCallback(ByVal func As YRfidReader, ByVal timestamp As Double, ByVal eventType As String, ByVal eventData As String)

  Sub yInternalEventCallback(ByVal func As YRfidReader, ByVal value As String)
    func._internalEventHandler(value)
  End Sub
  REM --- (end of generated code: YRfidReader globals)

  REM --- (generated code: YRfidReader class start)

  '''*
  ''' <summary>
  '''   The <c>YRfidReader</c> class allows you to detect RFID tags, as well as
  '''   read and write on these tags if the security settings allow it.
  ''' <para>
  ''' </para>
  ''' <para>
  '''   Short reminder:
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   - A tag's memory is generally organized in fixed-size blocks.
  ''' </para>
  ''' <para>
  '''   - At tag level, each block must be read and written in its entirety.
  ''' </para>
  ''' <para>
  '''   - Some blocks are special configuration blocks, and may alter the tag's behavior
  '''   if they are rewritten with arbitrary data.
  ''' </para>
  ''' <para>
  '''   - Data blocks can be set to read-only mode, but on many tags, this operation is irreversible.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   By default, the RfidReader class automatically manages these blocks so that
  '''   arbitrary size data  can be manipulated of  without risk and without knowledge of
  '''   tag architecture.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YRfidReader
    Inherits YFunction
    REM --- (end of generated code: YRfidReader class start)

    REM --- (generated code: YRfidReader definitions)
    Public Const NTAGS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const REFRESHRATE_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of generated code: YRfidReader definitions)

    REM --- (generated code: YRfidReader attributes declaration)
    Protected _nTags As Integer
    Protected _refreshRate As Integer
    Protected _valueCallbackRfidReader As YRfidReaderValueCallback
    Protected _eventCallback As YEventCallback
    Protected _isFirstCb As Boolean
    Protected _prevCbPos As Integer
    Protected _eventPos As Integer
    Protected _eventStamp As Integer
    REM --- (end of generated code: YRfidReader attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "RfidReader"
      REM --- (generated code: YRfidReader attributes initialization)
      _nTags = NTAGS_INVALID
      _refreshRate = REFRESHRATE_INVALID
      _valueCallbackRfidReader = Nothing
      _prevCbPos = 0
      _eventPos = 0
      _eventStamp = 0
      REM --- (end of generated code: YRfidReader attributes initialization)
    End Sub

    REM --- (generated code: YRfidReader private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("nTags") Then
        _nTags = CInt(json_val.getLong("nTags"))
      End If
      If json_val.has("refreshRate") Then
        _refreshRate = CInt(json_val.getLong("refreshRate"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YRfidReader private methods declaration)

    REM --- (generated code: YRfidReader public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the number of RFID tags currently detected.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of RFID tags currently detected
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YRfidReader.NTAGS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nTags() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NTAGS_INVALID
        End If
      End If
      res = Me._nTags
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the tag list refresh rate, measured in Hz.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the tag list refresh rate, measured in Hz
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YRfidReader.REFRESHRATE_INVALID</c>.
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
    '''   Changes the present tag list refresh rate, measured in Hz.
    ''' <para>
    '''   The reader will do
    '''   its best to respect it. Note that the reader cannot detect tag arrival or removal
    '''   while it is  communicating with a tag.  Maximum frequency is limited to 100Hz,
    '''   but in real life it will be difficult to do better than 50Hz.  A zero value
    '''   will power off the device radio.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the present tag list refresh rate, measured in Hz
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
    '''   Retrieves a RFID reader for a given identifier.
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
    '''   This function does not require that the RFID reader is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YRfidReader.isOnline()</c> to test if the RFID reader is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a RFID reader by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the RFID reader, for instance
    '''   <c>MyDevice.rfidReader</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YRfidReader</c> object allowing you to drive the RFID reader.
    ''' </returns>
    '''/
    Public Shared Function FindRfidReader(func As String) As YRfidReader
      Dim obj As YRfidReader
      obj = CType(YFunction._FindFromCache("RfidReader", func), YRfidReader)
      If ((obj Is Nothing)) Then
        obj = New YRfidReader(func)
        YFunction._AddToCache("RfidReader", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YRfidReaderValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackRfidReader = callback
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
      If (Not (Me._valueCallbackRfidReader Is Nothing)) Then
        Me._valueCallbackRfidReader(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function _chkerror(tagId As String, json As Byte(), ByRef status As YRfidStatus) As Integer
      Dim jsonStr As String
      Dim errCode As Integer = 0
      Dim errBlk As Integer = 0
      Dim fab As Integer = 0
      Dim lab As Integer = 0
      Dim retcode As Integer = 0

      If ((json).Length = 0) Then
        errCode = Me.get_errorType()
        errBlk = -1
        fab = -1
        lab = -1
      Else
        jsonStr = YAPI.DefaultEncoding.GetString(json)
        errCode = YAPI._atoi(Me._json_get_key(json, "err"))
        errBlk = YAPI._atoi(Me._json_get_key(json, "errBlk"))-1
        If (jsonStr.IndexOf("""fab"":") >= 0) Then
          fab = YAPI._atoi(Me._json_get_key(json, "fab"))-1
          lab = YAPI._atoi(Me._json_get_key(json, "lab"))-1
        Else
          fab = -1
          lab = -1
        End If
      End If
      status.imm_init(tagId, errCode, errBlk, fab, lab)
      retcode = status.get_yapiError()
      If Not(retcode = YAPI.SUCCESS) Then
        me._throw(retcode, status.get_errorMessage())
        return retcode
      end if
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function reset() As Integer
      Dim json As Byte() = New Byte(){}
      Dim status As YRfidStatus
      status = New YRfidStatus()

      json = Me._download("rfid.json?a=reset")
      Return Me._chkerror("", json, status)
    End Function

    '''*
    ''' <summary>
    '''   Returns the list of RFID tags currently detected by the reader.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of strings, corresponding to each tag identifier (UID).
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list.
    ''' </para>
    '''/
    Public Overridable Function get_tagIdList() As List(Of String)
      Dim json As Byte() = New Byte(){}
      Dim jsonList As List(Of Byte()) = New List(Of Byte())()
      Dim taglist As List(Of String) = New List(Of String)()

      json = Me._download("rfid.json?a=list")
      taglist.Clear()
      If ((json).Length > 3) Then
        jsonList = Me._json_get_array(json)
        Dim ii_0 As Integer
        For ii_0 = 0 To jsonList.Count - 1
          taglist.Add(Me._json_get_string(jsonList(ii_0)))
        Next ii_0
      End If
      Return taglist
    End Function

    '''*
    ''' <summary>
    '''   Returns a description of the properties of an existing RFID tag.
    ''' <para>
    '''   This function can cause communications with the tag.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to check
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   a <c>YRfidTagInfo</c> object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty <c>YRfidTagInfo</c> objact.
    '''   When it happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function get_tagInfo(tagId As String, ByRef status As YRfidStatus) As YRfidTagInfo
      Dim url As String
      Dim json As Byte() = New Byte(){}
      Dim tagType As Integer = 0
      Dim size As Integer = 0
      Dim usable As Integer = 0
      Dim blksize As Integer = 0
      Dim fblk As Integer = 0
      Dim lblk As Integer = 0
      Dim res As YRfidTagInfo
      url = "rfid.json?a=info&t=" + tagId

      json = Me._download(url)
      Me._chkerror(tagId, json, status)
      tagType = YAPI._atoi(Me._json_get_key(json, "type"))
      size = YAPI._atoi(Me._json_get_key(json, "size"))
      usable = YAPI._atoi(Me._json_get_key(json, "usable"))
      blksize = YAPI._atoi(Me._json_get_key(json, "blksize"))
      fblk = YAPI._atoi(Me._json_get_key(json, "fblk"))
      lblk = YAPI._atoi(Me._json_get_key(json, "lblk"))
      res = New YRfidTagInfo()
      res.imm_init(tagId, tagType, size, usable, blksize, fblk, lblk)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Changes an RFID tag configuration to prevents any further write to
    '''   the selected blocks.
    ''' <para>
    '''   This operation is definitive and irreversible.
    '''   Depending on the tag type and block index, adjascent blocks may become
    '''   read-only as well, based on the locking granularity.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   first block to lock
    ''' </param>
    ''' <param name="nBlocks">
    '''   number of blocks to lock
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagLockBlocks(tagId As String, firstBlock As Integer, nBlocks As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      optstr = options.imm_getParams()
      url = "rfid.json?a=lock&t=" + tagId + "&b=" + Convert.ToString(firstBlock) + "&n=" + Convert.ToString(nBlocks) + "" + optstr

      json = Me._download(url)
      Return Me._chkerror(tagId, json, status)
    End Function

    '''*
    ''' <summary>
    '''   Reads the locked state for RFID tag memory data blocks.
    ''' <para>
    '''   FirstBlock cannot be a special block, and any special
    '''   block encountered in the middle of the read operation will be
    '''   skipped automatically.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   number of the first block to check
    ''' </param>
    ''' <param name="nBlocks">
    '''   number of blocks to check
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   a list of booleans with the lock state of selected blocks
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function get_tagLockState(tagId As String, firstBlock As Integer, nBlocks As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As List(Of Boolean)
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      Dim binRes As Byte() = New Byte(){}
      Dim res As List(Of Boolean) = New List(Of Boolean)()
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim isLocked As Boolean
      optstr = options.imm_getParams()
      url = "rfid.json?a=chkl&t=" + tagId + "&b=" + Convert.ToString(firstBlock) + "&n=" + Convert.ToString(nBlocks) + "" + optstr

      json = Me._download(url)
      Me._chkerror(tagId, json, status)
      If (status.get_yapiError() <> YAPI.SUCCESS) Then
        Return res
      End If

      binRes = YAPI._hexStrToBin(Me._json_get_key(json, "bitmap"))
      idx = 0
      While (idx < nBlocks)
        val = binRes((idx >> 3))
        isLocked = (((val) And ((1 << ((idx) And (7))))) <> 0)
        res.Add(isLocked)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Tells which block of a RFID tag memory are special and cannot be used
    '''   to store user data.
    ''' <para>
    '''   Mistakely writing a special block can lead to
    '''   an irreversible alteration of the tag.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   number of the first block to check
    ''' </param>
    ''' <param name="nBlocks">
    '''   number of blocks to check
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   a list of booleans with the lock state of selected blocks
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function get_tagSpecialBlocks(tagId As String, firstBlock As Integer, nBlocks As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As List(Of Boolean)
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      Dim binRes As Byte() = New Byte(){}
      Dim res As List(Of Boolean) = New List(Of Boolean)()
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim isLocked As Boolean
      optstr = options.imm_getParams()
      url = "rfid.json?a=chks&t=" + tagId + "&b=" + Convert.ToString(firstBlock) + "&n=" + Convert.ToString(nBlocks) + "" + optstr

      json = Me._download(url)
      Me._chkerror(tagId, json, status)
      If (status.get_yapiError() <> YAPI.SUCCESS) Then
        Return res
      End If

      binRes = YAPI._hexStrToBin(Me._json_get_key(json, "bitmap"))
      idx = 0
      While (idx < nBlocks)
        val = binRes((idx >> 3))
        isLocked = (((val) And ((1 << ((idx) And (7))))) <> 0)
        res.Add(isLocked)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads data from an RFID tag memory, as an hexadecimal string.
    ''' <para>
    '''   The read operation may span accross multiple blocks if the requested
    '''   number of bytes is larger than the RFID tag block size. By default
    '''   firstBlock cannot be a special block, and any special block encountered
    '''   in the middle of the read operation will be skipped automatically.
    '''   If you rather want to read special blocks, use the <c>EnableRawAccess</c>
    '''   field from the <c>options</c> parameter.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   block number where read should start
    ''' </param>
    ''' <param name="nBytes">
    '''   total number of bytes to read
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   an hexadecimal string if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty binary buffer. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagReadHex(tagId As String, firstBlock As Integer, nBytes As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As String
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      Dim hexbuf As String
      optstr = options.imm_getParams()
      url = "rfid.json?a=read&t=" + tagId + "&b=" + Convert.ToString(firstBlock) + "&n=" + Convert.ToString(nBytes) + "" + optstr

      json = Me._download(url)
      Me._chkerror(tagId, json, status)
      If (status.get_yapiError() = YAPI.SUCCESS) Then
        hexbuf = Me._json_get_key(json, "res")
      Else
        hexbuf = ""
      End If
      Return hexbuf
    End Function

    '''*
    ''' <summary>
    '''   Reads data from an RFID tag memory, as a binary buffer.
    ''' <para>
    '''   The read operation
    '''   may span accross multiple blocks if the requested number of bytes
    '''   is larger than the RFID tag block size.  By default
    '''   firstBlock cannot be a special block, and any special block encountered
    '''   in the middle of the read operation will be skipped automatically.
    '''   If you rather want to read special blocks, use the <c>EnableRawAccess</c>
    '''   field frrm the <c>options</c> parameter.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   block number where read should start
    ''' </param>
    ''' <param name="nBytes">
    '''   total number of bytes to read
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   a binary object with the data read if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty binary buffer. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagReadBin(tagId As String, firstBlock As Integer, nBytes As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As Byte()
      Return YAPI._hexStrToBin(Me.tagReadHex(tagId, firstBlock, nBytes, options, status))
    End Function

    '''*
    ''' <summary>
    '''   Reads data from an RFID tag memory, as a byte list.
    ''' <para>
    '''   The read operation
    '''   may span accross multiple blocks if the requested number of bytes
    '''   is larger than the RFID tag block size.  By default
    '''   firstBlock cannot be a special block, and any special block encountered
    '''   in the middle of the read operation will be skipped automatically.
    '''   If you rather want to read special blocks, use the <c>EnableRawAccess</c>
    '''   field from the <c>options</c> parameter.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   block number where read should start
    ''' </param>
    ''' <param name="nBytes">
    '''   total number of bytes to read
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   a byte list with the data read if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagReadArray(tagId As String, firstBlock As Integer, nBytes As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As List(Of Integer)
      Dim blk As Byte() = New Byte(){}
      Dim idx As Integer = 0
      Dim endidx As Integer = 0
      Dim res As List(Of Integer) = New List(Of Integer)()
      blk = Me.tagReadBin(tagId, firstBlock, nBytes, options, status)
      endidx = (blk).Length

      idx = 0
      While (idx < endidx)
        res.Add(blk(idx))
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads data from an RFID tag memory, as a text string.
    ''' <para>
    '''   The read operation
    '''   may span accross multiple blocks if the requested number of bytes
    '''   is larger than the RFID tag block size.  By default
    '''   firstBlock cannot be a special block, and any special block encountered
    '''   in the middle of the read operation will be skipped automatically.
    '''   If you rather want to read special blocks, use the <c>EnableRawAccess</c>
    '''   field from the <c>options</c> parameter.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   block number where read should start
    ''' </param>
    ''' <param name="nChars">
    '''   total number of characters to read
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   a text string with the data read if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagReadStr(tagId As String, firstBlock As Integer, nChars As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As String
      Return YAPI.DefaultEncoding.GetString(Me.tagReadBin(tagId, firstBlock, nChars, options, status))
    End Function

    '''*
    ''' <summary>
    '''   Writes data provided as a binary buffer to an RFID tag memory.
    ''' <para>
    '''   The write operation may span accross multiple blocks if the
    '''   number of bytes to write is larger than the RFID tag block size.
    '''   By default firstBlock cannot be a special block, and any special block
    '''   encountered in the middle of the write operation will be skipped
    '''   automatically. The last data block affected by the operation will
    '''   be automatically padded with zeros if neccessary.  If you rather want
    '''   to rewrite special blocks as well,
    '''   use the <c>EnableRawAccess</c> field from the <c>options</c> parameter.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   block number where write should start
    ''' </param>
    ''' <param name="buff">
    '''   the binary buffer to write
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagWriteBin(tagId As String, firstBlock As Integer, buff As Byte(), options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim optstr As String
      Dim hexstr As String
      Dim buflen As Integer = 0
      Dim fname As String
      Dim json As Byte() = New Byte(){}
      buflen = (buff).Length
      If (buflen <= 16) Then
        REM // short data, use an URL-based command
        hexstr = YAPI._bytesToHexStr(buff, 0, buff.Length)
        Return Me.tagWriteHex(tagId, firstBlock, hexstr, options, status)
      Else
        REM // long data, use an upload command
        optstr = options.imm_getParams()
        fname = "Rfid:t=" + tagId + "&b=" + Convert.ToString(firstBlock) + "&n=" + Convert.ToString(buflen) + "" + optstr
        json = Me._uploadEx(fname, buff)
        Return Me._chkerror(tagId, json, status)
      End If
    End Function

    '''*
    ''' <summary>
    '''   Writes data provided as a list of bytes to an RFID tag memory.
    ''' <para>
    '''   The write operation may span accross multiple blocks if the
    '''   number of bytes to write is larger than the RFID tag block size.
    '''   By default firstBlock cannot be a special block, and any special block
    '''   encountered in the middle of the write operation will be skipped
    '''   automatically. The last data block affected by the operation will
    '''   be automatically padded with zeros if neccessary.
    '''   If you rather want to rewrite special blocks as well,
    '''   use the <c>EnableRawAccess</c> field from the <c>options</c> parameter.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   block number where write should start
    ''' </param>
    ''' <param name="byteList">
    '''   a list of byte to write
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagWriteArray(tagId As String, firstBlock As Integer, byteList As List(Of Integer), options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim bufflen As Integer = 0
      Dim buff As Byte() = New Byte(){}
      Dim idx As Integer = 0
      Dim hexb As Integer = 0
      bufflen = byteList.Count
      ReDim buff(bufflen-1)
      idx = 0
      While (idx < bufflen)
        hexb = byteList(idx)
        buff(idx) = Convert.ToByte(hexb And &HFF)
        idx = idx + 1
      End While

      Return Me.tagWriteBin(tagId, firstBlock, buff, options, status)
    End Function

    '''*
    ''' <summary>
    '''   Writes data provided as an hexadecimal string to an RFID tag memory.
    ''' <para>
    '''   The write operation may span accross multiple blocks if the
    '''   number of bytes to write is larger than the RFID tag block size.
    '''   By default firstBlock cannot be a special block, and any special block
    '''   encountered in the middle of the write operation will be skipped
    '''   automatically. The last data block affected by the operation will
    '''   be automatically padded with zeros if neccessary.
    '''   If you rather want to rewrite special blocks as well,
    '''   use the <c>EnableRawAccess</c> field from the <c>options</c> parameter.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   block number where write should start
    ''' </param>
    ''' <param name="hexString">
    '''   a string of hexadecimal byte codes to write
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagWriteHex(tagId As String, firstBlock As Integer, hexString As String, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim bufflen As Integer = 0
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      Dim buff As Byte() = New Byte(){}
      Dim idx As Integer = 0
      Dim hexb As Integer = 0
      bufflen = (hexString).Length
      bufflen = (bufflen >> 1)
      If (bufflen <= 16) Then
        REM // short data, use an URL-based command
        optstr = options.imm_getParams()
        url = "rfid.json?a=writ&t=" + tagId + "&b=" + Convert.ToString(firstBlock) + "&w=" + hexString + "" + optstr
        json = Me._download(url)
        Return Me._chkerror(tagId, json, status)
      Else
        REM // long data, use an upload command
        ReDim buff(bufflen-1)
        idx = 0
        While (idx < bufflen)
          hexb = Convert.ToInt32((hexString).Substring(2 * idx, 2), 16)
          buff(idx) = Convert.ToByte(hexb And &HFF)
          idx = idx + 1
        End While
        Return Me.tagWriteBin(tagId, firstBlock, buff, options, status)
      End If
    End Function

    '''*
    ''' <summary>
    '''   Writes data provided as an ASCII string to an RFID tag memory.
    ''' <para>
    '''   The write operation may span accross multiple blocks if the
    '''   number of bytes to write is larger than the RFID tag block size.
    '''   Note that only the characters présent  in  the provided string
    '''   will be written, there is no notion of string length. If your
    '''   string data have variable length, you'll have to encode the
    '''   string length yourself, with a terminal zero for instannce.
    ''' </para>
    ''' <para>
    '''   This function only works with ISO-latin characters, if you wish to
    '''   write strings encoded with alternate character sets, you'll have to
    '''   use tagWriteBin() function.
    ''' </para>
    ''' <para>
    '''   By default firstBlock cannot be a special block, and any special block
    '''   encountered in the middle of the write operation will be skipped
    '''   automatically. The last data block affected by the operation will
    '''   be automatically padded with zeros if neccessary.
    '''   If you rather want to rewrite special blocks as well,
    '''   use the <c>EnableRawAccess</c> field from the <c>options</c> parameter
    '''   (definitely not recommanded).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="firstBlock">
    '''   block number where write should start
    ''' </param>
    ''' <param name="text">
    '''   the text string to write
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagWriteStr(tagId As String, firstBlock As Integer, text As String, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim buff As Byte() = New Byte(){}
      buff = YAPI.DefaultEncoding.GetBytes(text)

      Return Me.tagWriteBin(tagId, firstBlock, buff, options, status)
    End Function

    '''*
    ''' <summary>
    '''   Reads an RFID tag AFI byte (ISO 15693 only).
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   the AFI value (0...255)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagGetAFI(tagId As String, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      Dim res As Integer = 0
      optstr = options.imm_getParams()
      url = "rfid.json?a=rdsf&t=" + tagId + "&b=0" + optstr

      json = Me._download(url)
      Me._chkerror(tagId, json, status)
      If (status.get_yapiError() = YAPI.SUCCESS) Then
        res = YAPI._atoi(Me._json_get_key(json, "res"))
      Else
        res = status.get_yapiError()
      End If
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Changes an RFID tag AFI byte (ISO 15693 only).
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="afi">
    '''   the AFI value to write (0...255)
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagSetAFI(tagId As String, afi As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      optstr = options.imm_getParams()
      url = "rfid.json?a=wrsf&t=" + tagId + "&b=0&v=" + Convert.ToString(afi) + "" + optstr

      json = Me._download(url)
      Return Me._chkerror(tagId, json, status)
    End Function

    '''*
    ''' <summary>
    '''   Locks the RFID tag AFI byte (ISO 15693 only).
    ''' <para>
    '''   This operation is definitive and irreversible.
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagLockAFI(tagId As String, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      optstr = options.imm_getParams()
      url = "rfid.json?a=lksf&t=" + tagId + "&b=0" + optstr

      json = Me._download(url)
      Return Me._chkerror(tagId, json, status)
    End Function

    '''*
    ''' <summary>
    '''   Reads an RFID tag DSFID byte (ISO 15693 only).
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   the DSFID value (0...255)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagGetDSFID(tagId As String, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      Dim res As Integer = 0
      optstr = options.imm_getParams()
      url = "rfid.json?a=rdsf&t=" + tagId + "&b=1" + optstr

      json = Me._download(url)
      Me._chkerror(tagId, json, status)
      If (status.get_yapiError() = YAPI.SUCCESS) Then
        res = YAPI._atoi(Me._json_get_key(json, "res"))
      Else
        res = status.get_yapiError()
      End If
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Changes an RFID tag DSFID byte (ISO 15693 only).
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="dsfid">
    '''   the DSFID value to write (0...255)
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagSetDSFID(tagId As String, dsfid As Integer, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      optstr = options.imm_getParams()
      url = "rfid.json?a=wrsf&t=" + tagId + "&b=1&v=" + Convert.ToString(dsfid) + "" + optstr

      json = Me._download(url)
      Return Me._chkerror(tagId, json, status)
    End Function

    '''*
    ''' <summary>
    '''   Locks the RFID tag DSFID byte (ISO 15693 only).
    ''' <para>
    '''   This operation is definitive and irreversible.
    ''' </para>
    ''' </summary>
    ''' <param name="tagId">
    '''   identifier of the tag to use
    ''' </param>
    ''' <param name="options">
    '''   an <c>YRfidOptions</c> object with the optional
    '''   command execution parameters, such as security key
    '''   if required
    ''' </param>
    ''' <param name="status">
    '''   an <c>RfidStatus</c> object that will contain
    '''   the detailled status of the operation
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code. When it
    '''   happens, you can get more information from the <c>status</c> object.
    ''' </para>
    '''/
    Public Overridable Function tagLockDSFID(tagId As String, options As YRfidOptions, ByRef status As YRfidStatus) As Integer
      Dim optstr As String
      Dim url As String
      Dim json As Byte() = New Byte(){}
      optstr = options.imm_getParams()
      url = "rfid.json?a=lksf&t=" + tagId + "&b=1" + optstr

      json = Me._download(url)
      Return Me._chkerror(tagId, json, status)
    End Function

    '''*
    ''' <summary>
    '''   Returns a string with last tag arrival/removal events observed.
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

      content = Me._download("events.txt?pos=0")
      Return YAPI.DefaultEncoding.GetString(content)
    End Function

    '''*
    ''' <summary>
    '''   Registers a callback function to be called each time that an RFID tag appears or
    '''   disappears.
    ''' <para>
    '''   The callback is invoked only during the execution of
    '''   <c>ySleep</c> or <c>yHandleEvents</c>. This provides control over the time when
    '''   the callback is triggered. For good responsiveness, remember to call one of these
    '''   two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer.
    '''   The callback function should take four arguments:
    '''   the <c>YRfidReader</c> object that emitted the event, the
    '''   UTC timestamp of the event, a character string describing
    '''   the type of event ("+" or "-") and a character string with the
    '''   RFID tag identifier.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </param>
    '''/
    Public Overridable Function registerEventCallback(callback As YEventCallback) As Integer
      Me._eventCallback = callback
      Me._isFirstCb = True
      If (Not (callback Is Nothing)) Then
        Me.registerValueCallback(AddressOf yInternalEventCallback)
      Else
        Me.registerValueCallback(CType(Nothing, YRfidReaderValueCallback))
      End If
      Return 0
    End Function

    Public Overridable Function _internalEventHandler(cbVal As String) As Integer
      Dim cbPos As Integer = 0
      Dim cbDPos As Integer = 0
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
      Dim intStamp As Integer = 0
      Dim binMStamp As Byte() = New Byte(){}
      Dim msStamp As Integer = 0
      Dim evtStamp As Double = 0
      Dim evtType As String
      Dim evtData As String
      REM // detect possible power cycle of the reader to clear event pointer
      cbPos = YAPI._atoi(cbVal)
      cbPos = (cbPos \ 1000)
      cbDPos = ((cbPos - Me._prevCbPos) And (&H7ffff))
      Me._prevCbPos = cbPos
      If (cbDPos > 16384) Then
        Me._eventPos = 0
      End If
      If (Not (Not (Me._eventCallback Is Nothing))) Then
        Return YAPI.SUCCESS
      End If
      If (Me._isFirstCb) Then
        REM // first emulated value callback caused by registerValueCallback:
        REM // retrieve arrivals of all tags currently present to emulate arrival
        Me._isFirstCb = False
        Me._eventStamp = 0
        content = Me._download("events.txt")
        contentStr = YAPI.DefaultEncoding.GetString(content)
        eventArr = New List(Of String)(contentStr.Split(vbLf.ToCharArray()))
        arrLen = eventArr.Count
        If Not(arrLen > 0) Then
          me._throw(YAPI.IO_ERROR, "fail to download events")
          return YAPI.IO_ERROR
        end if
        REM // first element of array is the new position preceeded by '@'
        arrPos = 1
        lenStr = eventArr(0)
        lenStr = (lenStr).Substring(1, (lenStr).Length-1)
        REM // update processed event position pointer
        Me._eventPos = YAPI._atoi(lenStr)
      Else
        REM // load all events since previous call
        url = "events.txt?pos=" + Convert.ToString(Me._eventPos)
        content = Me._download(url)
        contentStr = YAPI.DefaultEncoding.GetString(content)
        eventArr = New List(Of String)(contentStr.Split(vbLf.ToCharArray()))
        arrLen = eventArr.Count
        If Not(arrLen > 0) Then
          me._throw(YAPI.IO_ERROR, "fail to download events")
          return YAPI.IO_ERROR
        end if
        REM // last element of array is the new position preceeded by '@'
        arrPos = 0
        arrLen = arrLen - 1
        lenStr = eventArr(arrLen)
        lenStr = (lenStr).Substring(1, (lenStr).Length-1)
        REM // update processed event position pointer
        Me._eventPos = YAPI._atoi(lenStr)
      End If
      REM // now generate callbacks for each real event
      While (arrPos < arrLen)
        eventStr = eventArr(arrPos)
        eventLen = (eventStr).Length
        typePos = eventStr.IndexOf(":")+1
        If ((eventLen >= 14) AndAlso (typePos > 10)) Then
          hexStamp = (eventStr).Substring(0, 8)
          intStamp = Convert.ToInt32(hexStamp, 16)
          If (intStamp >= Me._eventStamp) Then
            Me._eventStamp = intStamp
            binMStamp = YAPI.DefaultEncoding.GetBytes((eventStr).Substring(8, 2))
            msStamp = (binMStamp(0)-64) * 32 + binMStamp(1)
            evtStamp = intStamp + (0.001 * msStamp)
            dataPos = eventStr.IndexOf("=")+1
            evtType = (eventStr).Substring(typePos, 1)
            evtData = ""
            If (dataPos > 10) Then
              evtData = (eventStr).Substring(dataPos, eventLen-dataPos)
            End If
            If (Not (Me._eventCallback Is Nothing)) Then
              Me._eventCallback(Me, evtStamp, evtType, evtData)
            End If
          End If
        End If
        arrPos = arrPos + 1
      End While
      Return YAPI.SUCCESS
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of RFID readers started using <c>yFirstRfidReader()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned RFID readers order.
    '''   If you want to find a specific a RFID reader, use <c>RfidReader.findRfidReader()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRfidReader</c> object, corresponding to
    '''   a RFID reader currently online, or a <c>Nothing</c> pointer
    '''   if there are no more RFID readers to enumerate.
    ''' </returns>
    '''/
    Public Function nextRfidReader() As YRfidReader
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YRfidReader.FindRfidReader(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of RFID readers currently accessible.
    ''' <para>
    '''   Use the method <c>YRfidReader.nextRfidReader()</c> to iterate on
    '''   next RFID readers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRfidReader</c> object, corresponding to
    '''   the first RFID reader currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstRfidReader() As YRfidReader
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("RfidReader", 0, p, size, neededsize, errmsg)
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
      Return YRfidReader.FindRfidReader(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YRfidReader public methods declaration)

  End Class

  REM --- (generated code: YRfidReader functions)

  '''*
  ''' <summary>
  '''   Retrieves a RFID reader for a given identifier.
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
  '''   This function does not require that the RFID reader is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YRfidReader.isOnline()</c> to test if the RFID reader is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a RFID reader by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the RFID reader, for instance
  '''   <c>MyDevice.rfidReader</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YRfidReader</c> object allowing you to drive the RFID reader.
  ''' </returns>
  '''/
  Public Function yFindRfidReader(ByVal func As String) As YRfidReader
    Return YRfidReader.FindRfidReader(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of RFID readers currently accessible.
  ''' <para>
  '''   Use the method <c>YRfidReader.nextRfidReader()</c> to iterate on
  '''   next RFID readers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YRfidReader</c> object, corresponding to
  '''   the first RFID reader currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstRfidReader() As YRfidReader
    Return YRfidReader.FirstRfidReader()
  End Function


  REM --- (end of generated code: YRfidReader functions)

End Module
