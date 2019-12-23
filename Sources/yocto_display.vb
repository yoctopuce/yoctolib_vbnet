'*********************************************************************
'*
'* $Id: yocto_display.vb 38899 2019-12-20 17:21:03Z mvuilleu $
'*
'* Implements yFindDisplay(), the high-level API for Display functions
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

Module yocto_display

  Public Const Y_NB_MAX_DISPLAY_LAYER As Integer = 10

  REM --- (generated code: YDisplayLayer globals)

 Public Enum  Y_ALIGN
  TOP_LEFT = 0
  CENTER_LEFT = 1
  BASELINE_LEFT = 2
  BOTTOM_LEFT = 3
  TOP_CENTER = 4
  CENTER = 5
  BASELINE_CENTER = 6
  BOTTOM_CENTER = 7
  TOP_DECIMAL = 8
  CENTER_DECIMAL = 9
  BASELINE_DECIMAL = 10
  BOTTOM_DECIMAL = 11
  TOP_RIGHT = 12
  CENTER_RIGHT = 13
  BASELINE_RIGHT = 14
  BOTTOM_RIGHT = 15
end enum
  REM --- (end of generated code: YDisplayLayer globals)
  REM --- (generated code: YDisplay globals)

  Public Const Y_ENABLED_FALSE As Integer = 0
  Public Const Y_ENABLED_TRUE As Integer = 1
  Public Const Y_ENABLED_INVALID As Integer = -1
  Public Const Y_STARTUPSEQ_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_BRIGHTNESS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ORIENTATION_LEFT As Integer = 0
  Public Const Y_ORIENTATION_UP As Integer = 1
  Public Const Y_ORIENTATION_RIGHT As Integer = 2
  Public Const Y_ORIENTATION_DOWN As Integer = 3
  Public Const Y_ORIENTATION_INVALID As Integer = -1
  Public Const Y_DISPLAYWIDTH_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_DISPLAYHEIGHT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_DISPLAYTYPE_MONO As Integer = 0
  Public Const Y_DISPLAYTYPE_GRAY As Integer = 1
  Public Const Y_DISPLAYTYPE_RGB As Integer = 2
  Public Const Y_DISPLAYTYPE_INVALID As Integer = -1
  Public Const Y_LAYERWIDTH_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_LAYERHEIGHT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_LAYERCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YDisplayValueCallback(ByVal func As YDisplay, ByVal value As String)
  Public Delegate Sub YDisplayTimedReportCallback(ByVal func As YDisplay, ByVal measure As YMeasure)
  REM --- (end of generated code: YDisplay globals)

  REM --- (generated code: YDisplayLayer class start)

  '''*
  ''' <summary>
  '''   Each <c>DisplayLayer</c> represents an image layer containing objects
  '''   to display (bitmaps, text, etc.).
  ''' <para>
  '''   The content is displayed only when
  '''   the layer is active on the screen (and not masked by other
  '''   overlapping layers).
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDisplayLayer
    REM --- (end of generated code: YDisplayLayer class start)

    REM --- (generated code: YDisplayLayer definitions)
    REM --- (end of generated code: YDisplayLayer definitions)

    REM --- (generated code: YDisplayLayer attributes declaration)
    REM --- (end of generated code: YDisplayLayer attributes declaration)


    REM --- (generated code: YDisplayLayer private methods declaration)

    REM --- (end of generated code: YDisplayLayer private methods declaration)

    REM --- (generated code: YDisplayLayer public methods declaration)
    '''*
    ''' <summary>
    '''   Reverts the layer to its initial state (fully transparent, default settings).
    ''' <para>
    '''   Reinitializes the drawing pointer to the upper left position,
    '''   and selects the most visible pen color. If you only want to erase the layer
    '''   content, use the method <c>clear()</c> instead.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function reset() As Integer
      Me._hidden = False
      Return Me.command_flush("X")
    End Function

    '''*
    ''' <summary>
    '''   Erases the whole content of the layer (makes it fully transparent).
    ''' <para>
    '''   This method does not change any other attribute of the layer.
    '''   To reinitialize the layer attributes to defaults settings, use the method
    '''   <c>reset()</c> instead.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function clear() As Integer
      Return Me.command_flush("x")
    End Function

    '''*
    ''' <summary>
    '''   Selects the pen color for all subsequent drawing functions,
    '''   including text drawing.
    ''' <para>
    '''   The pen color is provided as an RGB value.
    '''   For grayscale or monochrome displays, the value is
    '''   automatically converted to the proper range.
    ''' </para>
    ''' </summary>
    ''' <param name="color">
    '''   the desired pen color, as a 24-bit RGB value
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function selectColorPen(color As Integer) As Integer
      Return Me.command_push("c" + (color).ToString("x06"))
    End Function

    '''*
    ''' <summary>
    '''   Selects the pen gray level for all subsequent drawing functions,
    '''   including text drawing.
    ''' <para>
    '''   The gray level is provided as a number between
    '''   0 (black) and 255 (white, or whichever the lightest color is).
    '''   For monochrome displays (without gray levels), any value
    '''   lower than 128 is rendered as black, and any value equal
    '''   or above to 128 is non-black.
    ''' </para>
    ''' </summary>
    ''' <param name="graylevel">
    '''   the desired gray level, from 0 to 255
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function selectGrayPen(graylevel As Integer) As Integer
      Return Me.command_push("g" + Convert.ToString(graylevel))
    End Function

    '''*
    ''' <summary>
    '''   Selects an eraser instead of a pen for all subsequent drawing functions,
    '''   except for bitmap copy functions.
    ''' <para>
    '''   Any point drawn using the eraser
    '''   becomes transparent (as when the layer is empty), showing the other
    '''   layers beneath it.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function selectEraser() As Integer
      Return Me.command_push("e")
    End Function

    '''*
    ''' <summary>
    '''   Enables or disables anti-aliasing for drawing oblique lines and circles.
    ''' <para>
    '''   Anti-aliasing provides a smoother aspect when looked from far enough,
    '''   but it can add fuzziness when the display is looked from very close.
    '''   At the end of the day, it is your personal choice.
    '''   Anti-aliasing is enabled by default on grayscale and color displays,
    '''   but you can disable it if you prefer. This setting has no effect
    '''   on monochrome displays.
    ''' </para>
    ''' </summary>
    ''' <param name="mode">
    '''   <c>true</c> to enable anti-aliasing, <c>false</c> to
    '''   disable it.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function setAntialiasingMode(mode As Boolean) As Integer
      Return Me.command_push("a" + YAPI._boolToStr(mode))
    End Function

    '''*
    ''' <summary>
    '''   Draws a single pixel at the specified position.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of layer, in pixels
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of layer, in pixels
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drawPixel(x As Integer, y As Integer) As Integer
      Return Me.command_flush("P" + Convert.ToString(x) + "," + Convert.ToString(y))
    End Function

    '''*
    ''' <summary>
    '''   Draws an empty rectangle at a specified position.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="x1">
    '''   the distance from left of layer to the left border of the rectangle, in pixels
    ''' </param>
    ''' <param name="y1">
    '''   the distance from top of layer to the top border of the rectangle, in pixels
    ''' </param>
    ''' <param name="x2">
    '''   the distance from left of layer to the right border of the rectangle, in pixels
    ''' </param>
    ''' <param name="y2">
    '''   the distance from top of layer to the bottom border of the rectangle, in pixels
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drawRect(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Integer
      Return Me.command_flush("R" + Convert.ToString(x1) + "," + Convert.ToString(y1) + "," + Convert.ToString(x2) + "," + Convert.ToString(y2))
    End Function

    '''*
    ''' <summary>
    '''   Draws a filled rectangular bar at a specified position.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="x1">
    '''   the distance from left of layer to the left border of the rectangle, in pixels
    ''' </param>
    ''' <param name="y1">
    '''   the distance from top of layer to the top border of the rectangle, in pixels
    ''' </param>
    ''' <param name="x2">
    '''   the distance from left of layer to the right border of the rectangle, in pixels
    ''' </param>
    ''' <param name="y2">
    '''   the distance from top of layer to the bottom border of the rectangle, in pixels
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drawBar(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Integer
      Return Me.command_flush("B" + Convert.ToString(x1) + "," + Convert.ToString(y1) + "," + Convert.ToString(x2) + "," + Convert.ToString(y2))
    End Function

    '''*
    ''' <summary>
    '''   Draws an empty circle at a specified position.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of layer to the center of the circle, in pixels
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of layer to the center of the circle, in pixels
    ''' </param>
    ''' <param name="r">
    '''   the radius of the circle, in pixels
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drawCircle(x As Integer, y As Integer, r As Integer) As Integer
      Return Me.command_flush("C" + Convert.ToString(x) + "," + Convert.ToString(y) + "," + Convert.ToString(r))
    End Function

    '''*
    ''' <summary>
    '''   Draws a filled disc at a given position.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of layer to the center of the disc, in pixels
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of layer to the center of the disc, in pixels
    ''' </param>
    ''' <param name="r">
    '''   the radius of the disc, in pixels
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drawDisc(x As Integer, y As Integer, r As Integer) As Integer
      Return Me.command_flush("D" + Convert.ToString(x) + "," + Convert.ToString(y) + "," + Convert.ToString(r))
    End Function

    '''*
    ''' <summary>
    '''   Selects a font to use for the next text drawing functions, by providing the name of the
    '''   font file.
    ''' <para>
    '''   You can use a built-in font as well as a font file that you have previously
    '''   uploaded to the device built-in memory. If you experience problems selecting a font
    '''   file, check the device logs for any error message such as missing font file or bad font
    '''   file format.
    ''' </para>
    ''' </summary>
    ''' <param name="fontname">
    '''   the font file name, embedded fonts are 8x8.yfm, Small.yfm, Medium.yfm, Large.yfm (not available on
    '''   Yocto-MiniDisplay).
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function selectFont(fontname As String) As Integer
      Return Me.command_push("&" + fontname + "" + Chr(27))
    End Function

    '''*
    ''' <summary>
    '''   Draws a text string at the specified position.
    ''' <para>
    '''   The point of the text that is aligned
    '''   to the specified pixel position is called the anchor point, and can be chosen among
    '''   several options. Text is rendered from left to right, without implicit wrapping.
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of layer to the text anchor point, in pixels
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of layer to the text anchor point, in pixels
    ''' </param>
    ''' <param name="anchor">
    '''   the text anchor point, chosen among the <c>Y_ALIGN</c> enumeration:
    '''   <c>Y_ALIGN_TOP_LEFT</c>,    <c>Y_ALIGN_CENTER_LEFT</c>,    <c>Y_ALIGN_BASELINE_LEFT</c>,   
    '''   <c>Y_ALIGN_BOTTOM_LEFT</c>,
    '''   <c>Y_ALIGN_TOP_CENTER</c>,  <c>Y_ALIGN_CENTER</c>,         <c>Y_ALIGN_BASELINE_CENTER</c>, 
    '''   <c>Y_ALIGN_BOTTOM_CENTER</c>,
    '''   <c>Y_ALIGN_TOP_DECIMAL</c>, <c>Y_ALIGN_CENTER_DECIMAL</c>, <c>Y_ALIGN_BASELINE_DECIMAL</c>,
    '''   <c>Y_ALIGN_BOTTOM_DECIMAL</c>,
    '''   <c>Y_ALIGN_TOP_RIGHT</c>,   <c>Y_ALIGN_CENTER_RIGHT</c>,   <c>Y_ALIGN_BASELINE_RIGHT</c>,  
    '''   <c>Y_ALIGN_BOTTOM_RIGHT</c>.
    ''' </param>
    ''' <param name="text">
    '''   the text string to draw
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drawText(x As Integer, y As Integer, anchor As Y_ALIGN, text As String) As Integer
      Return Me.command_flush("T" + Convert.ToString(x) + "," + Convert.ToString(y) + "," + Convert.ToString(anchor) + "," + text + "" + Chr(27))
    End Function

    '''*
    ''' <summary>
    '''   Draws a GIF image at the specified position.
    ''' <para>
    '''   The GIF image must have been previously
    '''   uploaded to the device built-in memory. If you experience problems using an image
    '''   file, check the device logs for any error message such as missing image file or bad
    '''   image file format.
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of layer to the left of the image, in pixels
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of layer to the top of the image, in pixels
    ''' </param>
    ''' <param name="imagename">
    '''   the GIF file name
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drawImage(x As Integer, y As Integer, imagename As String) As Integer
      Return Me.command_flush("*" + Convert.ToString(x) + "," + Convert.ToString(y) + "," + imagename + "" + Chr(27))
    End Function

    '''*
    ''' <summary>
    '''   Draws a bitmap at the specified position.
    ''' <para>
    '''   The bitmap is provided as a binary object,
    '''   where each pixel maps to a bit, from left to right and from top to bottom.
    '''   The most significant bit of each byte maps to the leftmost pixel, and the least
    '''   significant bit maps to the rightmost pixel. Bits set to 1 are drawn using the
    '''   layer selected pen color. Bits set to 0 are drawn using the specified background
    '''   gray level, unless -1 is specified, in which case they are not drawn at all
    '''   (as if transparent).
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of layer to the left of the bitmap, in pixels
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of layer to the top of the bitmap, in pixels
    ''' </param>
    ''' <param name="w">
    '''   the width of the bitmap, in pixels
    ''' </param>
    ''' <param name="bitmap">
    '''   a binary object
    ''' </param>
    ''' <param name="bgcol">
    '''   the background gray level to use for zero bits (0 = black,
    '''   255 = white), or -1 to leave the pixels unchanged
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drawBitmap(x As Integer, y As Integer, w As Integer, bitmap As Byte(), bgcol As Integer) As Integer
      Dim destname As String
      destname = "layer" + Convert.ToString(Me._id) + ":" + Convert.ToString(w) + "," + Convert.ToString(bgcol) + "@" + Convert.ToString(x) + "," + Convert.ToString(y)
      Return Me._display.upload(destname,bitmap)
    End Function

    '''*
    ''' <summary>
    '''   Moves the drawing pointer of this layer to the specified position.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of layer, in pixels
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of layer, in pixels
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function moveTo(x As Integer, y As Integer) As Integer
      Return Me.command_push("@" + Convert.ToString(x) + "," + Convert.ToString(y))
    End Function

    '''*
    ''' <summary>
    '''   Draws a line from current drawing pointer position to the specified position.
    ''' <para>
    '''   The specified destination pixel is included in the line. The pointer position
    '''   is then moved to the end point of the line.
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of layer to the end point of the line, in pixels
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of layer to the end point of the line, in pixels
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function lineTo(x As Integer, y As Integer) As Integer
      Return Me.command_flush("-" + Convert.ToString(x) + "," + Convert.ToString(y))
    End Function

    '''*
    ''' <summary>
    '''   Outputs a message in the console area, and advances the console pointer accordingly.
    ''' <para>
    '''   The console pointer position is automatically moved to the beginning
    '''   of the next line when a newline character is met, or when the right margin
    '''   is hit. When the new text to display extends below the lower margin, the
    '''   console area is automatically scrolled up.
    ''' </para>
    ''' </summary>
    ''' <param name="text">
    '''   the message to display
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function consoleOut(text As String) As Integer
      Return Me.command_flush("!" + text + "" + Chr(27))
    End Function

    '''*
    ''' <summary>
    '''   Sets up display margins for the <c>consoleOut</c> function.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="x1">
    '''   the distance from left of layer to the left margin, in pixels
    ''' </param>
    ''' <param name="y1">
    '''   the distance from top of layer to the top margin, in pixels
    ''' </param>
    ''' <param name="x2">
    '''   the distance from left of layer to the right margin, in pixels
    ''' </param>
    ''' <param name="y2">
    '''   the distance from top of layer to the bottom margin, in pixels
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function setConsoleMargins(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Integer
      Return Me.command_push("m" + Convert.ToString(x1) + "," + Convert.ToString(y1) + "," + Convert.ToString(x2) + "," + Convert.ToString(y2))
    End Function

    '''*
    ''' <summary>
    '''   Sets up the background color used by the <c>clearConsole</c> function and by
    '''   the console scrolling feature.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bgcol">
    '''   the background gray level to use when scrolling (0 = black,
    '''   255 = white), or -1 for transparent
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function setConsoleBackground(bgcol As Integer) As Integer
      Return Me.command_push("b" + Convert.ToString(bgcol))
    End Function

    '''*
    ''' <summary>
    '''   Sets up the wrapping behavior used by the <c>consoleOut</c> function.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="wordwrap">
    '''   <c>true</c> to wrap only between words,
    '''   <c>false</c> to wrap on the last column anyway.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function setConsoleWordWrap(wordwrap As Boolean) As Integer
      Return Me.command_push("w" + YAPI._boolToStr(wordwrap))
    End Function

    '''*
    ''' <summary>
    '''   Blanks the console area within console margins, and resets the console pointer
    '''   to the upper left corner of the console.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function clearConsole() As Integer
      Return Me.command_flush("^")
    End Function

    '''*
    ''' <summary>
    '''   Sets the position of the layer relative to the display upper left corner.
    ''' <para>
    '''   When smooth scrolling is used, the display offset of the layer is
    '''   automatically updated during the next milliseconds to animate the move of the layer.
    ''' </para>
    ''' </summary>
    ''' <param name="x">
    '''   the distance from left of display to the upper left corner of the layer
    ''' </param>
    ''' <param name="y">
    '''   the distance from top of display to the upper left corner of the layer
    ''' </param>
    ''' <param name="scrollTime">
    '''   number of milliseconds to use for smooth scrolling, or
    '''   0 if the scrolling should be immediate.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function setLayerPosition(x As Integer, y As Integer, scrollTime As Integer) As Integer
      Return Me.command_flush("#" + Convert.ToString(x) + "," + Convert.ToString(y) + "," + Convert.ToString(scrollTime))
    End Function

    '''*
    ''' <summary>
    '''   Hides the layer.
    ''' <para>
    '''   The state of the layer is preserved but the layer is not displayed
    '''   on the screen until the next call to <c>unhide()</c>. Hiding the layer can positively
    '''   affect the drawing speed, since it postpones the rendering until all operations are
    '''   completed (double-buffering).
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function hide() As Integer
      Me.command_push("h")
      Me._hidden = True
      Return Me.flush_now()
    End Function

    '''*
    ''' <summary>
    '''   Shows the layer.
    ''' <para>
    '''   Shows the layer again after a hide command.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function unhide() As Integer
      Me._hidden = False
      Return Me.command_flush("s")
    End Function

    '''*
    ''' <summary>
    '''   Gets parent YDisplay.
    ''' <para>
    '''   Returns the parent YDisplay object of the current YDisplayLayer.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an <c>YDisplay</c> object
    ''' </returns>
    '''/
    Public Overridable Function get_display() As YDisplay
      Return Me._display
    End Function

    '''*
    ''' <summary>
    '''   Returns the display width, in pixels.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the display width, in pixels
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_DISPLAYWIDTH_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_displayWidth() As Integer
      Return Me._display.get_displayWidth()
    End Function

    '''*
    ''' <summary>
    '''   Returns the display height, in pixels.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the display height, in pixels
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_DISPLAYHEIGHT_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_displayHeight() As Integer
      Return Me._display.get_displayHeight()
    End Function

    '''*
    ''' <summary>
    '''   Returns the width of the layers to draw on, in pixels.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the width of the layers to draw on, in pixels
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_LAYERWIDTH_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_layerWidth() As Integer
      Return Me._display.get_layerWidth()
    End Function

    '''*
    ''' <summary>
    '''   Returns the height of the layers to draw on, in pixels.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the height of the layers to draw on, in pixels
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_LAYERHEIGHT_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_layerHeight() As Integer
      Return Me._display.get_layerHeight()
    End Function

    Public Overridable Function resetHiddenFlag() As Integer
      Me._hidden = False
      Return YAPI.SUCCESS
    End Function



    REM --- (end of generated code: YDisplayLayer public methods declaration)

    Private _cmdbuff As String = ""
    Private _display As YDisplay = Nothing
    Private _id As Integer = -1
    Private _hidden As Boolean = False


    Public Function flush_now() As Integer
      Dim res As Integer = YAPI.SUCCESS
      If (_cmdbuff <> "") Then
        res = _display.sendcommand(_cmdbuff)
        _cmdbuff = ""
      End If
      Return res
    End Function

    Private Function command_push(cmd As String) As Integer
      Dim res As Integer = YAPI.SUCCESS
      If (_cmdbuff.Length + cmd.Length >= 100) Then res = flush_now()
      If (_cmdbuff = "") Then _cmdbuff = _id.ToString()
      _cmdbuff = _cmdbuff + cmd
      Return YAPI.SUCCESS
    End Function

    Private Function command_flush(cmd As String) As Integer
      Dim res As Integer = command_push(cmd)
      If Not (_hidden) Then res = flush_now()
      Return res
    End Function

    Public Sub New(parent As YDisplay, id As String)
      Me._display = parent
      Me._id = CInt(id)
    End Sub

  End Class

  REM --- (generated code: YDisplayLayer functions)


  REM --- (end of generated code: YDisplayLayer functions)



  REM --- (generated code: YDisplay class start)

  '''*
  ''' <summary>
  '''   The <c>YDisplay</c> class allows to drive Yoctopuce displays.
  ''' <para>
  '''   Yoctopuce display interface has been designed to easily
  '''   show information and images. The device provides built-in
  '''   multi-layer rendering. Layers can be drawn offline, individually,
  '''   and freely moved on the display. It can also replay recorded
  '''   sequences (animations).
  ''' </para>
  ''' <para>
  '''   In order to draw on the screen, you should use the
  '''   <c>display.get_displayLayer</c> method to retrieve the layer(s) on
  '''   which you want to draw, and then use methods defined in
  '''   <c>YDisplayLayer</c> to draw on the layers.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDisplay
    Inherits YFunction
    REM --- (end of generated code: YDisplay class start)

    REM --- (generated code: YDisplay definitions)
    Public Const ENABLED_FALSE As Integer = 0
    Public Const ENABLED_TRUE As Integer = 1
    Public Const ENABLED_INVALID As Integer = -1
    Public Const STARTUPSEQ_INVALID As String = YAPI.INVALID_STRING
    Public Const BRIGHTNESS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ORIENTATION_LEFT As Integer = 0
    Public Const ORIENTATION_UP As Integer = 1
    Public Const ORIENTATION_RIGHT As Integer = 2
    Public Const ORIENTATION_DOWN As Integer = 3
    Public Const ORIENTATION_INVALID As Integer = -1
    Public Const DISPLAYWIDTH_INVALID As Integer = YAPI.INVALID_UINT
    Public Const DISPLAYHEIGHT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const DISPLAYTYPE_MONO As Integer = 0
    Public Const DISPLAYTYPE_GRAY As Integer = 1
    Public Const DISPLAYTYPE_RGB As Integer = 2
    Public Const DISPLAYTYPE_INVALID As Integer = -1
    Public Const LAYERWIDTH_INVALID As Integer = YAPI.INVALID_UINT
    Public Const LAYERHEIGHT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const LAYERCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of generated code: YDisplay definitions)

    REM --- (generated code: YDisplay attributes declaration)
    Protected _enabled As Integer
    Protected _startupSeq As String
    Protected _brightness As Integer
    Protected _orientation As Integer
    Protected _displayWidth As Integer
    Protected _displayHeight As Integer
    Protected _displayType As Integer
    Protected _layerWidth As Integer
    Protected _layerHeight As Integer
    Protected _layerCount As Integer
    Protected _command As String
    Protected _valueCallbackDisplay As YDisplayValueCallback
    REM --- (end of generated code: YDisplay attributes declaration)

    Protected _allDisplayLayers() As YDisplayLayer = Nothing
    Protected _recording As Boolean
    Protected _sequence As String

    Public Sub New(ByVal func As String)
      MyBase.new(func)
      _className = "Display"
      REM --- (generated code: YDisplay attributes initialization)
      _enabled = ENABLED_INVALID
      _startupSeq = STARTUPSEQ_INVALID
      _brightness = BRIGHTNESS_INVALID
      _orientation = ORIENTATION_INVALID
      _displayWidth = DISPLAYWIDTH_INVALID
      _displayHeight = DISPLAYHEIGHT_INVALID
      _displayType = DISPLAYTYPE_INVALID
      _layerWidth = LAYERWIDTH_INVALID
      _layerHeight = LAYERHEIGHT_INVALID
      _layerCount = LAYERCOUNT_INVALID
      _command = COMMAND_INVALID
      _valueCallbackDisplay = Nothing
      REM --- (end of generated code: YDisplay attributes initialization)
    End Sub

    REM --- (generated code: YDisplay private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("enabled") Then
        If (json_val.getInt("enabled") > 0) Then _enabled = 1 Else _enabled = 0
      End If
      If json_val.has("startupSeq") Then
        _startupSeq = json_val.getString("startupSeq")
      End If
      If json_val.has("brightness") Then
        _brightness = CInt(json_val.getLong("brightness"))
      End If
      If json_val.has("orientation") Then
        _orientation = CInt(json_val.getLong("orientation"))
      End If
      If json_val.has("displayWidth") Then
        _displayWidth = CInt(json_val.getLong("displayWidth"))
      End If
      If json_val.has("displayHeight") Then
        _displayHeight = CInt(json_val.getLong("displayHeight"))
      End If
      If json_val.has("displayType") Then
        _displayType = CInt(json_val.getLong("displayType"))
      End If
      If json_val.has("layerWidth") Then
        _layerWidth = CInt(json_val.getLong("layerWidth"))
      End If
      If json_val.has("layerHeight") Then
        _layerHeight = CInt(json_val.getLong("layerHeight"))
      End If
      If json_val.has("layerCount") Then
        _layerCount = CInt(json_val.getLong("layerCount"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YDisplay private methods declaration)

    REM --- (generated code: YDisplay public methods declaration)
    '''*
    ''' <summary>
    '''   Returns true if the screen is powered, false otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to true if the screen is powered,
    '''   false otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ENABLED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_enabled() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ENABLED_INVALID
        End If
      End If
      res = Me._enabled
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the power state of the display.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to the power state of the display
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
    Public Function set_enabled(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("enabled", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the name of the sequence to play when the displayed is powered on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of the sequence to play when the displayed is powered on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_STARTUPSEQ_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_startupSeq() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return STARTUPSEQ_INVALID
        End If
      End If
      res = Me._startupSeq
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the name of the sequence to play when the displayed is powered on.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the name of the sequence to play when the displayed is powered on
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
    Public Function set_startupSeq(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("startupSeq", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the luminosity of the  module informative LEDs (from 0 to 100).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the luminosity of the  module informative LEDs (from 0 to 100)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BRIGHTNESS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_brightness() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BRIGHTNESS_INVALID
        End If
      End If
      res = Me._brightness
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the brightness of the display.
    ''' <para>
    '''   The parameter is a value between 0 and
    '''   100. Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the brightness of the display
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
    Public Function set_brightness(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("brightness", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the currently selected display orientation.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_ORIENTATION_LEFT</c>, <c>Y_ORIENTATION_UP</c>, <c>Y_ORIENTATION_RIGHT</c> and
    '''   <c>Y_ORIENTATION_DOWN</c> corresponding to the currently selected display orientation
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ORIENTATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_orientation() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ORIENTATION_INVALID
        End If
      End If
      res = Me._orientation
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the display orientation.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_ORIENTATION_LEFT</c>, <c>Y_ORIENTATION_UP</c>, <c>Y_ORIENTATION_RIGHT</c> and
    '''   <c>Y_ORIENTATION_DOWN</c> corresponding to the display orientation
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
    Public Function set_orientation(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("orientation", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the display width, in pixels.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the display width, in pixels
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DISPLAYWIDTH_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_displayWidth() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DISPLAYWIDTH_INVALID
        End If
      End If
      res = Me._displayWidth
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the display height, in pixels.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the display height, in pixels
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DISPLAYHEIGHT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_displayHeight() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DISPLAYHEIGHT_INVALID
        End If
      End If
      res = Me._displayHeight
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the display type: monochrome, gray levels or full color.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_DISPLAYTYPE_MONO</c>, <c>Y_DISPLAYTYPE_GRAY</c> and <c>Y_DISPLAYTYPE_RGB</c>
    '''   corresponding to the display type: monochrome, gray levels or full color
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DISPLAYTYPE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_displayType() As Integer
      Dim res As Integer
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DISPLAYTYPE_INVALID
        End If
      End If
      res = Me._displayType
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the width of the layers to draw on, in pixels.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the width of the layers to draw on, in pixels
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LAYERWIDTH_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_layerWidth() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LAYERWIDTH_INVALID
        End If
      End If
      res = Me._layerWidth
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the height of the layers to draw on, in pixels.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the height of the layers to draw on, in pixels
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LAYERHEIGHT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_layerHeight() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LAYERHEIGHT_INVALID
        End If
      End If
      res = Me._layerHeight
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of available layers to draw on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of available layers to draw on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LAYERCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_layerCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LAYERCOUNT_INVALID
        End If
      End If
      res = Me._layerCount
      Return res
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
    '''   Retrieves a display for a given identifier.
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
    '''   This function does not require that the display is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YDisplay.isOnline()</c> to test if the display is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a display by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the display, for instance
    '''   <c>YD128X32.display</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YDisplay</c> object allowing you to drive the display.
    ''' </returns>
    '''/
    Public Shared Function FindDisplay(func As String) As YDisplay
      Dim obj As YDisplay
      obj = CType(YFunction._FindFromCache("Display", func), YDisplay)
      If ((obj Is Nothing)) Then
        obj = New YDisplay(func)
        YFunction._AddToCache("Display", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YDisplayValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackDisplay = callback
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
      If (Not (Me._valueCallbackDisplay Is Nothing)) Then
        Me._valueCallbackDisplay(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Clears the display screen and resets all display layers to their default state.
    ''' <para>
    '''   Using this function in a sequence will kill the sequence play-back. Don't use that
    '''   function to reset the display at sequence start-up.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function resetAll() As Integer
      Me.flushLayers()
      Me.resetHiddenLayerFlags()
      Return Me.sendCommand("Z")
    End Function

    '''*
    ''' <summary>
    '''   Smoothly changes the brightness of the screen to produce a fade-in or fade-out
    '''   effect.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="brightness">
    '''   the new screen brightness
    ''' </param>
    ''' <param name="duration">
    '''   duration of the brightness transition, in milliseconds.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function fade(brightness As Integer, duration As Integer) As Integer
      Me.flushLayers()
      Return Me.sendCommand("+" + Convert.ToString(brightness) + "," + Convert.ToString(duration))
    End Function

    '''*
    ''' <summary>
    '''   Starts to record all display commands into a sequence, for later replay.
    ''' <para>
    '''   The name used to store the sequence is specified when calling
    '''   <c>saveSequence()</c>, once the recording is complete.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function newSequence() As Integer
      Me.flushLayers()
      Me._sequence = ""
      Me._recording = True
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Stops recording display commands and saves the sequence into the specified
    '''   file on the display internal memory.
    ''' <para>
    '''   The sequence can be later replayed
    '''   using <c>playSequence()</c>.
    ''' </para>
    ''' </summary>
    ''' <param name="sequenceName">
    '''   the name of the newly created sequence
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function saveSequence(sequenceName As String) As Integer
      Me.flushLayers()
      Me._recording = False
      Me._upload(sequenceName, YAPI.DefaultEncoding.GetBytes(Me._sequence))
      REM //We need to use YPRINTF("") for Objective-C
      Me._sequence = ""
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Replays a display sequence previously recorded using
    '''   <c>newSequence()</c> and <c>saveSequence()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="sequenceName">
    '''   the name of the newly created sequence
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function playSequence(sequenceName As String) As Integer
      Me.flushLayers()
      Return Me.sendCommand("S" + sequenceName)
    End Function

    '''*
    ''' <summary>
    '''   Waits for a specified delay (in milliseconds) before playing next
    '''   commands in current sequence.
    ''' <para>
    '''   This method can be used while
    '''   recording a display sequence, to insert a timed wait in the sequence
    '''   (without any immediate effect). It can also be used dynamically while
    '''   playing a pre-recorded sequence, to suspend or resume the execution of
    '''   the sequence. To cancel a delay, call the same method with a zero delay.
    ''' </para>
    ''' </summary>
    ''' <param name="delay_ms">
    '''   the duration to wait, in milliseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function pauseSequence(delay_ms As Integer) As Integer
      Me.flushLayers()
      Return Me.sendCommand("W" + Convert.ToString(delay_ms))
    End Function

    '''*
    ''' <summary>
    '''   Stops immediately any ongoing sequence replay.
    ''' <para>
    '''   The display is left as is.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function stopSequence() As Integer
      Me.flushLayers()
      Return Me.sendCommand("S")
    End Function

    '''*
    ''' <summary>
    '''   Uploads an arbitrary file (for instance a GIF file) to the display, to the
    '''   specified full path name.
    ''' <para>
    '''   If a file already exists with the same path name,
    '''   its content is overwritten.
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
    '''   Copies the whole content of a layer to another layer.
    ''' <para>
    '''   The color and transparency
    '''   of all the pixels from the destination layer are set to match the source pixels.
    '''   This method only affects the displayed content, but does not change any
    '''   property of the layer object.
    '''   Note that layer 0 has no transparency support (it is always completely opaque).
    ''' </para>
    ''' </summary>
    ''' <param name="srcLayerId">
    '''   the identifier of the source layer (a number in range 0..layerCount-1)
    ''' </param>
    ''' <param name="dstLayerId">
    '''   the identifier of the destination layer (a number in range 0..layerCount-1)
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function copyLayerContent(srcLayerId As Integer, dstLayerId As Integer) As Integer
      Me.flushLayers()
      Return Me.sendCommand("o" + Convert.ToString(srcLayerId) + "," + Convert.ToString(dstLayerId))
    End Function

    '''*
    ''' <summary>
    '''   Swaps the whole content of two layers.
    ''' <para>
    '''   The color and transparency of all the pixels from
    '''   the two layers are swapped. This method only affects the displayed content, but does
    '''   not change any property of the layer objects. In particular, the visibility of each
    '''   layer stays unchanged. When used between one hidden layer and a visible layer,
    '''   this method makes it possible to easily implement double-buffering.
    '''   Note that layer 0 has no transparency support (it is always completely opaque).
    ''' </para>
    ''' </summary>
    ''' <param name="layerIdA">
    '''   the first layer (a number in range 0..layerCount-1)
    ''' </param>
    ''' <param name="layerIdB">
    '''   the second layer (a number in range 0..layerCount-1)
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function swapLayerContent(layerIdA As Integer, layerIdB As Integer) As Integer
      Me.flushLayers()
      Return Me.sendCommand("E" + Convert.ToString(layerIdA) + "," + Convert.ToString(layerIdB))
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of displays started using <c>yFirstDisplay()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned displays order.
    '''   If you want to find a specific a display, use <c>Display.findDisplay()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDisplay</c> object, corresponding to
    '''   a display currently online, or a <c>Nothing</c> pointer
    '''   if there are no more displays to enumerate.
    ''' </returns>
    '''/
    Public Function nextDisplay() As YDisplay
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YDisplay.FindDisplay(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of displays currently accessible.
    ''' <para>
    '''   Use the method <c>YDisplay.nextDisplay()</c> to iterate on
    '''   next displays.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDisplay</c> object, corresponding to
    '''   the first display currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstDisplay() As YDisplay
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Display", 0, p, size, neededsize, errmsg)
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
      Return YDisplay.FindDisplay(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YDisplay public methods declaration)


    '''*
    ''' <summary>
    '''   Returns a YDisplayLayer object that can be used to draw on the specified
    '''   layer.
    ''' <para>
    '''   The content is displayed only when the layer is active on the
    '''   screen (and not masked by other overlapping layers).
    ''' </para>
    ''' </summary>
    ''' <param name="layerId">
    '''   the identifier of the layer (a number in range 0..layerCount-1)
    ''' </param>
    ''' <returns>
    '''   an <c>YDisplayLayer</c> object
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Nothing</c>.
    ''' </para>
    '''/

    Public Function get_displayLayer(layerId As Integer) As YDisplayLayer
      Dim i As Integer
      Dim layercount As Integer = get_layerCount()

      If ((layerId < 0) Or (layerId >= layercount)) Then
        _throw(-1, "invalid DisplayLayer index, valid values are [0.." + (layercount - 1).ToString() + "]")
        Return Nothing
      End If

      If (_allDisplayLayers Is Nothing) Then
        ReDim _allDisplayLayers(layercount)
        For i = 0 To layercount
          _allDisplayLayers(i) = New YDisplayLayer(Me, CStr(i))

        Next i
      End If
      Return _allDisplayLayers(layerId)
    End Function


    Private Function flushLayers() As Integer
      Dim i As Integer
      If Not (_allDisplayLayers Is Nothing) Then
        For i = 0 To _allDisplayLayers.GetUpperBound(0)
          _allDisplayLayers(i).flush_now()
        Next i
      End If
      Return YAPI.SUCCESS
    End Function

    Private Sub resetHiddenLayerFlags()
      Dim i As Integer
      If Not (_allDisplayLayers Is Nothing) Then
        For i = 0 To _allDisplayLayers.GetUpperBound(0)
          _allDisplayLayers(i).resetHiddenFlag()
        Next i
      End If
    End Sub

    Public Function sendCommand(ByVal cmd As String) As Integer
      If Not (_recording) Then
        Return Me.set_command(cmd)
      End If
      _sequence = _sequence + cmd + Chr(13)
      Return YAPI.SUCCESS
    End Function


  End Class

  REM --- (generated code: YDisplay functions)

  '''*
  ''' <summary>
  '''   Retrieves a display for a given identifier.
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
  '''   This function does not require that the display is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YDisplay.isOnline()</c> to test if the display is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a display by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the display, for instance
  '''   <c>YD128X32.display</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YDisplay</c> object allowing you to drive the display.
  ''' </returns>
  '''/
  Public Function yFindDisplay(ByVal func As String) As YDisplay
    Return YDisplay.FindDisplay(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of displays currently accessible.
  ''' <para>
  '''   Use the method <c>YDisplay.nextDisplay()</c> to iterate on
  '''   next displays.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YDisplay</c> object, corresponding to
  '''   the first display currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstDisplay() As YDisplay
    Return YDisplay.FirstDisplay()
  End Function


  REM --- (end of generated code: YDisplay functions)


End Module
