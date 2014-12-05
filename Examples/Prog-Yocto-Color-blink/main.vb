
Imports System.IO
Imports System.Environment

Module Module1

  Sub Main()

    Dim errmsg As String = ""
    Dim led As YColorLed

    If (YAPI.RegisterHub("usb", errmsg) <> YAPI.SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      Environment.Exit(1)
    End If

    led = YColorLed.FirstColorLed()
    If (led Is Nothing) Then
      Console.WriteLine("No led connected (check USB cable)")
      Environment.Exit(1)
    End If

    led.resetBlinkSeq()                     REM cleans the sequence
    led.addRgbMoveToBlinkSeq(&HFF00, 500)   REM move to green in 500 ms
    led.addRgbMoveToBlinkSeq(&H0, 0)        REM switch to black instantaneously
    led.addRgbMoveToBlinkSeq(&H0, 250)      REM stays black for 250ms
    led.addRgbMoveToBlinkSeq(&HFF, 0)       REM switch to blue instantaneously
    led.addRgbMoveToBlinkSeq(&HFF, 100)     REM stays blue for 100ms
    led.addRgbMoveToBlinkSeq(&H0, 0)        REM switch to black instantaneously
    led.addRgbMoveToBlinkSeq(&H0, 250)      REM stays black for 250ms
    led.addRgbMoveToBlinkSeq(&HFF0000, 0)   REM switch to red instantaneously
    led.addRgbMoveToBlinkSeq(&HFF0000, 100) REM stays red for 100ms
    led.addRgbMoveToBlinkSeq(&H0, 0)        REM switch to black instantaneously
    led.addRgbMoveToBlinkSeq(&H0, 1000)     REM stays black for 1s
    led.startBlinkSeq()                     REM starts sequence

    Console.WriteLine("The Led is now blinking autonomously")
  End Sub

End Module
