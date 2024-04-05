' ********************************************************************
'
'  $Id: main.vb 60119 2024-03-22 09:43:37Z seb $
'
'  An example that shows how to use a  Yocto-Motor-DC
'
'  You can find more information on our web site:
'   Yocto-Motor-DC documentation:
'      https://www.yoctopuce.com/EN/products/yocto-motor-dc/doc.html
'   Visual Basic .Net API Reference:
'      https://www.yoctopuce.com/EN/doc/reference/yoctolib-vbnet-EN.html
'
' *********************************************************************

Module Module1

  Private Sub Usage()
    Dim execname = System.AppDomain.CurrentDomain.FriendlyName
    Console.WriteLine("Usage:")
    Console.WriteLine(execname + "  <serial_number>  power")
    Console.WriteLine(execname + "  <logical_name>  power")
    Console.WriteLine(execname + "  any power")
    Console.WriteLine("  power is a integer between -100 and 100%")
    Console.WriteLine("Example:")
    Console.WriteLine(execname + "  any 75")
    System.Threading.Thread.Sleep(2500)
    End
  End Sub

  Sub Main()
    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim errmsg As String = ""
    Dim target As String
    Dim power As Integer
    Dim motor As YMotor
    Dim current As YCurrent
    Dim voltage As YVoltage
    Dim temperature As YTemperature

    If argv.Length < 2 Then Usage()

    target = argv(1)
    power = Convert.ToInt32(argv(2))

    REM Setup the API to use local USB devices
    If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
      Console.WriteLine("RegisterHub error: " + errmsg)
      End
    End If

    If target = "any" Then
      motor = YMotor.FirstMotor()
      If motor Is Nothing Then
        Console.WriteLine("No module connected (check USB cable) ")
        End
      End If
      target = motor.get_module().get_serialNumber()
    End If

    motor = YMotor.FindMotor(target + ".motor")
    current = YCurrent.FindCurrent(target + ".current")
    voltage = YVoltage.FindVoltage(target + ".voltage")
    temperature = YTemperature.FindTemperature(target + ".temperature")

    If (motor.isOnline()) Then
      If (motor.get_motorStatus() >= Y_MOTORSTATUS_LOVOLT) Then motor.resetStatus()
      motor.drivingForceMove(power, 2000) REM ramp up to power in 2 seconds
      While (motor.isOnline())
        REM display motor status
        Console.WriteLine("Status=" + motor.get_advertisedValue() + "  " +
                 "Voltage=" + Str(voltage.get_currentValue()) + "V  " +
                 "Current=" + Str(current.get_currentValue() / 1000) + "A  " +
                 "Temp=" + Str(temperature.get_currentValue()) + "deg C")
        YAPI.Sleep(1000, errmsg) REM wait for one second
      End While

    Else
      Console.WriteLine("Module not connected (check USB cable) ")
    End If
    YAPI.FreeAPI()
  End Sub

End Module
