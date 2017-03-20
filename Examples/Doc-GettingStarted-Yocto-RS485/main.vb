Imports System.IO
Imports System.Environment

Module Module1

  Sub Main()

    Dim argv() As String = System.Environment.GetCommandLineArgs()
    Dim serialPort As YSerialPort
    Dim errmsg = ""
    Dim cmd As String
    Dim slave, reg, val As Integer

    REM Setup the API to use local USB devices. You can
    REM use an IP address instead of 'usb' if the device
    REM is connected to a network.

    If (YAPI.RegisterHub("usb", errmsg) <> YAPI.SUCCESS) Then
      Console.WriteLine("yInitAPI failed: " + errmsg)
      End
    End If

    If (argv.Length > 1 And argv(1) <> "any") Then
      serialPort = yFindSerialPort(argv(1))
    Else
      serialPort = yFirstSerialPort()
      If serialPort Is Nothing Then
        Console.WriteLine("No module connected (check USB cable)")
        End
      End If
    End If

    Console.WriteLine("Please enter the MODBUS slave address (1...255)")
    Console.WriteLine("Slave: ")
    slave = Convert.ToInt32(Console.ReadLine())

    Console.WriteLine("Please select a Coil No (>=1), Input Bit No (>=10001+),")
    Console.WriteLine("       Input Register No (>=30001) or Register No (>=40001)")
    Console.WriteLine("No: ")
    reg = Convert.ToInt32(Console.ReadLine())
    While (serialPort.isOnline())
      If reg >= 40001 Then
        val = serialPort.modbusReadRegisters(slave, reg - 40001, 1)(0)
      ElseIf (reg >= 30001) Then
        val = serialPort.modbusReadInputRegisters(slave, reg - 30001, 1)(0)
      ElseIf (reg >= 10001) Then
        val = serialPort.modbusReadInputBits(slave, reg - 10001, 1)(0)
      Else
        val = serialPort.modbusReadBits(slave, reg - 1, 1)(0)
      End If
      Console.WriteLine("Current value: " + Convert.ToString(val))
      Console.WriteLine("Press ENTER to read again, Q to quit")
      If ((reg Mod 30000) < 10000) Then Console.WriteLine(" or enter a new value")

      cmd = Console.ReadLine()
      If cmd = "q" Or cmd = "Q" Then End

      If (cmd <> "" And (reg Mod 30000) < 10000) Then
        val = Convert.ToInt32(cmd)
        If reg >= 30001 Then
          serialPort.modbusWriteRegister(slave, reg - 30001, val)
        Else
          serialPort.modbusWriteBit(slave, reg - 1, val)
        End If
      End If
    End While
    yFreeAPI()
  End Sub

End Module
