﻿Imports Microsoft.VisualBasic.ApplicationServices
Namespace My

  ' The following events are available for MyApplication:
  '
  ' Startup: Raised when the application starts, before the startup form is created.
  ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
  ' UnhandledException: Raised if the application encounters an unhandled exception.
  ' StartupNextInstance: Raised when launching a single-instance application and the application is already active.
  ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
  Partial Friend Class MyApplication
    Private Sub App_Startup(ByVal sender As Object, ByVal e As StartupEventArgs) Handles Me.Startup
      Dim errmsg As String = ""
      If (YAPI.RegisterHub("usb", errmsg) <> YAPI_SUCCESS) Then
        MessageBox.Show(errmsg)
        e.Cancel = True
      End If
      REM  YAPI.DisableExceptions()
    End Sub
  End Class


End Namespace

