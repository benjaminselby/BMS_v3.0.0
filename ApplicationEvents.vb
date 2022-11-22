Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    ' The following events are available for MyApplication:
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication
        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup

            ' Determine if this user should have access to the BMS Admin form, 
            ' or only the regular BMS Base form.

            Dim bmsForm As Form
            If ConfigHandler.currentUser.isBmsAdmin Then
                bmsForm = New BmsAdminMain()
            Else
                bmsForm = New BmsMain()
            End If

            My.Application.MainForm = bmsForm

        End Sub
    End Class
End Namespace
