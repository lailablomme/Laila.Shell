﻿Imports System.Reflection

Class Application
    Public Sub New()

    End Sub

    Private Sub Application_Startup(sender As Object, e As StartupEventArgs)
        AddHandler Me.DispatcherUnhandledException, AddressOf Application_DispatcherUnhandledException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf OnUnhandledException
        AddHandler TaskScheduler.UnobservedTaskException, AddressOf OnUnobservedTaskException
    End Sub

    Private Sub OnUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        logException(e.ExceptionObject)
    End Sub

    Private Sub OnUnobservedTaskException(sender As Object, e As UnobservedTaskExceptionEventArgs)
        logException(e.Exception)
        e.SetObserved()
    End Sub

    Private Sub Application_DispatcherUnhandledException(sender As Object, e As System.Windows.Threading.DispatcherUnhandledExceptionEventArgs)
        logException(e.Exception)
    End Sub

    Private Sub logException(ex As Exception)
        Dim folder As String = IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location)

        IO.File.AppendAllLines(IO.Path.Combine(folder, "error.log"), {String.Format("-----{0:d} {0:HH:mm:ss}-----", DateTime.Now)})

        While Not ex Is Nothing
            IO.File.AppendAllLines(IO.Path.Combine(folder, "error.log"), {ex.Message, ex.StackTrace?.ToString()})
            ex = ex.InnerException
        End While
    End Sub
End Class
