Imports System.Reflection

Class Application
    Public Sub New()

    End Sub

    Private Sub Application_DispatcherUnhandledException(sender As Object, e As System.Windows.Threading.DispatcherUnhandledExceptionEventArgs)
        Dim folder As String = IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location)

        IO.File.AppendAllLines(IO.Path.Combine(folder, "error.log"), {String.Format("-----{0:d} {0:HH:mm:ss}-----", DateTime.Now)})

        Dim ex As Exception = e.Exception
        While Not ex Is Nothing
            IO.File.AppendAllLines(IO.Path.Combine(folder, "error.log"), {e.Exception.Message, e.Exception.StackTrace.ToString()})
            ex = ex.InnerException
        End While
    End Sub
End Class
