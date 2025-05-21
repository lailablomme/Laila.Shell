Imports System.Globalization
Imports System.Reflection
Imports System.Windows.Markup

Class Application
    Public Sub New()
        Dim culture As CultureInfo = CultureInfo.InstalledUICulture
        CultureInfo.DefaultThreadCurrentUICulture = culture
        CultureInfo.DefaultThreadCurrentCulture = culture

        ' Set default FlowDirection based on culture direction
        Dim direction = If(culture.TextInfo.IsRightToLeft, FlowDirection.RightToLeft, FlowDirection.LeftToRight)

        ' Apply to all windows (optional)
        FrameworkElement.LanguageProperty.OverrideMetadata(GetType(FrameworkElement),
            New FrameworkPropertyMetadata(XmlLanguage.GetLanguage(culture.IetfLanguageTag)))
        FrameworkElement.FlowDirectionProperty.OverrideMetadata(GetType(FrameworkElement),
            New FrameworkPropertyMetadata(direction))
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
        IO.File.AppendAllLines(getLogFileName(), {String.Format("-----{0:d} {0:HH:mm:ss}-----", DateTime.Now)})

        While Not ex Is Nothing
            IO.File.AppendAllLines(getLogFileName(), {$"{ex.Message} [{ex.GetType().ToString()}]", ex.StackTrace?.ToString()})
            ex = ex.InnerException
        End While
    End Sub


    Private Shared Function getLogFileName() As String
        Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
        If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
        Return IO.Path.Combine(path, "error.log")
    End Function
End Class
