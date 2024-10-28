Class MainWindow
    Private _model As MainWindowViewModel

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _model = New MainWindowViewModel(Me)
        Me.DataContext = _model
    End Sub

    Private Sub BackButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim navigation As Laila.Shell.Controls.Navigation = button.Tag
        navigation.Back()
    End Sub

    Private Sub ForwardButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim navigation As Laila.Shell.Controls.Navigation = button.Tag
        navigation.Forward()
    End Sub

    Private Sub UpButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim navigation As Laila.Shell.Controls.Navigation = button.Tag
        navigation.Up()
    End Sub

    Private Sub RefreshButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim navigation As Laila.Shell.Controls.Navigation = button.Tag
        navigation.Folder.RefreshItemsAsync()
    End Sub
End Class
