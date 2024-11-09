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

    Private Sub newItemMenuButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        If Not menus.NewItemMenu Is Nothing Then
            menus.NewItemMenu.Placement = Primitives.PlacementMode.Bottom
            menus.NewItemMenu.PlacementTarget = button
            menus.NewItemMenu.IsOpen = True
        Else
            MsgBox("no new items")
        End If
    End Sub
End Class
