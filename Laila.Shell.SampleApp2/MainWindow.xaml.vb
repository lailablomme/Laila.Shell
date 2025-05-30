﻿Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Controls.Primitives
Imports Laila.MetroWindow.Data
Imports Laila.Shell.Controls
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Class MainWindow
    Private Const WINDOWPOSITION_FILENAME As String = "Laila.Shell.SampleApp.WindowPosition.dat"

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AddHandler Me.PreviewKeyUp,
            Sub(s As Object, e As KeyEventArgs)
                focusTextBlock.Text = Keyboard.FocusedElement.ToString() & " " & If(TypeOf Keyboard.FocusedElement Is FrameworkElement, CType(Keyboard.FocusedElement, FrameworkElement).Name, "")
            End Sub
        AddHandler Me.PreviewKeyDown,
            Sub(s As Object, e As KeyEventArgs)
                focusTextBlock.Text = Keyboard.FocusedElement.ToString() & " " & If(TypeOf Keyboard.FocusedElement Is FrameworkElement, CType(Keyboard.FocusedElement, FrameworkElement).Name, "")
            End Sub
    End Sub

    Public Overrides Sub OnLoadPosition()
        ' load position from disk
        If File.Exists(Path.Combine(Path.GetTempPath(), WINDOWPOSITION_FILENAME)) Then
            Me.Position = WindowPositionData.Deserialize(File.ReadAllText(Path.Combine(Path.GetTempPath(), WINDOWPOSITION_FILENAME)))
        End If
    End Sub

    Public Overrides Sub OnSavePosition()
        ' write position to disk
        IO.File.WriteAllText(Path.Combine(Path.GetTempPath(), WINDOWPOSITION_FILENAME), Me.Position.Serialize())
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

    Private _newItemMenu As NewItemMenu
    Private Sub NewItemMenuButton_Checked(sender As Object, e As RoutedEventArgs)
        Dim button As ToggleButton = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        If Not menus.NewItemMenu Is Nothing Then
            If Not menus.NewItemMenu.Equals(_newItemMenu) Then
                AddHandler menus.NewItemMenu.Closed,
                    Sub(s2 As Object, e2 As EventArgs)
                        button.IsChecked = False
                    End Sub
                _newItemMenu = menus.NewItemMenu
            End If
            menus.NewItemMenu.Placement = Primitives.PlacementMode.Bottom
            menus.NewItemMenu.PlacementTarget = button
            menus.NewItemMenu.IsOpen = True
        End If
    End Sub

    Private Sub CutButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        Clipboard.CutFiles(menus.SelectedItems)
        menus.UpdateButtons()
    End Sub

    Private Sub CopyButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        Clipboard.CopyFiles(menus.SelectedItems)
        menus.UpdateButtons()
    End Sub

    Private Sub PasteButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        Clipboard.PasteFiles(menus.Folder)
    End Sub

    Private Sub RenameButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim folderView As Laila.Shell.Controls.FolderView = button.Tag
        folderView.DoRename(folderView.SelectedItems(0))
    End Sub

    Private Sub ShareButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        Menus.DoShare(menus.SelectedItems)
    End Sub

    Private Sub DeleteButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button As Button = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        Menus.DoDelete(menus.SelectedItems)
    End Sub

    Private Sub SortMenuButton_Checked(sender As Object, e As RoutedEventArgs)
        Dim button As ToggleButton = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        Dim sortMenu As SortMenu = New SortMenu() With {.Folder = menus.Folder}
        AddHandler sortMenu.Closed,
            Sub(s2 As Object, e2 As EventArgs)
                button.IsChecked = False
            End Sub
        sortMenu.Placement = Primitives.PlacementMode.Bottom
        sortMenu.PlacementTarget = button
        sortMenu.IsOpen = True
    End Sub

    Private Sub ViewMenuButton_Checked(sender As Object, e As RoutedEventArgs)
        Dim button As ToggleButton = sender
        Dim menus As Laila.Shell.Controls.Menus = button.Tag
        Dim viewMenu As ViewMenu = New ViewMenu() With {.Folder = menus.Folder}
        AddHandler viewMenu.Closed,
            Sub(s2 As Object, e2 As EventArgs)
                button.IsChecked = False
            End Sub
        viewMenu.Placement = Primitives.PlacementMode.Bottom
        viewMenu.PlacementTarget = button
        viewMenu.IsOpen = True
    End Sub

    Private _savedPreviewWidth As GridLength
    Private Sub PreviewMenuButton_Checked(sender As Object, e As RoutedEventArgs)
        Dim button As ToggleButton = sender
        Dim previewer As Previewer = button.Tag
        Dim splitter As GridSplitter = previewer.Tag
        Dim grid As Grid = UIHelper.GetParentOfType(Of Grid)(previewer)
        previewer.Visibility = Visibility.Visible
        splitter.Visibility = Visibility.Visible
        grid.ColumnDefinitions(3).Width = New GridLength(5, GridUnitType.Pixel)
        grid.ColumnDefinitions(4).Width = _savedPreviewWidth
    End Sub

    Private Sub PreviewMenuButton_Unchecked(sender As Object, e As RoutedEventArgs)
        Dim button As ToggleButton = sender
        Dim previewer As Previewer = button.Tag
        Dim splitter As GridSplitter = previewer.Tag
        Dim grid As Grid = UIHelper.GetParentOfType(Of Grid)(previewer)
        previewer.Visibility = Visibility.Collapsed
        splitter.Visibility = Visibility.Collapsed
        grid.ColumnDefinitions(3).Width = New GridLength(0, GridUnitType.Pixel)
        _savedPreviewWidth = grid.ColumnDefinitions(4).Width
        grid.ColumnDefinitions(4).Width = New GridLength(0, GridUnitType.Pixel)
    End Sub
End Class
