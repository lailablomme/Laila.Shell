﻿Imports Laila.AutoCompleteTextBox
Imports Laila.Shell.Helpers
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Environment
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media

Namespace Controls
    Public Class AddressBar
        Inherits Laila.AutoCompleteTextBox.AutoCompleteTextBox
        Implements IDisposable

        Private Const INVALID_VALUE As String = "5e979b53-746b-4a0c-9f5f-00fdd22c91d8"

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(AddressBar), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private PART_NavigationButtonsPanel As StackPanel
        Private PART_NavigationButtons As Border
        Private PART_ClickToEdit As Border
        Private _visibleFolders As List(Of Folder) = New List(Of Folder)
        Private disposedValue As Boolean
        Private ReadOnly _lock As New SemaphoreSlim(1, 1)

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(AddressBar), New FrameworkPropertyMetadata(GetType(AddressBar)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_NavigationButtonsPanel = Template.FindName("PART_NavigationButtonsPanel", Me)
            PART_NavigationButtons = Template.FindName("PART_NavigationButtons", Me)
            PART_ClickToEdit = Template.FindName("PART_ClickToEdit", Me)

            AddHandler Me.SizeChanged,
                Sub(s As Object, e As SizeChangedEventArgs)
                    Me.ShowNavigationButtons(Me.Folder, False)
                End Sub
            AddHandler Me.PART_ClickToEdit.MouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    _isSettingTextInternally = True
                    Me.Text = Me.Folder.AddressBarDisplayName
                    _isSettingTextInternally = False
                    PART_NavigationButtons.Visibility = Visibility.Hidden
                    Me.PART_TextBox.IsEnabled = True
                    Me.PART_TextBox.SelectionStart = Me.Text.Length
                    Me.PART_TextBox.Focus()
                End Sub
            AddHandler Me.PART_TextBox.PreviewKeyDown,
                Sub(s As Object, e As KeyEventArgs)
                    If Not Me.IsDropDownOpen OrElse Me.IsLoadingSuggestions Then
                        Select Case e.Key
                            Case Key.Enter
                                If Not Me.IsDirty Then
                                    Me.OnItemSelected()
                                End If
                            Case Key.Escape
                                If Not Me.IsDirty Then
                                    Me.Cancel()
                                End If
                        End Select
                    End If
                End Sub
        End Sub

        Protected Overrides Sub OnItemSelected()
            Me.PART_TextBox.IsEnabled = False

            If INVALID_VALUE.Equals(Me.SelectedValue) Then
                Dim item As Item = Item.FromParsingName(Me.PART_TextBox.Text, Nothing)
                If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                    CType(item, Folder).LastScrollOffset = New Point()
                    Me.Folder = item
                    Me.SelectedItem = item
                    CType(item, Folder).IsVisibleInAddressBar = True
                Else
                    System.Media.SystemSounds.Asterisk.Play()
                    Me.Cancel()
                    Me.IsLoading = True
                    Me.ShowNavigationButtons(Me.Folder, False)
                End If
            ElseIf Not Me.SelectedItem Is Nothing Then
                Dim doShow As Boolean = Me.SelectedItem.Equals(Me.Folder)
                CType(Me.SelectedItem, Folder).LastScrollOffset = New Point()
                Me.Folder = Me.SelectedItem
                If doShow Then
                    Me.IsLoading = True
                    Me.ShowNavigationButtons(Me.Folder, False)
                End If
            Else
                Me.Cancel()
            End If
        End Sub

        Protected Overrides Sub Cancel()
            MyBase.Cancel()

            Me.PART_TextBox.IsEnabled = False
            Me.IsLoading = True
            Me.ShowNavigationButtons(Me.Folder, False)

            clearSuggestionItems()
        End Sub

        Private Sub clearSuggestionItems()
            Dim provider As AddressBarSuggestionProvider = Me.Provider
            provider._lock.Wait()
            Try
                If Not provider._items Is Nothing Then
                    For Each item In provider._items
                        item.MaybeDispose()
                    Next
                End If
                provider._items = Nothing
            Finally
                provider._lock.Release()
            End Try
        End Sub

        Protected Overrides Sub TextBox_LostFocus(s As Object, e As RoutedEventArgs)
            Me.Cancel()
        End Sub

        Public Async Function ShowNavigationButtons(folder As Folder, isWithDelay As Boolean) As Task
            If isWithDelay Then Await Task.Delay(150)

            Await _lock.WaitAsync()
            Try
                Dim previousVisibleFolders As List(Of Folder) = _visibleFolders.ToList()
                _visibleFolders.Clear()

                folder.IsVisibleInAddressBar = True

                Dim buttons As List(Of StackPanel) = New List(Of StackPanel)()
                Dim totalWidth As Double = 0
                Dim buttonStyle As Style = TryFindResource("lailaShell_AddressBarButtonStyle")
                Dim chevronButtonStyle As Style = TryFindResource("lailaShell_AddressBarChevronButtonStyle")
                Dim moreButtonStyle As Style = TryFindResource("lailaShell_AddressBarMoreButtonStyle")

                Dim standardPanel As StackPanel = New StackPanel() With {.Orientation = Orientation.Horizontal, .Focusable = False}
                Dim computerImage As Image = New Image() With {.Width = 16, .Height = 16, .VerticalAlignment = VerticalAlignment.Center, .Margin = New Thickness(4, 0, 4, 0)}
                'computerImage.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/monitor16.png")
                standardPanel.Children.Add(computerImage)
                Dim specialFoldersButton As ToggleButton = New ToggleButton()
                If Not chevronButtonStyle Is Nothing Then specialFoldersButton.Style = chevronButtonStyle
                Dim specialFoldersContextMenu As ContextMenu = New ContextMenu()
                specialFoldersContextMenu.PlacementTarget = specialFoldersButton
                specialFoldersContextMenu.Placement = Primitives.PlacementMode.Bottom
                For Each specialFolder In Shell.SpecialFolders.Values.OrderBy(Function(f) f.DisplayName)
                    Dim specialFoldersMenuItem As MenuItem = New MenuItem()
                    specialFoldersMenuItem.Header = specialFolder.DisplayName
                    specialFoldersMenuItem.Tag = specialFolder
                    specialFoldersMenuItem.Icon = New Image() With {.Width = 16, .Height = 16, .Source = specialFolder.Icon(16)}
                    AddHandler specialFoldersMenuItem.Click,
                        Sub(s As Object, e As EventArgs)
                            CType(specialFoldersMenuItem.Tag, Folder).LastScrollOffset = New Point()
                            Me.Folder = specialFoldersMenuItem.Tag
                        End Sub
                    specialFoldersContextMenu.Items.Add(specialFoldersMenuItem)
                Next
                AddHandler specialFoldersButton.Checked,
                            Sub(s As Object, e As EventArgs)
                                specialFoldersContextMenu.IsOpen = True
                            End Sub
                AddHandler specialFoldersContextMenu.Closed,
                            Sub(s As Object, e As EventArgs)
                                specialFoldersButton.IsChecked = False
                            End Sub
                standardPanel.Children.Add(specialFoldersButton)

                Dim currentFolder As Folder = folder
                While Not currentFolder Is Nothing
                    Dim panel As StackPanel = New StackPanel() With {.Orientation = Orientation.Horizontal}
                    Dim folderButton As Button = New Button()
                    If Not buttonStyle Is Nothing Then folderButton.Style = buttonStyle
                    folderButton.Content = currentFolder.DisplayName
                    folderButton.Tag = currentFolder
                    _visibleFolders.Add(currentFolder)
                    currentFolder.IsVisibleInAddressBar = True
                    AddHandler folderButton.Click,
                    Sub(s As Object, e As EventArgs)
                        CType(folderButton.Tag, Folder).LastScrollOffset = New Point()
                        Me.Folder = folderButton.Tag
                    End Sub
                    panel.Children.Add(folderButton)
                    Dim subFoldersButton As ToggleButton = New ToggleButton()
                    subFoldersButton.IsEnabled = False
                    subFoldersButton.Tag = currentFolder
                    If Not chevronButtonStyle Is Nothing Then subFoldersButton.Style = chevronButtonStyle
                    panel.Children.Add(subFoldersButton)
                    Dim func As Func(Of ToggleButton, Task) =
                        Async Function(button As ToggleButton) As Task
                            Dim subFolder As Folder
                            UIHelper.OnUIThread(
                                Sub()
                                    subFolder = button.Tag
                                End Sub)
                            If (Await subFolder.GetItemsAsync()).Where(Function(i) TypeOf i Is Folder).Count > 0 _
                                OrElse Not subFolder.Equals(folder) Then
                                AddHandler button.Checked,
                                     Sub(s As Object, e As EventArgs)
                                         Dim subFoldersContextMenu As ItemsContextMenu = New ItemsContextMenu()
                                         subFoldersContextMenu.PlacementTarget = button
                                         subFoldersContextMenu.Placement = Primitives.PlacementMode.Bottom
                                         Dim view As ListCollectionView = New ListCollectionView(subFolder.Items)
                                         view.Filter = Function(i) TypeOf i Is Folder
                                         view.SortDescriptions.Add(New SortDescription() With {
                                             .PropertyName = "ItemNameDisplaySortValue",
                                             .Direction = ListSortDirection.Ascending
                                         })
                                         subFoldersContextMenu.ItemsSource = view

                                         AddHandler subFoldersContextMenu.ItemClicked,
                                             Sub(clickedSubFolder As Folder, e2 As EventArgs)
                                                 clickedSubFolder.LastScrollOffset = New Point()
                                                 Me.Folder = clickedSubFolder
                                             End Sub
                                         AddHandler subFoldersContextMenu.Closed,
                                             Sub(s2 As Object, e2 As EventArgs)
                                                 button.IsChecked = False
                                             End Sub

                                         subFoldersContextMenu.IsOpen = True
                                     End Sub
                            ElseIf subFolder.Equals(folder) Then
                                UIHelper.OnUIThread(
                                    Sub()
                                        button.Visibility = Visibility.Collapsed
                                    End Sub)
                            End If
                            UIHelper.OnUIThread(
                                Sub()
                                    button.IsEnabled = True
                                End Sub)
                        End Function
                    Task.Run(Sub() func(subFoldersButton))
                    currentFolder.MaybeDispose()

                    panel.Measure(New Size(1000, 1000))
                    If totalWidth + panel.DesiredSize.Width < Me.ActualWidth - 60 Then
                        totalWidth += panel.DesiredSize.Width
                        buttons.Add(panel)
                    Else
                        Exit While
                    End If
                    Dim parent As Folder = If(currentFolder._logicalParent Is Nothing, currentFolder.GetParent(), currentFolder._logicalParent)
                    If parent Is Nothing OrElse parent.FullPath = Shell.Desktop.FullPath Then
                        computerImage.Source = currentFolder.Icon(16)
                    End If
                    currentFolder = parent
                    If Not currentFolder Is Nothing AndAlso currentFolder.FullPath = Shell.Desktop.FullPath Then
                        currentFolder.MaybeDispose()
                        currentFolder = Nothing
                    End If
                End While

                If Not currentFolder Is Nothing Then
                    Dim moreButton As ToggleButton = New ToggleButton()
                    If Not moreButtonStyle Is Nothing Then moreButton.Style = moreButtonStyle
                    moreButton.Tag = currentFolder
                    standardPanel.Children.Add(moreButton)
                    standardPanel.Measure(New Size(1000, 1000))

                    Dim moreContextMenu As ContextMenu = New ContextMenu()
                    moreContextMenu.PlacementTarget = moreButton
                    moreContextMenu.Placement = Primitives.PlacementMode.Bottom

                    AddHandler moreContextMenu.Closed,
                            Sub(s2 As Object, e2 As EventArgs)
                                moreButton.IsChecked = False
                            End Sub

                    While Not currentFolder Is Nothing
                        Dim moreMenuItem As MenuItem = New MenuItem()
                        moreMenuItem.Header = currentFolder.DisplayName
                        moreMenuItem.Tag = currentFolder
                        _visibleFolders.Add(currentFolder)
                        currentFolder.IsVisibleInAddressBar = True
                        moreMenuItem.Icon = New Image() With {.Width = 16, .Height = 16, .Source = currentFolder.Icon(16)}
                        AddHandler moreMenuItem.Click,
                                Sub(s2 As Object, e2 As EventArgs)
                                    CType(moreMenuItem.Tag, Folder).LastScrollOffset = New Point()
                                    Me.Folder = moreMenuItem.Tag
                                End Sub
                        moreContextMenu.Items.Add(moreMenuItem)
                        Dim parent As Folder = If(currentFolder._logicalParent Is Nothing, currentFolder.GetParent(), currentFolder._logicalParent)
                        If parent Is Nothing OrElse parent.FullPath = Shell.Desktop.FullPath Then
                            computerImage.Source = currentFolder.Icon(16)
                        End If
                        currentFolder = parent
                        If Not currentFolder Is Nothing AndAlso currentFolder.FullPath = Shell.Desktop.FullPath Then
                            currentFolder.MaybeDispose()
                            currentFolder = Nothing
                        End If
                    End While

                    AddHandler moreButton.Checked,
                    Sub(s As Object, e As EventArgs)
                        moreContextMenu.IsOpen = True
                    End Sub
                End If

                buttons.Add(standardPanel)
                buttons.Reverse()
                If (folder Is Nothing AndAlso Me.Folder Is Nothing) OrElse (Not folder Is Nothing AndAlso folder.Equals(Me.Folder)) Then ' if we're not too late with this async
                    Me.PART_NavigationButtonsPanel.Children.Clear()
                    For Each button In buttons
                        Me.PART_NavigationButtonsPanel.Children.Add(button)
                    Next
                    Me.PART_NavigationButtons.Visibility = Visibility.Visible
                End If

                Me.IsLoading = False
                clearSuggestionItems()

                For Each f In previousVisibleFolders.Where(Function(f2) Not _visibleFolders.Contains(f2))
                    f.IsVisibleInAddressBar = False
                Next
            Finally
                _lock.Release()
            End Try
        End Function

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim ab As AddressBar = TryCast(d, AddressBar)
            ab.SelectedItem = e.NewValue
            ab.IsLoading = True
            ab.ShowNavigationButtons(e.NewValue, True)
        End Sub

        Public Property IsLoading As Boolean
            Get
                Return GetValue(IsLoadingProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsLoadingProperty, value)
            End Set
        End Property

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    For Each f In _visibleFolders
                        f.IsVisibleInAddressBar = False
                    Next
                    _visibleFolders.Clear()

                    clearSuggestionItems()
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace