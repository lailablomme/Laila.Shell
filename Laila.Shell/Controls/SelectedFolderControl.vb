Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Helpers
Imports System.Windows.Controls.Primitives
Imports System.Threading
Imports System.Windows.Input
Imports System.Windows.Data
Imports System.ComponentModel
Imports Laila.Shell.Themes

Namespace Controls
    Public Class SelectedFolderControl
        Inherits StackPanel
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(SelectedFolderControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColor", GetType(Boolean), GetType(SelectedFolderControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColorOverride", GetType(Boolean?), GetType(SelectedFolderControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged))
        Public Shared ReadOnly IsTabStopProperty As DependencyProperty = DependencyProperty.Register("IsTabStop", GetType(Boolean), GetType(SelectedFolderControl), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ColorsProperty As DependencyProperty = DependencyProperty.Register("Colors", GetType(StandardColors), GetType(SelectedFolderControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private ReadOnly _lock As New SemaphoreSlim(1, 1)
        Private _visibleFolders As List(Of Folder) = New List(Of Folder)
        Private _menu As RightClickMenu
        Private disposedValue As Boolean

        Shared Sub New()
            ' Override the metadata for MaxWidthProperty
            MaxWidthProperty.OverrideMetadata(GetType(SelectedFolderControl),
                New FrameworkPropertyMetadata(Double.PositiveInfinity, AddressOf OnMaxWidthChanged))
        End Sub

        Private Shared Sub OnMaxWidthChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim sfc = TryCast(d, SelectedFolderControl)
            Dim __ = sfc.ShowNavigationButtons(sfc.Folder, False)
        End Sub

        Public Sub New()
            Me.Colors = New StandardColors()
            Me.Orientation = Orientation.Horizontal
            Me.Focusable = False
            AddHandler Me.Loaded,
                Sub(sender As Object, e As EventArgs)
                    Dim __ = ShowNavigationButtons(Me.Folder, False)
                End Sub
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Shell.AddToControlCache(Me)

            AddHandler Shell.Settings.PropertyChanged,
                Sub(s As Object, e As PropertyChangedEventArgs)
                    Select Case e.PropertyName
                        Case "DoShowEncryptedOrCompressedFilesInColor"
                            setDoShowEncryptedOrCompressedFilesInColor()
                    End Select
                End Sub
            setDoShowEncryptedOrCompressedFilesInColor()
        End Sub

        Public Async Function ShowNavigationButtons(folder As Folder, isWithDelay As Boolean) As Task
            If folder Is Nothing Then
                Me.Children.Clear()
                Dim dummyButton As Button = New Button()
                dummyButton.Style = TryFindResource("lailaShell_AddressBarButtonStyle")
                dummyButton.Visibility = Visibility.Hidden
                dummyButton.Content = "X"
                Me.Children.Add(dummyButton)
                Return
            End If
            'If isWithDelay Then Await Task.Delay(150)

            Await _lock.WaitAsync()
            Try
                Dim previousVisibleFolders As List(Of Folder) = _visibleFolders.ToList()
                _visibleFolders.Clear()

                folder.IsVisibleInAddressBar = True

                Dim panels As List(Of StackPanel) = New List(Of StackPanel)()
                Dim totalWidth As Double = 0
                Dim buttonStyle As Style = TryFindResource("lailaShell_AddressBarButtonStyle")
                Dim chevronButtonStyle As Style = TryFindResource("lailaShell_AddressBarChevronButtonStyle")
                Dim moreButtonStyle As Style = TryFindResource("lailaShell_AddressBarMoreButtonStyle")
                Dim panelStyle As Style = TryFindResource("lailaShell_AddressBarPanelContainerStyle")

                Dim standardPanel As StackPanel = New StackPanel() With {.Orientation = Orientation.Horizontal, .Focusable = False}
                Dim computerImage As Image = New Image() With {.Width = 16, .Height = 16, .VerticalAlignment = VerticalAlignment.Center, .Margin = New Thickness(4, 0, 4, 0)}
                'computerImage.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/monitor16.png")
                standardPanel.Children.Add(computerImage)
                Dim specialFoldersButton As ToggleButton = New ToggleButton()
                If Not chevronButtonStyle Is Nothing Then specialFoldersButton.Style = chevronButtonStyle
                Dim specialFoldersContextMenu As ContextMenu = New ContextMenu()
                specialFoldersContextMenu.Colors = Me.Colors
                specialFoldersContextMenu.PlacementTarget = specialFoldersButton
                specialFoldersContextMenu.Placement = Primitives.PlacementMode.Bottom
                For Each specialFolder In Shell.GetSpecialFolders().Values.OrderBy(Function(f) f.DisplayName)
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
                standardPanel.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
                totalWidth += standardPanel.DesiredSize.Width

                Dim moreButton As ToggleButton = New ToggleButton()
                If Not moreButtonStyle Is Nothing Then moreButton.Style = moreButtonStyle
                moreButton.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
                totalWidth += moreButton.DesiredSize.Width

                Dim currentFolder As Folder = folder
                While Not currentFolder Is Nothing
                    Dim panel As StackPanel = New StackPanel() With {.Orientation = Orientation.Horizontal}
                    Dim folderButton As Button = New Button()
                    If Not buttonStyle Is Nothing Then folderButton.Style = buttonStyle
                    folderButton.Content = currentFolder.DisplayName
                    folderButton.Tag = currentFolder
                    folderButton.SetBinding(Button.ForegroundProperty, New Binding("Colors.Foreground") With {.Source = Me})
                    _visibleFolders.Add(currentFolder)
                    currentFolder.IsVisibleInAddressBar = True
                    AddHandler folderButton.Click,
                        Sub(s As Object, e As EventArgs)
                            Using Shell.OverrideCursor(Cursors.Wait)
                                CType(folderButton.Tag, Folder).LastScrollOffset = New Point()
                                Me.Folder = folderButton.Tag
                            End Using
                        End Sub
                    panel.Children.Add(folderButton)
                    Dim subFoldersButton As ToggleButton = New ToggleButton()
                    subFoldersButton.Tag = currentFolder
                    If Not chevronButtonStyle Is Nothing Then subFoldersButton.Style = chevronButtonStyle
                    panel.Children.Add(subFoldersButton)
                    subFoldersButton.Visibility = If(currentFolder.Equals(folder),
                        If(currentFolder.Items.ToList().Where(Function(i) TypeOf i Is Folder).Count = 0, Visibility.Collapsed, Visibility.Visible),
                        Visibility.Visible)
                    Dim func As Func(Of ToggleButton, Task) =
                        Async Function(button As ToggleButton) As Task
                            Dim subFolder As Folder = Nothing
                            Dim subFoldersContextMenu As ItemsContextMenu = Nothing

                            UIHelper.OnUIThread(
                                Sub()
                                    subFolder = button.Tag

                                    subFoldersContextMenu = New ItemsContextMenu()
                                    subFoldersContextMenu.Colors = Me.Colors
                                    subFoldersContextMenu.PlacementTarget = button
                                    subFoldersContextMenu.Placement = Primitives.PlacementMode.Bottom
                                    subFoldersContextMenu.DoShowEncryptedOrCompressedFilesInColorOverride = Me.DoShowEncryptedOrCompressedFilesInColorOverride
                                    Dim view As ListCollectionView = New ListCollectionView(subFolder.Items)
                                    view.Filter = Function(i) TypeOf i Is Folder
                                    view.SortDescriptions.Add(New SortDescription() With {
                                        .PropertyName = "ItemNameDisplaySortValue",
                                        .Direction = ListSortDirection.Ascending
                                    })
                                    subFoldersContextMenu.ItemsSource = view
                                End Sub)

                            AddHandler subFoldersContextMenu.ItemClicked,
                                Async Sub(clickedItem As Item, e2 As EventArgs)
                                    If TypeOf clickedItem Is Folder Then
                                        CType(clickedItem, Folder).LastScrollOffset = New Point()
                                        Me.Folder = clickedItem
                                    Else
                                        If Not _menu Is Nothing Then
                                            _menu.Dispose()
                                        End If
                                        _menu = New RightClickMenu() With {
                                            .Folder = subFolder,
                                            .SelectedItems = {clickedItem},
                                            .IsDefaultOnly = True
                                        }
                                        Await _menu.Make()
                                        Await _menu.InvokeCommand(_menu.DefaultId)
                                    End If
                                End Sub
                            AddHandler subFoldersContextMenu.Closed,
                                Sub(s2 As Object, e2 As EventArgs)
                                    button.IsChecked = False
                                End Sub

                            AddHandler button.Checked,
                                Sub(s As Object, e As EventArgs)
                                    Dim ___ = subFolder.GetItemsAsync()
                                    subFoldersContextMenu.IsOpen = True
                                End Sub

                            If subFolder.Equals(folder) Then
                                If (Not TypeOf subFolder Is SearchFolder AndAlso
                                    (Await subFolder.GetItemsAsync()).Where(Function(i) TypeOf i Is Folder).Count > 0) Then
                                    UIHelper.OnUIThread(
                                        Sub()
                                            button.Visibility = Visibility.Visible
                                        End Sub)
                                Else
                                    UIHelper.OnUIThread(
                                        Sub()
                                            button.Visibility = Visibility.Collapsed
                                        End Sub)
                                End If
                            End If
                        End Function
                    Dim __ = Task.Run(Sub() func(subFoldersButton))

                    panel.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
                    If totalWidth + panel.DesiredSize.Width < Me.MaxWidth Then
                        totalWidth += panel.DesiredSize.Width
                        panels.Add(panel)
                    Else
                        Exit While
                    End If
                    If TypeOf currentFolder Is SearchFolder OrElse currentFolder.LogicalParent Is Nothing OrElse Item.ArePathsEqual(currentFolder.LogicalParent.FullPath, Shell.Desktop.FullPath) Then
                        computerImage.Source = currentFolder.Icon(16)
                        currentFolder = Nothing
                    End If
                    If Not currentFolder Is Nothing Then
                        currentFolder = currentFolder.LogicalParent
                        If Not currentFolder Is Nothing AndAlso Item.ArePathsEqual(currentFolder.FullPath, Shell.Desktop.FullPath) Then
                            currentFolder = Nothing
                        End If
                    End If
                End While

                Dim hasMoreButton As Boolean = False
                If Not currentFolder Is Nothing Then
                    moreButton.Tag = currentFolder
                    hasMoreButton = True

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
                        If TypeOf currentFolder Is SearchFolder OrElse currentFolder.LogicalParent Is Nothing OrElse Item.ArePathsEqual(currentFolder.LogicalParent.FullPath, Shell.Desktop.FullPath) Then
                            computerImage.Source = currentFolder.Icon(16)
                            currentFolder = Nothing
                        End If
                        If Not currentFolder Is Nothing Then
                            currentFolder = currentFolder.LogicalParent
                            If Not currentFolder Is Nothing AndAlso Item.ArePathsEqual(currentFolder.FullPath, Shell.Desktop.FullPath) Then
                                currentFolder = Nothing
                            End If
                        End If
                    End While

                    AddHandler moreButton.Checked,
                        Sub(s As Object, e As EventArgs)
                            moreContextMenu.IsOpen = True
                        End Sub
                End If

                If hasMoreButton Then
                    Dim morePanel As StackPanel = New StackPanel() With {.Orientation = Orientation.Horizontal}
                    morePanel.Children.Add(moreButton)
                    panels.Add(morePanel)
                End If
                panels.Add(standardPanel)
                panels.Reverse()
                If (folder Is Nothing AndAlso Me.Folder Is Nothing) OrElse (Not folder Is Nothing AndAlso folder.Equals(Me.Folder)) Then ' if we're not too late with this async
                    Me.Children.Clear()
                    For Each panel In panels
                        Dim container As Button = New Button() With {.Content = panel, .Style = panelStyle, .IsTabStop = False}
                        Me.Children.Add(container)
                        AddHandler container.PreviewKeyDown,
                            Sub(s As Object, e As KeyEventArgs)
                                Select Case e.Key
                                    Case Key.Left
                                        Dim idx As Integer = Me.Children.IndexOf(container)
                                        If idx > 0 Then
                                            CType(Me.Children(idx - 1), Button).Focus()
                                        End If
                                        e.Handled = True
                                    Case Key.Right
                                        Dim idx As Integer = Me.Children.IndexOf(container)
                                        If idx < Me.Children.Count - 1 Then
                                            CType(Me.Children(idx + 1), Button).Focus()
                                        End If
                                        e.Handled = True
                                    Case Key.Space, Key.Enter
                                        If TypeOf panel.Children(0) Is Button Then
                                            CType(panel.Children(0), Button).RaiseEvent(New RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
                                        End If
                                        e.Handled = True
                                    Case Key.Down
                                        If TypeOf panel.Children(panel.Children.Count - 1) Is ToggleButton Then
                                            CType(panel.Children(panel.Children.Count - 1), ToggleButton).IsChecked = True
                                        End If
                                        e.Handled = True
                                End Select
                            End Sub
                    Next
                    If Me.Children.Count > 0 Then CType(Me.Children(0), Button).IsTabStop = Me.IsTabStop
                End If

                For Each f In previousVisibleFolders.Where(Function(f2) Not _visibleFolders.Contains(f2))
                    f.IsVisibleInAddressBar = False
                Next
            Finally
                _lock.Release()
            End Try
        End Function

        Public ReadOnly Property VisibleFolders As List(Of Folder)
            Get
                Return _visibleFolders
            End Get
        End Property

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim sfc As SelectedFolderControl = TryCast(d, SelectedFolderControl)
            Dim __ = sfc.ShowNavigationButtons(e.NewValue, True)
        End Sub

        Public Property DoShowEncryptedOrCompressedFilesInColor As Boolean
            Get
                Return GetValue(DoShowEncryptedOrCompressedFilesInColorProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowEncryptedOrCompressedFilesInColorProperty, value)
            End Set
        End Property

        Private Sub setDoShowEncryptedOrCompressedFilesInColor()
            If Me.DoShowEncryptedOrCompressedFilesInColorOverride.HasValue Then
                Me.DoShowEncryptedOrCompressedFilesInColor = Me.DoShowEncryptedOrCompressedFilesInColorOverride.Value
            Else
                Me.DoShowEncryptedOrCompressedFilesInColor = Shell.Settings.DoShowEncryptedOrCompressedFilesInColor
            End If
        End Sub

        Public Property DoShowEncryptedOrCompressedFilesInColorOverride As Boolean?
            Get
                Return GetValue(DoShowEncryptedOrCompressedFilesInColorOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowEncryptedOrCompressedFilesInColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim sfc As SelectedFolderControl = TryCast(d, SelectedFolderControl)
            sfc.setDoShowEncryptedOrCompressedFilesInColor()
        End Sub

        Public Overloads Property IsTabStop As Boolean
            Get
                Return GetValue(IsTabStopProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsTabStopProperty, value)
            End Set
        End Property

        Public Property Colors As StandardColors
            Get
                Return GetValue(ColorsProperty)
            End Get
            Protected Set(ByVal value As StandardColors)
                SetCurrentValue(ColorsProperty, value)
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

                    Shell.RemoveFromControlCache(Me)
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