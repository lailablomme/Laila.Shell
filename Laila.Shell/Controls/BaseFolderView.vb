Imports System.ComponentModel
Imports System.Media
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Namespace Controls
    Public MustInherit Class BaseFolderView
        Inherits Control
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ColumnsInProperty As DependencyProperty = DependencyProperty.Register("ColumnsIn", GetType(Behaviors.GridViewExtBehavior.ColumnsInData), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly IsSelectingProperty As DependencyProperty = DependencyProperty.Register("IsSelecting", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Friend Host As FolderView
        Friend PART_ListBox As System.Windows.Controls.ListBox
        Private PART_Grid As Grid
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _mouseItemOver As Item
        Private _timeSpentTimer As Timer
        Protected _scrollViewer As ScrollViewer
        Private _lastScrollOffset As Point
        Private _lastScrollSize As Size
        Private _isLoaded As Boolean
        Private _typeToSearchTimer As Timer
        Private _typeToSearchString As String = ""
        Private _menu As RightClickMenu
        Private _ignoreSelection As Boolean
        Private _checkBoxVisibility As Visibility = Visibility.Collapsed
        Private disposedValue As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(BaseFolderView), New FrameworkPropertyMetadata(GetType(BaseFolderView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_ListBox = Template.FindName("PART_ListView", Me)
            Me.PART_Grid = Template.FindName("PART_Grid", Me)

            'Me.PART_ListBox.Visibility = Visibility.Hidden

            If Not Me.Folder Is Nothing Then
                Me.MakeBinding(Me.Folder)
            End If

            AddHandler Shell.ShuttingDown,
                Sub(s As Object, e As EventArgs)
                    Me.Dispose()
                End Sub

            AddHandler PART_ListBox.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        _selectionHelper = New SelectionHelper(Of Item)(Me.PART_ListBox)
                        _selectionHelper.SelectionChanged =
                            Sub()
                                If Not Me.Folder Is Nothing AndAlso Not _ignoreSelection Then _
                                    Me.SelectedItems = _selectionHelper.SelectedItems
                            End Sub
                        _selectionHelper.SetSelectedItems(Me.SelectedItems)

                        _scrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(Me.PART_ListBox)(0)
                        AddHandler _scrollViewer.ScrollChanged,
                            Sub(s2 As Object, e2 As ScrollChangedEventArgs)
                                If Not Me.Folder Is Nothing Then
                                    _lastScrollOffset = New Point(_scrollViewer.HorizontalOffset, _scrollViewer.VerticalOffset)
                                    _lastScrollSize = New Size(_scrollViewer.ScrollableWidth, _scrollViewer.ScrollableHeight)
                                End If
                                'UIHelper.OnUIThreadAsync(
                                'Sub()
                                '    GC.Collect()
                                'End Sub)
                            End Sub

                        If Not Me.Folder Is Nothing Then
                            setGrouping(Me.Folder)
                        End If
                    End If
                End Sub

            AddHandler Me.PART_ListBox.PreviewMouseMove, AddressOf OnListViewPreviewMouseMove
            AddHandler Me.PART_ListBox.PreviewMouseDown, AddressOf OnListViewPreviewMouseButtonDown
            AddHandler Me.PART_ListBox.PreviewMouseUp, AddressOf OnListViewPreviewMouseButtonUp
            AddHandler Me.PART_ListBox.MouseLeave, AddressOf OnListViewMouseLeave
            AddHandler Me.PreviewKeyDown, AddressOf OnListViewKeyDown
            AddHandler Me.PreviewTextInput, AddressOf OnListViewTextInput
        End Sub

        Public Sub SetSelectedItemsSoft(items As IEnumerable(Of Item))
            _ignoreSelection = True
            _selectionHelper.SetSelectedItems(items)
            _ignoreSelection = False
        End Sub

        Public Property CheckBoxVisibility As Visibility
            Get
                Return _checkBoxVisibility
            End Get
            Set(value As Visibility)
                _checkBoxVisibility = value
            End Set
        End Property

        Private Sub OnListViewKeyDown(sender As Object, e As KeyEventArgs)
            If Not TypeOf e.OriginalSource Is TextBox Then
                If e.Key = Key.C AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) _
                AndAlso Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 Then
                    Clipboard.CopyFiles(Me.SelectedItems)
                ElseIf e.Key = Key.X AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) _
                AndAlso Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 Then
                    Clipboard.CutFiles(Me.SelectedItems)
                ElseIf e.Key = Key.Enter AndAlso Keyboard.Modifiers = ModifierKeys.None _
                AndAlso Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 Then
                    If TypeOf Me.SelectedItems(0) Is Folder Then
                        Me.Host.Folder = Me.SelectedItems(0)
                    Else
                        invokeDefaultCommand(Me.SelectedItems(0))
                    End If
                    e.Handled = True
                End If
            End If
        End Sub

        Private Sub OnListViewTextInput(sender As Object, e As TextCompositionEventArgs)
            If Not TypeOf e.OriginalSource Is TextBox Then
                If Not _typeToSearchTimer Is Nothing Then
                    _typeToSearchTimer.Dispose()
                End If

                _typeToSearchTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                _typeToSearchString = ""
                                _typeToSearchTimer.Dispose()
                                _typeToSearchTimer = Nothing
                            End Sub)
                    End Sub), Nothing, 650, Timeout.Infinite)

                _typeToSearchString &= e.Text
                Dim foundItem As Item =
                Me.Folder.Items.Skip(Me.Folder.Items.IndexOf(Me.SelectedItems(0)) + 1) _
                        .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                If foundItem Is Nothing Then
                    foundItem =
                    Me.Folder.Items.Take(Me.Folder.Items.IndexOf(Me.SelectedItems(0))) _
                            .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                End If
                If Not foundItem Is Nothing Then
                    Me.SelectedItems = {foundItem}
                    e.Handled = True
                Else
                    SystemSounds.Asterisk.Play()
                End If
            End If
        End Sub

        Private Sub invokeDefaultCommand(item As Item)
            getMenu(Me.Folder, {item}, True)
            _menu.Make()
            _menu.InvokeCommand(_menu.DefaultId)
        End Sub

        Private Sub OnListViewPreviewMouseMove(sender As Object, e As MouseEventArgs)
            Dim listBoxItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
            Dim overItem As Item = TryCast(listBoxItem?.DataContext, Item)
            If Not overItem Is Nothing AndAlso Not overItem.Equals(_mouseItemOver) Then
                Dim toolTip As String = overItem.InfoTip
                listBoxItem.ToolTip = If(String.IsNullOrWhiteSpace(toolTip), Nothing, toolTip)
                _mouseItemOver = overItem
            End If

            If Not _mouseItemDown Is Nothing AndAlso Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 _
                AndAlso (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                Dim currentPointDown As Point = e.GetPosition(Me)
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                    Drag.Start(Me.SelectedItems, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                End If
            End If
        End Sub

        Public Sub OnListViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(Me)

            ' this prevents a multiple selection getting replaced by the single clicked item
            If Not e.OriginalSource Is Nothing Then
                Dim listBoxItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
                Dim clickedItem As Item = TryCast(listBoxItem?.DataContext, Item)
                _mouseItemDown = clickedItem
                If clickedItem Is Nothing Then
                    Me.PART_ListBox.Focus()
                Else
                    listBoxItem.Focus()
                End If
                If e.LeftButton = MouseButtonState.Pressed AndAlso e.ClickCount = 2 AndAlso Not clickedItem Is Nothing Then
                    Using Shell.OverrideCursor(Cursors.Wait)
                        Me.SelectedItems = {clickedItem}
                        If TypeOf clickedItem Is Folder Then
                            CType(clickedItem, Folder).LastScrollOffset = New Point()
                            Me.Host.Folder = clickedItem
                            UIHelper.OnUIThread(
                                Sub()
                                End Sub, Threading.DispatcherPriority.Render)
                        Else
                            invokeDefaultCommand(clickedItem)
                        End If
                    End Using
                ElseIf e.LeftButton = MouseButtonState.Pressed AndAlso Not clickedItem Is Nothing Then
                    If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 _
                            AndAlso Me.SelectedItems.Contains(clickedItem) Then
                        Dim checkBox As CheckBox = UIHelper.GetParentOfType(Of CheckBox)(e.OriginalSource)
                        If Not checkBox Is Nothing AndAlso checkBox.IsChecked Then
                            checkBox.IsChecked = False
                            e.Handled = True
                        ElseIf Keyboard.Modifiers = ModifierKeys.None Then
                            e.Handled = True
                        End If
                    End If
                ElseIf e.RightButton = MouseButtonState.Pressed AndAlso
                        UIHelper.GetParentOfType(Of Primitives.ScrollBar)(e.OriginalSource) Is Nothing AndAlso
                        UIHelper.GetParentOfType(Of GridViewHeaderRowPresenter)(e.OriginalSource) Is Nothing Then

                    If (Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0 OrElse Not Me.SelectedItems.Contains(clickedItem)) _
                            AndAlso Not clickedItem Is Nothing Then
                        Me.SelectedItems = {clickedItem}
                    ElseIf clickedItem Is Nothing Then
                        Me.SelectedItems = Nothing
                    End If

                    Me.PART_ListBox.ContextMenu = getMenu(Me.Folder, Me.SelectedItems, False)
                    e.Handled = True
                ElseIf clickedItem Is Nothing AndAlso
                        UIHelper.GetParentOfType(Of System.Windows.Controls.Primitives.ScrollBar)(e.OriginalSource) Is Nothing Then
                    Me.SelectedItems = Nothing
                End If
            Else
                _mouseItemDown = Nothing
            End If
        End Sub

        Public Sub OnListViewPreviewMouseButtonUp(sender As Object, e As MouseButtonEventArgs)
            _mouseItemDown = Nothing
        End Sub

        Public Sub OnListViewMouseLeave(sender As Object, e As MouseEventArgs)
            _mouseItemDown = Nothing
        End Sub

        Private Function getMenu(folder As Folder, selectedItems As IEnumerable(Of Item), isDefaultOnly As Boolean) As RightClickMenu
            If Not _menu Is Nothing Then
                _menu.Dispose()
            End If

            _menu = New RightClickMenu()
            _menu.Folder = folder
            _menu.SelectedItems = selectedItems
            _menu.IsDefaultOnly = isDefaultOnly

            AddHandler _menu.CommandInvoked,
                Sub(s As Object, e2 As CommandInvokedEventArgs)
                    Select Case e2.Verb
                        Case "open"
                            If Not Me.SelectedItems Is Nothing _
                                    AndAlso Me.SelectedItems.Count = 1 _
                                    AndAlso TypeOf Me.SelectedItems(0) Is Folder Then
                                Me.Host.Folder = Me.SelectedItems(0)
                                e2.IsHandled = True
                            End If
                        Case "rename"
                            Me.DoRename(Me.SelectedItems(0))
                            e2.IsHandled = True
                        Case "laila.shell.(un)pin"
                            If e2.IsChecked Then
                                PinnedItems.PinItem(Me.SelectedItems(0).FullPath)
                            Else
                                PinnedItems.UnpinItem(Me.SelectedItems(0).FullPath)
                            End If
                            e2.IsHandled = True
                    End Select
                End Sub
            AddHandler _menu.RenameRequest,
                Async Sub(s As Object, e As RenameRequestEventArgs)
                    e.IsHandled = Await Me.DoRename(e.Pidl)
                End Sub

            Return _menu
        End Function

        Protected MustOverride Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                          ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)

        Public Async Function DoRename(pidl As Pidl) As Task(Of Boolean)
            If Not Me.Folder Is Nothing Then
                Dim item As Item = (Await Me.Folder.GetItemsAsync()).FirstOrDefault(Function(i) i.Pidl.Equals(pidl))
                If Not item Is Nothing Then
                    Me.SelectedItems = {item}
                    DoRename(item)
                    Return True
                End If
            End If
            Return False
        End Function

        Public Sub DoRename(item As Item)
            Me.PART_ListBox.ScrollIntoView(item)
            UIHelper.OnUIThread(
                Sub()
                End Sub, Threading.DispatcherPriority.ContextIdle)
            Dim listBoxItem As ListBoxItem = Me.PART_ListBox.ItemContainerGenerator.ContainerFromItem(item)
            DoRename(listBoxItem)
        End Sub

        Public Sub DoRename(listBoxItem As ListBoxItem)
            Dim point As Point, size As Size, textAlignment As TextAlignment, fontSize As Double
            Me.GetItemNameCoordinates(listBoxItem, textAlignment, point, size, fontSize)
            Menus.DoRename(point, size, textAlignment, fontSize, listBoxItem.DataContext, Me.PART_Grid)
        End Sub

        Public Overridable Property SelectedItems As IEnumerable(Of Item)
            Get
                Return GetValue(SelectedItemsProperty)
            End Get
            Set(value As IEnumerable(Of Item))
                SetCurrentValue(SelectedItemsProperty, value)
            End Set
        End Property

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Public Property IsLoading As Boolean
            Get
                Return GetValue(IsLoadingProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(IsLoadingProperty, value)
            End Set
        End Property

        Public Property IsSelecting As Boolean
            Get
                Return GetValue(IsSelectingProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(IsSelectingProperty, value)
            End Set
        End Property

        Protected Overridable Sub ClearBinding()
            If Not Me.PART_ListBox Is Nothing Then
                Me.PART_ListBox.ItemsSource = Nothing
            End If
        End Sub

        Protected Overridable Sub MakeBinding(folder As Folder)
            If Not Me.PART_ListBox Is Nothing Then
                setGrouping(folder)

                Me.PART_ListBox.ItemsSource = folder.Items
            End If
        End Sub

        Protected Overridable Sub Folder_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "ItemsSortPropertyName", "ItemsSortDirection", "ItemsGroupByPropertyName", "View"
                    Dim folder As Folder = CType(sender, Folder)

                    UIHelper.OnUIThread(
                        Sub()
                            If e.PropertyName = "ItemsGroupByPropertyName" Then
                                setGrouping(folder)
                            End If

                            If Not folder.IsRefreshingItems AndAlso Not Me.Folder Is Nothing AndAlso Not TypeOf folder Is SearchFolder Then
                                Dim folderViewState As FolderViewState = FolderViewState.FromViewName(folder.FullPath)
                                folderViewState.SortPropertyName = folder.ItemsSortPropertyName
                                folderViewState.SortDirection = folder.ItemsSortDirection
                                folderViewState.GroupByPropertyName = folder.ItemsGroupByPropertyName
                                folderViewState.View = folder.View
                                folderViewState.Persist()
                            End If
                        End Sub)
            End Select
        End Sub

        Private Sub setGrouping(folder As Folder)
            If Not Me.PART_ListBox.Resources("groupStyle") Is Nothing Then
                If Not String.IsNullOrWhiteSpace(folder.ItemsGroupByPropertyName) Then
                    If Me.PART_ListBox.GroupStyle.Count = 0 Then
                        Me.PART_ListBox.GroupStyle.Add(Me.PART_ListBox.Resources("groupStyle"))
                    End If
                Else
                    Me.PART_ListBox.GroupStyle.Clear()
                End If
            End If
        End Sub

        Shared Async Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = TryCast(d, BaseFolderView)

            'bfv.IsLoading = True

            ' stop recording time spent
            If Not bfv._timeSpentTimer Is Nothing Then bfv._timeSpentTimer.Dispose()

            ' hide listview so no-one sees us binding to the new folder and restoring the scroll position
            If Not bfv.PART_ListBox Is Nothing Then
                bfv.PART_ListBox.Visibility = Visibility.Hidden
            End If

            If Not e.OldValue Is Nothing Then
                Dim oldValue As Folder = e.OldValue

                ' stop listening for property changes
                RemoveHandler oldValue.PropertyChanged, AddressOf bfv.Folder_PropertyChanged

                ' record last scroll value for use with the back and forward navigation buttons
                oldValue.LastScrollOffset = bfv._lastScrollOffset
                oldValue.LastScrollSize = bfv._lastScrollSize
                bfv._scrollViewer.ScrollToHorizontalOffset(0)
                bfv._scrollViewer.ScrollToVerticalOffset(0)

                ' clear view binding
                bfv.ClearBinding()
            End If

            bfv.CheckBoxVisibility = If(Shell.Settings.DoShowCheckBoxesToSelect, Visibility.Visible, Visibility.Collapsed)

            If Not e.NewValue Is Nothing Then
                Dim newValue As Folder = e.NewValue

                ' track recent/frequent folders (in a task because for some folders this might take a while)
                Dim func As Func(Of Task) =
                    Async Function() As Task
                        FrequentFolders.Track(newValue)
                    End Function
                Task.Run(func)

                ' track time spent to find frequent folders
                bfv._timeSpentTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                FrequentFolders.RecordTimeSpent(newValue, 2)
                            End Sub)
                    End Sub), Nothing, 1000 * 60 * 2, 1000 * 60 * 2)

                ' set sorting and grouping
                If Not TypeOf newValue Is SearchFolder Then
                    Dim folderViewState As FolderViewState = FolderViewState.FromViewName(newValue.FullPath)
                    newValue.ItemsSortPropertyName = folderViewState.SortPropertyName
                    newValue.ItemsSortDirection = folderViewState.SortDirection
                    newValue.ItemsGroupByPropertyName = folderViewState.GroupByPropertyName
                End If

                ' get notified of folder property changes
                AddHandler newValue.PropertyChanged, AddressOf bfv.Folder_PropertyChanged

                ' bind view
                bfv.MakeBinding(e.NewValue)

                ' load items
                Await newValue.GetItemsAsync()

                ' async because otherwise we're a tad too early
                UIHelper.OnUIThreadAsync(
                    Async Sub()
                        Await Task.Delay(50)

                        bfv.OnBeforeRestoreScrollOffset()

                        ' restore folder scroll position
                        bfv._lastScrollOffset = newValue.LastScrollOffset
                        bfv._lastScrollSize = newValue.LastScrollSize
                        bfv._scrollViewer.ScrollToHorizontalOffset(If(bfv._lastScrollSize.Width = 0, 0, bfv._lastScrollOffset.X * bfv._scrollViewer.ScrollableWidth / bfv._lastScrollSize.Width))
                        bfv._scrollViewer.ScrollToVerticalOffset(If(bfv._lastScrollSize.Height = 0, 0, bfv._lastScrollOffset.Y * bfv._scrollViewer.ScrollableHeight / bfv._lastScrollSize.Height))

                        ' show listview
                        bfv.PART_ListBox.Visibility = Visibility.Visible
                    End Sub, Threading.DispatcherPriority.ContextIdle)
            End If

            'bfv.IsLoading = False
        End Sub

        Protected Overridable Sub OnBeforeRestoreScrollOffset()

        End Sub

        Shared Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim dlv As BaseFolderView = TryCast(d, BaseFolderView)

            If Not dlv._selectionHelper Is Nothing Then
                dlv._selectionHelper.SetSelectedItems(e.NewValue)
            End If
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    If Not _timeSpentTimer Is Nothing Then
                        _timeSpentTimer.Dispose()
                        _timeSpentTimer = Nothing
                    End If

                    If Not _typeToSearchTimer Is Nothing Then
                        _typeToSearchTimer.Dispose()
                        _typeToSearchTimer = Nothing
                    End If
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