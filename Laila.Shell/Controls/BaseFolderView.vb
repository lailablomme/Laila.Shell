Imports System.ComponentModel
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography.X509Certificates
Imports System.Threading
Imports System.Windows
Imports System.Windows.Annotations
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Shell
Imports System.Xml
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.SevenZip

Namespace Controls
    Public MustInherit Class BaseFolderView
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ColumnsInProperty As DependencyProperty = DependencyProperty.Register("ColumnsIn", GetType(Behaviors.GridViewExtBehavior.ColumnsInData), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly MenusProperty As DependencyProperty = DependencyProperty.Register("Menus", GetType(Menus), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Friend PART_ListView As System.Windows.Controls.ListView
        Private PART_Grid As Grid
        Private _columnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
        Private _isLoading As Boolean
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _mouseItemOver As Item
        Private _menus As Laila.Shell.Controls.Menus
        Private _timeSpentTimer As Timer
        Private _scrollViewer As ScrollViewer
        Private _lastScrollOffset As Point
        Private _lastScrollSize As Size
        Private _isLoaded As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(BaseFolderView), New FrameworkPropertyMetadata(GetType(BaseFolderView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_ListView = Template.FindName("PART_ListView", Me)
            Me.PART_Grid = Template.FindName("PART_Grid", Me)

            Me.PART_ListView.Visibility = Visibility.Hidden

            If Not Me.Folder Is Nothing Then
                Me.MakeBinding(Me.Folder)
            End If

            AddHandler PART_ListView.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        _selectionHelper = New SelectionHelper(Of Item)(Me.PART_ListView)
                        _selectionHelper.SelectionChanged =
                            Sub()
                                If Not Me.Folder Is Nothing Then _
                                    Me.SelectedItems = _selectionHelper.SelectedItems
                            End Sub
                        _selectionHelper.SetSelectedItems(Me.SelectedItems)

                        _scrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(Me.PART_ListView)(0)
                        AddHandler _scrollViewer.ScrollChanged,
                            Sub(s2 As Object, e2 As ScrollChangedEventArgs)
                                If Not Me.Folder Is Nothing Then
                                    _lastScrollOffset = New Point(_scrollViewer.HorizontalOffset, _scrollViewer.VerticalOffset)
                                    _lastScrollSize = New Size(_scrollViewer.ScrollableWidth, _scrollViewer.ScrollableHeight)
                                End If
                                UIHelper.OnUIThreadAsync(
                                Sub()
                                    GC.Collect()
                                End Sub)
                            End Sub
                    End If
                End Sub

            AddHandler Me.PART_ListView.PreviewMouseMove, AddressOf OnListViewPreviewMouseMove
            AddHandler Me.PART_ListView.PreviewMouseDown, AddressOf OnListViewPreviewMouseButtonDown
            AddHandler Me.PART_ListView.PreviewMouseUp, AddressOf OnListViewPreviewMouseButtonUp
            AddHandler Me.PART_ListView.MouseLeave, AddressOf OnListViewMouseLeave
            AddHandler Me.PreviewKeyDown, AddressOf OnListViewKeyDown
        End Sub

        Private Sub OnListViewKeyDown(sender As Object, e As KeyEventArgs)
            If e.Key = Key.C AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) AndAlso Me.SelectedItems.Count > 0 Then
                Clipboard.CopyFiles(Me.SelectedItems)
            ElseIf e.Key = Key.X AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) AndAlso Me.SelectedItems.Count > 0 Then
                Clipboard.CutFiles(Me.SelectedItems)
            End If
        End Sub

        Private Sub OnListViewPreviewMouseMove(sender As Object, e As MouseEventArgs)
            Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
            Dim overItem As Item = TryCast(listViewItem?.DataContext, Item)
            If Not overItem Is Nothing AndAlso Not overItem.Equals(_mouseItemOver) Then
                Dim toolTip As String = overItem.InfoTip
                listViewItem.ToolTip = If(String.IsNullOrWhiteSpace(toolTip), Nothing, toolTip)
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
                Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
                Dim clickedItem As Item = TryCast(listViewItem?.DataContext, Item)
                _mouseItemDown = clickedItem
                If clickedItem Is Nothing Then
                    Me.PART_ListView.Focus()
                Else
                    listViewItem.Focus()
                End If
                If e.LeftButton = MouseButtonState.Pressed AndAlso e.ClickCount = 2 AndAlso Not clickedItem Is Nothing Then
                    Using Shell.OverrideCursor(Cursors.Wait)
                        Me.SelectedItems = {clickedItem}
                        If TypeOf clickedItem Is Folder Then
                            Me.Folder = clickedItem
                            UIHelper.OnUIThread(
                                Sub()
                                End Sub, Threading.DispatcherPriority.Render)
                        Else
                            hookMenus()
                            If Not _menus Is Nothing Then
                                Dim contextMenu As ContextMenu = _menus.GetDefaultContextMenu()
                                _menus.InvokeCommand(contextMenu, _menus.DefaultId)
                            End If
                        End If
                    End Using
                ElseIf e.LeftButton = MouseButtonState.Pressed AndAlso Not clickedItem Is Nothing Then
                    If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 _
                            AndAlso Me.SelectedItems.Contains(clickedItem) _
                            AndAlso Keyboard.Modifiers = ModifierKeys.None Then
                        e.Handled = True
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

                    hookMenus()
                    If Not _menus Is Nothing Then
                        _menus.Update()
                        Me.PART_ListView.ContextMenu = _menus.ItemContextMenu
                    End If
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

        Private Sub hookMenus()
            If Not EqualityComparer(Of Menus).Default.Equals(_menus, Me.Menus) Then
                _menus = Me.Menus
                AddHandler _menus.CommandInvoked,
                    Sub(s As Object, e2 As CommandInvokedEventArgs)
                        Select Case e2.Verb
                            Case "open"
                                If Not Me.SelectedItems Is Nothing _
                                        AndAlso Me.SelectedItems.Count = 1 _
                                        AndAlso TypeOf Me.SelectedItems(0) Is Folder Then
                                    Me.Folder = Me.SelectedItems(0)
                                    e2.IsHandled = True
                                End If
                            Case "rename"
                                Me.DoRename()
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
            End If
        End Sub

        Protected MustOverride Sub GetItemNameCoordinates(listViewItem As ListViewItem, ByRef textAlignment As TextAlignment,
                                                          ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)

        Public Sub DoRename()
            Dim listViewItem As ListViewItem = Me.PART_ListView.ItemContainerGenerator.ContainerFromItem(Me.SelectedItems(0))
            DoRename(listViewItem)
        End Sub

        Public Sub DoRename(listViewItem As ListViewItem)
            Dim point As Point, size As Size, textAlignment As TextAlignment, fontSize As Double
            Me.GetItemNameCoordinates(listViewItem, textAlignment, point, size, fontSize)
            Menus.DoRename(point, size, textAlignment, fontSize, listViewItem.DataContext, Me.PART_Grid)
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

        Public Property Menus As Menus
            Get
                Return GetValue(MenusProperty)
            End Get
            Set(ByVal value As Menus)
                SetCurrentValue(MenusProperty, value)
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

        Protected Overridable Sub ClearBinding()
            If Not Me.PART_ListView Is Nothing Then
                Me.PART_ListView.ItemsSource = Nothing
            End If
        End Sub

        Protected Overridable Sub MakeBinding(folder As Folder)
            If Not Me.PART_ListView Is Nothing Then
                If Not String.IsNullOrWhiteSpace(folder.ItemsGroupByPropertyName) Then
                    Me.PART_ListView.GroupStyle.Add(Me.PART_ListView.Resources("groupStyle"))
                Else
                    Me.PART_ListView.GroupStyle.Clear()
                End If

                Me.PART_ListView.ItemsSource = folder.Items
            End If
        End Sub

        Protected Overridable Sub Folder_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "IsRefreshingItems"
                    UIHelper.OnUIThread(
                        Sub()
                            Me.IsLoading = CType(s, Folder).IsRefreshingItems
                        End Sub)
                Case "ItemsSortPropertyName", "ItemsSortDirection", "ItemsGroupByPropertyName", "View"
                    Dim folder As Folder = CType(s, Folder)
                    Dim folderViewState As FolderViewState = FolderViewState.FromViewName(folder.FullPath)
                    folderViewState.SortPropertyName = folder.ItemsSortPropertyName
                    folderViewState.SortDirection = folder.ItemsSortDirection
                    folderViewState.GroupByPropertyName = folder.ItemsGroupByPropertyName
                    folderViewState.View = folder.View
                    folderViewState.Persist()

                    If Not String.IsNullOrWhiteSpace(folder.ItemsGroupByPropertyName) Then
                        Me.PART_ListView.GroupStyle.Add(Me.PART_ListView.Resources("groupStyle"))
                    Else
                        Me.PART_ListView.GroupStyle.Clear()
                    End If
            End Select
        End Sub

        Shared Async Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = TryCast(d, BaseFolderView)

            bfv.IsLoading = True

            ' stop recording time spent
            If Not bfv._timeSpentTimer Is Nothing Then bfv._timeSpentTimer.Dispose()

            ' hide listview so no-one sees us binding to the new folder and restoring the scroll position
            If Not bfv.PART_ListView Is Nothing Then
                bfv.PART_ListView.Visibility = Visibility.Hidden
            End If

            If Not e.OldValue Is Nothing Then
                Dim oldValue As Folder = e.OldValue

                ' stop listening for property changes
                RemoveHandler oldValue.PropertyChanged, AddressOf bfv.Folder_PropertyChanged

                ' record last scroll value for use with the back and forward navigation buttons
                oldValue.LastScrollOffset = bfv._lastScrollOffset
                oldValue.LastScrollSize = bfv._lastScrollSize

                ' clear view binding
                bfv.ClearBinding()
            End If

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

                ' load items
                Await newValue.GetItemsAsync()

                ' get notified of folder property changes
                AddHandler newValue.PropertyChanged, AddressOf bfv.Folder_PropertyChanged

                ' bind view
                bfv.MakeBinding(e.NewValue)

                ' async because otherwise we're a tad too early
                UIHelper.OnUIThreadAsync(
                    Async Sub()
                        Await Task.Delay(50)

                        ' restore folder scroll position
                        bfv._lastScrollOffset = newValue.LastScrollOffset
                        bfv._lastScrollSize = newValue.LastScrollSize
                        bfv._scrollViewer.ScrollToHorizontalOffset(If(bfv._lastScrollSize.Width = 0, 0, bfv._lastScrollOffset.X * bfv._scrollViewer.ScrollableWidth / bfv._lastScrollSize.Width))
                        bfv._scrollViewer.ScrollToVerticalOffset(If(bfv._lastScrollSize.Height = 0, 0, bfv._lastScrollOffset.Y * bfv._scrollViewer.ScrollableHeight / bfv._lastScrollSize.Height))

                        ' show listview
                        bfv.PART_ListView.Visibility = Visibility.Visible
                    End Sub)
            End If

            bfv.IsLoading = False
        End Sub

        Shared Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim dlv As BaseFolderView = TryCast(d, BaseFolderView)

            If Not dlv._selectionHelper Is Nothing Then
                dlv._selectionHelper.SetSelectedItems(e.NewValue)
            End If
        End Sub
    End Class
End Namespace