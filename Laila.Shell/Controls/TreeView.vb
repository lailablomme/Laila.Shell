Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.PinnedItems

Namespace Controls
    Public Class TreeView
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ItemsProperty As DependencyProperty = DependencyProperty.Register("Items", GetType(ObservableCollection(Of Item)), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private PART_Grid As Grid
        Private PART_ListBox As ListBox
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _isSettingSelectedFolder As Boolean
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _dropTarget As IDropTarget
        Private _menu As Laila.Shell.ContextMenu
        Private _fequentUpdateTimer As Timer

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(TreeView), New FrameworkPropertyMetadata(GetType(TreeView)))
        End Sub

        Public Sub New()
            Me.Items = New ObservableCollection(Of Item)()

            Dim view As ICollectionView = CollectionViewSource.GetDefaultView(Me.Items)
            view.Filter = AddressOf filter
            view.SortDescriptions.Add(New SortDescription() With {
                .PropertyName = "TreeSortKey",
                .Direction = ListSortDirection.Ascending
            })
            AddHandler Me.Items.CollectionChanged, AddressOf items_CollectionChanged

            AddHandler System.Windows.Application.Current.MainWindow.Closed,
                Sub()
                    WpfDragTargetProxy.RevokeDragDrop(PART_ListBox)
                End Sub

            AddHandler Shell.FolderNotification,
                 Async Sub(s As Object, e As FolderNotificationEventArgs)
                     Select Case e.Event
                         Case SHCNE.RMDIR, SHCNE.DELETE, SHCNE.DRIVEREMOVED
                             If Not Me.SelectedItem Is Nothing AndAlso Not e.Folder.LogicalParent Is Nothing _
                                AndAlso Me.GetIsSelectionDownFolder(e.Folder) Then
                                 Await Me.SetSelectedFolder(e.Folder.LogicalParent)
                             End If
                     End Select
                 End Sub

            ' home and galery
            Shell.SpecialFolders("Home").TreeRootIndex = TreeRootSection.SYSTEM + 0 : Items.Add(Shell.SpecialFolders("Home"))
            Shell.SpecialFolders("Gallery").TreeRootIndex = TreeRootSection.SYSTEM + 1 : Items.Add(Shell.SpecialFolders("Gallery"))

            ' separators
            Items.Add(New SeparatorFolder() With {.TreeRootIndex = TreeRootSection.PINNED - 1})
            Items.Add(New SeparatorFolder() With {.TreeRootIndex = TreeRootSection.ENVIRONMENT - 1})

            ' this computer & network
            Shell.SpecialFolders("This computer").TreeRootIndex = TreeRootSection.ENVIRONMENT + 0 : Items.Add(Shell.SpecialFolders("This computer"))
            Shell.SpecialFolders("Network").TreeRootIndex = TreeRootSection.ENVIRONMENT + 1 : Items.Add(Shell.SpecialFolders("Network"))

            updatePinnedItems()
            updateFrequentFolders()

            ' frequent folders
            _fequentUpdateTimer = New Timer(New TimerCallback(
                        Sub()
                            Application.Current.Dispatcher.Invoke(
                                Async Function() As Task
                                    Await updateFrequentFolders()
                                    CollectionViewSource.GetDefaultView(Me.Items).Refresh()
                                End Function)
                        End Sub), Nothing, 1000 * 60, 1000 * 60)

            AddHandler PinnedItems.ItemPinned,
                        Async Sub(s2 As Object, e2 As PinnedItemEventArgs)
                            Await updatePinnedItems()
                            Await updateFrequentFolders()
                            CollectionViewSource.GetDefaultView(Me.Items).Refresh()
                        End Sub
            AddHandler PinnedItems.ItemUnpinned,
                        Async Sub(s2 As Object, e2 As PinnedItemEventArgs)
                            Await updatePinnedItems()
                            Await updateFrequentFolders()
                            CollectionViewSource.GetDefaultView(Me.Items).Refresh()
                        End Sub

            CollectionViewSource.GetDefaultView(Me.Items).Refresh()
        End Sub

        Private Function GetIsSelectionDownFolder(folder As Folder) As Boolean
            If Not Me.SelectedItem Is Nothing AndAlso Me.SelectedItem.Equals(folder) Then
                Return True
            Else
                For Each f In folder.Items
                    If TypeOf f Is Folder AndAlso Me.GetIsSelectionDownFolder(f) Then
                        Return True
                    End If
                Next
                Return False
            End If
        End Function

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_ListBox = Template.FindName("PART_ListBox", Me)
            PART_Grid = Template.FindName("PART_Grid", Me)

            _selectionHelper = New SelectionHelper(Of Item)(PART_ListBox)
            _selectionHelper.SelectionChanged =
                        Sub()
                            UIHelper.OnUIThread(
                                Sub()
                                    Me.OnSelectionChanged()
                                End Sub)
                        End Sub


            AddHandler PART_ListBox.PreviewMouseMove, AddressOf OnTreeViewPreviewMouseMove
            AddHandler PART_ListBox.PreviewMouseLeftButtonDown, AddressOf OnTreeViewPreviewMouseButtonDown
            AddHandler PART_ListBox.PreviewMouseRightButtonDown, AddressOf OnTreeViewPreviewMouseButtonDown
            AddHandler PART_ListBox.PreviewMouseUp, AddressOf OnTreeViewPreviewMouseButtonUp
            AddHandler PART_ListBox.MouseLeave, AddressOf OnTreeViewMouseLeave

            _dropTarget = New TreeViewDropTarget(Me)
            WpfDragTargetProxy.RegisterDragDrop(PART_ListBox, _dropTarget)
        End Sub

        Private Async Function updatePinnedItems() As Task
            ' get pinned items
            Dim pinnedItemsList As IEnumerable(Of Item) = PinnedItems.GetPinnedItems()
            Dim count As Integer = 0
            For Each pinnedItem In pinnedItemsList
                Dim existingPinnedItem As Item = Me.Items.FirstOrDefault(Function(i) _
                    i.FullPath = pinnedItem.FullPath AndAlso i.LogicalParent Is Nothing _
                        AndAlso i.TreeRootIndex >= TreeRootSection.PINNED AndAlso i.TreeRootIndex < TreeRootSection.FREQUENT)
                If existingPinnedItem Is Nothing Then
                    ' insert new frequent folder
                    pinnedItem.TreeRootIndex = TreeRootSection.PINNED + count
                    pinnedItem.IsPinned = True
                    Me.Items.Add(pinnedItem)
                Else
                    ' update exising frequent folder
                    pinnedItem.Dispose()
                    existingPinnedItem.TreeRootIndex = TreeRootSection.PINNED + count
                End If
                count += 1
            Next
            ' remove pinned items no longer in the list
            For Each pinnedItem As Item In Me.Items.Where(Function(i) _
                Not TypeOf i Is SeparatorFolder AndAlso i.LogicalParent Is Nothing _
                    AndAlso i.TreeRootIndex >= TreeRootSection.PINNED AndAlso i.TreeRootIndex < TreeRootSection.FREQUENT _
                        AndAlso Not pinnedItemsList.ToList().Exists(Function(f) f.FullPath = i.FullPath)).ToList()
                If (TypeOf pinnedItem Is Folder AndAlso GetIsSelectionDownFolder(pinnedItem)) OrElse pinnedItem.Equals(Me.SelectedItem) Then
                    _selectionHelper.SetSelectedItems({})
                End If
                Me.Items.Remove(pinnedItem)
            Next
        End Function

        Private Async Function updateFrequentFolders() As Task
            ' get frequent folders
            Dim frequentFoldersList As IEnumerable(Of Folder) = FrequentFolders.GetMostFrequent()
            Dim count As Integer = 0
            For Each frequentFolder In frequentFoldersList
                Dim existingFrequentFolder As Folder = Me.Items.FirstOrDefault(Function(i) _
                    i.FullPath = frequentFolder.FullPath AndAlso i.LogicalParent Is Nothing _
                        AndAlso i.TreeRootIndex >= TreeRootSection.FREQUENT AndAlso i.TreeRootIndex < TreeRootSection.ENVIRONMENT)
                If existingFrequentFolder Is Nothing Then
                    ' insert new frequent folder
                    frequentFolder.TreeRootIndex = TreeRootSection.FREQUENT + count
                    Me.Items.Add(frequentFolder)
                Else
                    ' update exising frequent folder
                    frequentFolder.Dispose()
                    existingFrequentFolder.TreeRootIndex = TreeRootSection.FREQUENT + count
                End If
                count += 1
            Next
            ' remove frequent folders no longer in the list
            For Each frequentFolder As Folder In Me.Items.Where(Function(i) _
                Not TypeOf i Is SeparatorFolder AndAlso i.LogicalParent Is Nothing _
                    AndAlso i.TreeRootIndex >= TreeRootSection.FREQUENT AndAlso i.TreeRootIndex < TreeRootSection.ENVIRONMENT _
                        AndAlso Not frequentFoldersList.ToList().Exists(Function(f) f.FullPath = i.FullPath)).ToList()
                If GetIsSelectionDownFolder(frequentFolder) Then
                    _selectionHelper.SetSelectedItems({})
                End If
                Me.Items.Remove(frequentFolder)
            Next
        End Function

        Protected Overridable Sub OnSelectionChanged()
            If Not Me.SelectedItem Is Nothing Then
                If TypeOf Me.SelectedItem Is Folder AndAlso Not Me.SelectedItem.Equals(Me.Folder) Then
                    Me.Folder = Me.SelectedItem
                End If
            End If
        End Sub

        Public ReadOnly Property SelectedItem As Item
            Get
                If _selectionHelper.SelectedItems.Count = 1 Then
                    Return _selectionHelper.SelectedItems(0)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public Async Function SetSelectedFolder(folder As Folder, Optional callback As Action(Of Folder) = Nothing) As Task
            If Not _isSettingSelectedFolder Then
                Await Task.Delay(100)

                _isSettingSelectedFolder = True

                Debug.WriteLine("SetSelectedFolder " & folder?.FullPath)
                Dim list As List(Of Folder) = New List(Of Folder)()
                Dim currentFolder As Folder = folder
                If Not currentFolder Is Nothing Then
                    While Not currentFolder.LogicalParent Is Nothing
                        list.Add(currentFolder)
                        Debug.WriteLine("SetSelectedFolder Added parent " & currentFolder.FullPath)
                        currentFolder = currentFolder.LogicalParent
                    End While

                    Dim tf As Folder
                    If Me.Items.ToList().Exists(Function(f1) f1.FullPath = currentFolder.FullPath AndAlso f1.LogicalParent Is Nothing) Then
                        tf = Me.Items.First(Function(f1) f1.FullPath = currentFolder.FullPath AndAlso f1.LogicalParent Is Nothing)
                    Else
                        tf = Nothing
                    End If

                    If Not tf Is Nothing Then
                        list.Reverse()

                        Dim func As Func(Of Folder, Func(Of Boolean, Task), Task) =
                            Async Function(item As Folder, callback2 As Func(Of Boolean, Task)) As Task
                                Dim tf2 = (Await tf.GetItemsAsync()).FirstOrDefault(Function(f2) f2.FullPath = item.FullPath)
                                If Not tf2 Is Nothing Then
                                    Debug.WriteLine("SetSelectedFolder found " & tf2.FullPath)
                                    tf.IsExpanded = True
                                    tf = tf2
                                    Await callback2(False)
                                Else
                                    Debug.WriteLine("SetSelectedFolder didn't find " & item.FullPath)
                                    Await callback2(True)
                                End If
                            End Function
                        Dim en As IEnumerator(Of Folder) = list.GetEnumerator()
                        Dim cb As System.Func(Of Boolean, Task) =
                            Async Function(cancel As Boolean) As Task
                                If Not cancel Then
                                    If en.MoveNext() Then
                                        Await func(en.Current, cb)
                                    Else
                                        _selectionHelper.SetSelectedItems({tf})

                                        If Not callback Is Nothing Then
                                            callback(tf)
                                        End If

                                        _isSettingSelectedFolder = False
                                    End If
                                Else
                                    _selectionHelper.SetSelectedItems({})

                                    If Not callback Is Nothing Then
                                        callback(Nothing)
                                    End If

                                    _isSettingSelectedFolder = False
                                End If
                            End Function
                        If en.MoveNext() Then
                            Await func(en.Current, cb)
                        Else
                            If Not tf Is Nothing Then Await cb(False)
                            _isSettingSelectedFolder = False
                        End If
                    Else
                        Me.Folder = folder

                        If Not callback Is Nothing Then
                            callback(folder)
                        End If

                        _isSettingSelectedFolder = False
                    End If
                Else
                    _isSettingSelectedFolder = False
                End If
            End If
        End Function

        Private Sub OnTreeViewPreviewMouseMove(sender As Object, e As MouseEventArgs)
            If Not _mouseItemDown Is Nothing AndAlso Not Me.SelectedItem Is Nothing AndAlso
                (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                Dim currentPointDown As Point = e.GetPosition(PART_ListBox)
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                    Drag.Start({_mouseItemDown}, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                End If
            End If
        End Sub

        Private Async Sub OnTreeViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(PART_ListBox)

            If Not e.OriginalSource Is Nothing Then
                Dim treeViewItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
                Dim clickedItem As Item = treeViewItem?.DataContext
                If Not TypeOf clickedItem Is SeparatorFolder Then
                    _mouseItemDown = clickedItem
                    If e.RightButton = MouseButtonState.Pressed Then
                        If Not clickedItem Is Nothing Then
                            If Me.SelectedItem Is Nothing Then Await Me.SetSelectedFolder(clickedItem)

                            Dim parent As Folder = clickedItem.Parent
                            If parent Is Nothing Then parent = Shell.Desktop

                            _menu = New Laila.Shell.ContextMenu()
                            AddHandler _menu.CommandInvoked,
                                Sub(s As Object, e2 As CommandInvokedEventArgs)
                                    Select Case e2.Verb
                                        Case "open"
                                            If TypeOf clickedItem Is Folder Then
                                                Me.SetSelectedFolder(clickedItem)
                                                e2.IsHandled = True
                                            End If
                                        Case "rename"
                                            Dim pt As Point = Me.PointFromScreen(treeViewItem.PointToScreen(New Point(0, 0)))
                                            pt.X += clickedItem.TreeMargin.Left + 37
                                            pt.Y -= 1
                                            _menu.DoRename(pt, Me.ActualWidth - pt.X - 2, clickedItem, Me.PART_Grid)
                                            e2.IsHandled = True
                                        Case "laila.shell.(un)pin"
                                            If e2.IsChecked Then
                                                PinnedItems.PinItem(clickedItem)
                                            Else
                                                PinnedItems.UnpinItem(clickedItem)
                                            End If
                                            e2.IsHandled = True
                                    End Select
                                End Sub

                            Dim contextMenu As Controls.ContextMenu = _menu.GetContextMenu(parent, {clickedItem}, False)
                            PART_ListBox.ContextMenu = contextMenu
                            e.Handled = True
                        Else
                            PART_ListBox.ContextMenu = Nothing
                        End If
                    ElseIf e.LeftButton = MouseButtonState.Pressed AndAlso e.ClickCount = 2 Then
                        If Not clickedItem Is Nothing AndAlso TypeOf clickedItem Is Folder Then
                            CType(clickedItem, Folder).IsExpanded = Not CType(clickedItem, Folder).IsExpanded
                        End If
                    End If
                Else
                    _mouseItemDown = Nothing
                    e.Handled = True
                End If
            Else
                _mouseItemDown = Nothing
            End If
        End Sub

        Public Sub OnTreeViewPreviewMouseButtonUp(sender As Object, e As MouseButtonEventArgs)
            _mouseItemDown = Nothing
        End Sub

        Public Sub OnTreeViewMouseLeave(sender As Object, e As MouseEventArgs)
            _mouseItemDown = Nothing
        End Sub

        Private Sub items_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add
                    For Each item In e.NewItems
                        If TypeOf item Is Folder Then
                            Dim folder As Folder = item
                            AddHandler folder.PropertyChanged, AddressOf folder_PropertyChanged
                            AddHandler folder._items.CollectionChanged, AddressOf folder_CollectionChanged
                            For Each item2 In folder._items.Where(Function(i) TypeOf i Is Folder)
                                UIHelper.OnUIThreadAsync(
                                    Sub()
                                        If Not Me.Items.Contains(item2) Then
                                            Me.Items.Add(item2)
                                        End If
                                    End Sub)
                            Next
                        End If
                    Next
                Case NotifyCollectionChangedAction.Remove
                    For Each item In e.OldItems
                        If TypeOf item Is Folder Then
                            Dim folder As Folder = item
                            RemoveHandler folder.PropertyChanged, AddressOf folder_PropertyChanged
                            RemoveHandler folder._items.CollectionChanged, AddressOf folder_CollectionChanged
                            For Each item2 In Me.Items.Where(Function(i) TypeOf i Is Folder _
                                AndAlso Not i.LogicalParent Is Nothing AndAlso i.LogicalParent.Equals(folder))
                                UIHelper.OnUIThreadAsync(
                                    Sub()
                                        If Me.Items.Contains(item2) Then
                                            Me.Items.Remove(item2)
                                        End If
                                    End Sub)
                            Next
                        End If
                    Next
                Case NotifyCollectionChangedAction.Reset
                    Throw New NotSupportedException()
            End Select
        End Sub

        Private Sub folder_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add
                    For Each item In e.NewItems
                        If TypeOf item Is Folder AndAlso Not TypeOf item Is DummyFolder Then
                            Me.Items.Add(item)
                        End If
                    Next
                Case NotifyCollectionChangedAction.Remove
                    For Each item In e.OldItems
                        If TypeOf item Is Folder Then
                            Me.Items.Remove(item)
                        End If
                    Next
                Case NotifyCollectionChangedAction.Replace
                    For Each item In e.NewItems
                        If TypeOf item Is Folder AndAlso Not TypeOf item Is DummyFolder Then
                            Me.Items.Add(item)
                        End If
                    Next
                    For Each item In e.OldItems
                        If TypeOf item Is Folder Then
                            Me.Items.Remove(item)
                        End If
                    Next
                Case NotifyCollectionChangedAction.Reset
                    Throw New NotSupportedException()
            End Select
        End Sub

        Private Sub folder_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
            Dim folder As Folder = s

            Select Case e.PropertyName
                Case "IsExpanded"
                    CollectionViewSource.GetDefaultView(Me.Items).Refresh()
                Case "TreeSortKey"
                    For Each item2 In Me.Items.Where(Function(i) TypeOf i Is Folder _
                        AndAlso Not i.LogicalParent Is Nothing AndAlso i.LogicalParent.Equals(folder))
                        item2.NotifyOfPropertyChange("TreeSortKey")
                    Next
            End Select
        End Sub

        Private Function filter(i As Object) As Boolean
            Dim item As Item = i
            Return item.LogicalParent Is Nothing OrElse item.LogicalParent.IsExpanded
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
            Dim tv As TreeView = TryCast(d, TreeView)
            tv.SetSelectedFolder(e.NewValue)
        End Sub

        Public Property Items As ObservableCollection(Of Item)
            Get
                Return GetValue(ItemsProperty)
            End Get
            Set(ByVal value As ObservableCollection(Of Item))
                SetCurrentValue(ItemsProperty, value)
            End Set
        End Property

        Public Enum TreeRootSection As Long
            SYSTEM = 0
            PINNED = 100
            FREQUENT = Long.MaxValue - 100
            ENVIRONMENT = Long.MaxValue - 5
        End Enum
    End Class
End Namespace