Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class TreeView
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ItemsProperty As DependencyProperty = DependencyProperty.Register("Items", GetType(ObservableCollection(Of Item)), GetType(TreeView), New FrameworkPropertyMetadata(New ObservableCollection(Of Item)(), FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private PART_ListBox As ListBox
        Private _selectionHelper As SelectionHelper(Of Folder) = Nothing
        Private _isSettingSelectedFolder As Boolean
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _dropTarget As IDropTarget
        Private _menu As Laila.Shell.ContextMenu
        Private _rootIndex As Long = 0

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(TreeView), New FrameworkPropertyMetadata(GetType(TreeView)))
        End Sub

        Public Sub New()
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
                                AndAlso Await Me.IsDownSelection(e.Folder) Then
                                Me.SetSelectedFolder(e.Folder.LogicalParent)
                            End If
                    End Select
                End Sub

            AddHandler Shell.RequestSetSelectedFolder,
                Sub(s As Object, e As RequestSetSelectedFolderEventArgs)
                    Me.SetSelectedFolder(e.RequestedFolder, e.Callback)
                End Sub
        End Sub

        Private Async Function IsDownSelection(folder As Folder) As Task(Of Boolean)
            If Not Me.SelectedItem Is Nothing AndAlso Me.SelectedItem.Equals(folder) Then
                Return True
            Else
                For Each f In Await folder.GetItems()
                    If TypeOf f Is Folder AndAlso Await Me.IsDownSelection(f) Then
                        Return True
                    End If
                Next
                Return False
            End If
        End Function

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_ListBox = Template.FindName("PART_ListBox", Me)

            AddHandler PART_ListBox.Loaded,
                Async Sub(s As Object, e As EventArgs)
                    Dim items As List(Of Folder) = New List(Of Folder)()

                    ' home and galery
                    items.Add(Shell.SpecialFolders("Home"))
                    items.Add(Shell.SpecialFolders("Gallery"))

                    ' all the special folders under quick launch
                    Dim recentFolder As Folder = Folder.FromParsingName(Environment.GetFolderPath(Environment.SpecialFolder.Recent), Nothing)
                    For Each f In (Await Shell.SpecialFolders("Home") _
                        .GetItems()).Where(Function(i) TypeOf i Is Folder AndAlso Not IO.File.Exists(i.FullPath))
                        If Not items.ToList().Exists(Function(f2) f2.FullPath = f.FullPath AndAlso f2.LogicalParent Is Nothing) Then
                            Dim tvf As Folder = Folder.FromParsingName(f.FullPath, Nothing)
                            'tvf.IsPinned = True
                            items.Add(tvf)
                        End If
                    Next

                    ' 5 first regular folders under quicklaunch
                    For Each f In (Await recentFolder _
                        .GetItems()).Where(Function(i) TypeOf i Is Folder).Take(5)
                        If Not items.ToList().Exists(Function(f2) f2.FullPath = f.FullPath) Then
                            items.Add(Folder.FromParsingName(f.FullPath, Nothing))
                        End If
                    Next

                    ' all special folders under user profile that we'rent added yet
                    For Each f In (Await CType(Folder.FromParsingName(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), Nothing), Folder) _
                        .GetItems()).Where(Function(i) TypeOf i Is Folder AndAlso
                            Not items.ToList().Exists(Function(f2) f2.FullPath = i.FullPath))
                        Dim fpure As Folder = Folder.FromParsingName(f.FullPath, Nothing)
                        If Not items.ToList().Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                            (fpure.Parent Is Nothing OrElse fpure.Parent.Parent Is Nothing) Then
                            items.Insert(2, Folder.FromParsingName(f.FullPath, Nothing))
                        End If
                    Next

                    ' this computer & network
                    items.Add(Shell.SpecialFolders("This computer"))
                    items.Add(Shell.SpecialFolders("Network"))

                    For Each item In items
                        Me.Items.Add(item)
                    Next

                    Dim view As ICollectionView = CollectionViewSource.GetDefaultView(Me.Items)
                    view.Refresh()

                    _selectionHelper = New SelectionHelper(Of Folder)(PART_ListBox)
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

                    If Me.Folder Is Nothing Then Shell.SetSelectedFolder(Shell.SpecialFolders("This computer"), Nothing)
                End Sub
        End Sub

        Protected Overridable Sub OnSelectionChanged()
            If Not Me.SelectedItem Is Nothing Then
                If Not Me.SelectedItem.Equals(Me.Folder) Then
                    Me.Folder = Me.SelectedItem
                End If
            End If
        End Sub

        Public ReadOnly Property SelectedItem As Folder
            Get
                If _selectionHelper.SelectedItems.Count = 1 Then
                    Return _selectionHelper.SelectedItems(0)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public Sub SetSelectedFolder(folder As Folder, Optional callback As Action(Of Folder) = Nothing)
            If Not _isSettingSelectedFolder Then
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

                        Dim func As Action(Of Folder, Action(Of Boolean)) =
                            Sub(item As Folder, callback2 As Action(Of Boolean))
                                Dim tf2 As Folder
                                UIHelper.OnUIThreadAsync(
                                    Async Sub()
                                        tf2 = (Await tf.GetItems()).FirstOrDefault(Function(f2) f2.FullPath = item.FullPath)
                                        If Not tf2 Is Nothing Then
                                            Debug.WriteLine("SetSelectedFolder found " & tf2.FullPath)
                                            tf.IsExpanded = True
                                            tf = tf2
                                            callback2(False)
                                        Else
                                            Debug.WriteLine("SetSelectedFolder didn't find " & item.FullPath)
                                            callback2(True)
                                        End If
                                    End Sub)
                            End Sub
                        Dim en As IEnumerator(Of Folder) = list.GetEnumerator()
                        Dim cb As System.Action(Of Boolean) =
                            Sub(cancel As Boolean)
                                If Not cancel Then
                                    If en.MoveNext() Then
                                        func(en.Current, cb)
                                    Else
                                        ' wait for expanding to complete
                                        For i = 1 To 5
                                            Application.Current.Dispatcher.Invoke(
                                            Sub()
                                            End Sub, Threading.DispatcherPriority.ContextIdle)
                                            Application.Current.Dispatcher.Invoke(
                                            Sub()
                                            End Sub, Threading.DispatcherPriority.ContextIdle)
                                            Application.Current.Dispatcher.Invoke(
                                            Sub()
                                            End Sub, Threading.DispatcherPriority.ContextIdle)
                                        Next

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
                            End Sub
                        If en.MoveNext() Then
                            func(en.Current, cb)
                        Else
                            If Not tf Is Nothing Then cb(False)
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
        End Sub

        Private Sub OnTreeViewPreviewMouseMove(sender As Object, e As MouseEventArgs)
            If Not _mouseItemDown Is Nothing AndAlso Not Me.SelectedItem Is Nothing AndAlso
                (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                Dim currentPointDown As Point = e.GetPosition(PART_ListBox)
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                    Drag.Start({Me.SelectedItem}, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                End If
            End If
        End Sub

        Private Sub OnTreeViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(PART_ListBox)

            If Not e.OriginalSource Is Nothing Then
                Dim treeViewItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
                Dim clickedItem As Folder = treeViewItem?.DataContext
                _mouseItemDown = clickedItem
                If e.RightButton = MouseButtonState.Pressed Then
                    If Not clickedItem Is Nothing Then
                        If Me.SelectedItem Is Nothing Then Me.SetSelectedFolder(clickedItem)

                        Dim parent As Folder = clickedItem.Parent
                        If parent Is Nothing Then parent = Shell.Desktop

                        _menu = New Laila.Shell.ContextMenu()
                        AddHandler _menu.Click,
                        Sub(id As Integer, verb As String, ByRef isHandled As Boolean)
                            Select Case verb
                                Case "open"
                                    Me.SetSelectedFolder(clickedItem)
                                    isHandled = True
                            End Select
                        End Sub

                        Dim contextMenu As Controls.ContextMenu = _menu.GetContextMenu(parent, {clickedItem}, False)
                        PART_ListBox.ContextMenu = contextMenu
                        e.Handled = True
                    Else
                        PART_ListBox.ContextMenu = Nothing
                    End If
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
                            If folder.LogicalParent Is Nothing Then
                                folder._rootIndex = _rootIndex
                                _rootIndex += 1
                            End If
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
                                        If Not Me.Items.Contains(item2) Then
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
                    Dim view As ICollectionView = CollectionViewSource.GetDefaultView(Me.Items)
                    view.Refresh()
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

    End Class
End Namespace