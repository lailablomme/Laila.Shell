Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Runtime
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Controls
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Namespace ViewModels
    Public Class TreeViewModel
        Inherits NotifyPropertyChangedBase

        Friend _view As Controls.TreeView
        Private _folders1 As ObservableCollection(Of Folder) = New ObservableCollection(Of Folder)
        Private _folders2 As ObservableCollection(Of Folder) = New ObservableCollection(Of Folder)
        Private _folders3 As ObservableCollection(Of Folder) = New ObservableCollection(Of Folder)
        Private _selectionHelper1 As SelectionHelper(Of Folder) = Nothing
        Private _selectionHelper2 As SelectionHelper(Of Folder) = Nothing
        Private _selectionHelper3 As SelectionHelper(Of Folder) = Nothing
        Private _isSettingSelectedFolder As Boolean
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _dropTarget As IDropTarget
        Private _menu As ContextMenu = New ContextMenu()

        Public Sub New(view As Controls.TreeView)
            _view = view

            Dim isWorking As Boolean
            AddHandler _view.Loaded,
                Async Sub(s As Object, e As EventArgs)
                    ' home and galery
                    _folders1.Add(Folder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing))
                    _folders1.Add(Folder.FromParsingName("shell:::{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}", Nothing))

                    ' all the special folders under quick launch
                    Dim recentFolder As Folder = Folder.FromParsingName(Environment.GetFolderPath(Environment.SpecialFolder.Recent), Nothing)
                    For Each f In (Await CType(Folder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing), Folder) _
                        .GetItems()).Where(Function(i) TypeOf i Is Folder AndAlso Not IO.File.Exists(i.FullPath))
                        If Not _folders1.ToList().Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                            Not _folders2.ToList().Exists(Function(f2) f2.FullPath = f.FullPath) Then
                            Dim tvf As Folder = Folder.FromParsingName(f.FullPath, Nothing)
                            'tvf.IsPinned = True
                            _folders2.Add(tvf)
                        End If
                    Next

                    ' 5 first regular folders under quicklaunch
                    For Each f In (Await recentFolder _
                        .GetItems()).Where(Function(i) TypeOf i Is Folder).Take(5)
                        If Not _folders1.ToList().Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                            Not _folders2.ToList().Exists(Function(f2) f2.FullPath = f.FullPath) Then
                            _folders2.Add(Folder.FromParsingName(f.FullPath, Nothing))
                        End If
                    Next

                    ' all special folders under user profile that we'rent added yet
                    For Each f In (Await CType(Folder.FromParsingName(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), Nothing), Folder) _
                        .GetItems()).Where(Function(i) TypeOf i Is Folder AndAlso
                            Not _folders1.ToList().Exists(Function(f2) f2.FullPath = i.FullPath))
                        Dim fpure As Folder = Folder.FromParsingName(f.FullPath, Nothing)
                        If Not _folders1.ToList().Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                            Not _folders2.ToList().Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                            (fpure.Parent Is Nothing OrElse fpure.Parent.Parent Is Nothing) Then
                            _folders1.Insert(2, Folder.FromParsingName(f.FullPath, Nothing))
                        End If
                    Next

                    ' this computer & network
                    _folders3.Add(Folder.FromParsingName("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", Nothing))
                    _folders3.Add(Folder.FromParsingName("shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", Nothing))

                    _selectionHelper1 = New SelectionHelper(Of Folder)(_view.treeView1)
                    _selectionHelper1.SelectionChanged =
                        Sub()
                            UIHelper.OnUIThread(
                                Sub()
                                    If Not isWorking Then
                                        isWorking = True
                                        _selectionHelper2.SetSelectedItems({})
                                        _selectionHelper3.SetSelectedItems({})
                                        Me.OnSelectionChanged()
                                        NotifyOfPropertyChange("SelectedItem")
                                        NotifyOfPropertyChange("SelectedItems")
                                        isWorking = False
                                    End If
                                End Sub)
                        End Sub
                    _selectionHelper2 = New SelectionHelper(Of Folder)(_view.treeView2)
                    _selectionHelper2.SelectionChanged =
                        Sub()
                            UIHelper.OnUIThread(
                                Sub()
                                    If Not isWorking Then
                                        isWorking = True
                                        _selectionHelper1.SetSelectedItems({})
                                        _selectionHelper3.SetSelectedItems({})
                                        Me.OnSelectionChanged()
                                        NotifyOfPropertyChange("SelectedItem")
                                        NotifyOfPropertyChange("SelectedItems")
                                        isWorking = False
                                    End If
                                End Sub)
                        End Sub
                    _selectionHelper3 = New SelectionHelper(Of Folder)(_view.treeView3)
                    _selectionHelper3.SelectionChanged =
                        Sub()
                            UIHelper.OnUIThread(
                                Sub()
                                    If Not isWorking Then
                                        isWorking = True
                                        _selectionHelper1.SetSelectedItems({})
                                        _selectionHelper2.SetSelectedItems({})
                                        Me.OnSelectionChanged()
                                        NotifyOfPropertyChange("SelectedItem")
                                        NotifyOfPropertyChange("SelectedItems")
                                        isWorking = False
                                    End If
                                End Sub)
                        End Sub

                    AddHandler _view.PreviewMouseMove, AddressOf OnTreeViewPreviewMouseMove
                    AddHandler _view.PreviewMouseLeftButtonDown, AddressOf OnTreeViewPreviewMouseButtonDown
                    AddHandler _view.PreviewMouseRightButtonDown, AddressOf OnTreeViewPreviewMouseButtonDown
                    AddHandler _view.PreviewMouseUp, AddressOf OnTreeViewPreviewMouseButtonUp
                    AddHandler _view.MouseLeave, AddressOf OnTreeViewMouseLeave

                    _dropTarget = New TreeViewDropTarget(Me)
                    WpfDragTargetProxy.RegisterDragDrop(_view, _dropTarget)

                    If _view.Folder Is Nothing Then Shell.SetSelectedFolder(Shell.SpecialFolders("This computer"), Nothing)
                End Sub

            AddHandler System.Windows.Application.Current.MainWindow.Closed,
                Sub()
                    WpfDragTargetProxy.RevokeDragDrop(_view)
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
            If Me.SelectedItem.Equals(folder) Then
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

        Public ReadOnly Property Folders1 As ObservableCollection(Of Folder)
            Get
                Return _folders1
            End Get
        End Property

        Public ReadOnly Property Folders2 As ObservableCollection(Of Folder)
            Get
                Return _folders2
            End Get
        End Property

        Public ReadOnly Property Folders3 As ObservableCollection(Of Folder)
            Get
                Return _folders3
            End Get
        End Property

        Protected Overridable Sub OnSelectionChanged()
            If Not Me.SelectedItem Is Nothing Then
                If Not Me.SelectedItem.Equals(_view.Folder) Then
                    _view.Folder = Me.SelectedItem
                End If
            End If
        End Sub

        Public ReadOnly Property SelectedItem As Folder
            Get
                If _selectionHelper1.SelectedItems.Count = 1 Then
                    Return _selectionHelper1.SelectedItems(0)
                ElseIf _selectionHelper2.SelectedItems.Count = 1 Then
                    Return _selectionHelper2.SelectedItems(0)
                ElseIf _selectionHelper3.SelectedItems.Count = 1 Then
                    Return _selectionHelper3.SelectedItems(0)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        'Public Sub SetSelectedItem(value As Folder)
        '    If value Is Nothing Then
        '        _selectionHelper1.SetSelectedItems({})
        '        _selectionHelper2.SetSelectedItems({})
        '        _selectionHelper3.SetSelectedItems({})
        '    Else
        '        _selectionHelper1.SetSelectedItems({value})
        '        If Not _selectionHelper1.SelectedItems.Count = 1 Then
        '            _selectionHelper2.SetSelectedItems({value})
        '            If Not _selectionHelper2.SelectedItems.Count = 1 Then
        '                _selectionHelper3.SetSelectedItems({value})
        '            End If
        '        End If
        '    End If
        'End Sub

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

                    Dim tf As Folder, root As Integer
                    If _folders1.ToList().Exists(Function(f1) f1.FullPath = currentFolder.FullPath) Then
                        tf = _folders1.First(Function(f1) f1.FullPath = currentFolder.FullPath)
                        root = 1
                    ElseIf _folders2.ToList().Exists(Function(f1) f1.FullPath = currentFolder.FullPath) Then
                        tf = _folders2.First(Function(f1) f1.FullPath = currentFolder.FullPath)
                        root = 2
                    ElseIf _folders3.ToList().Exists(Function(f1) f1.FullPath = currentFolder.FullPath) Then
                        tf = _folders3.First(Function(f1) f1.FullPath = currentFolder.FullPath)
                        root = 3
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
                                        If Not tf.IsExpanded AndAlso Not tf.IsSelected Then tf._items = Nothing
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

                                        If Not callback Is Nothing Then
                                            callback(tf)
                                        Else
                                            Select Case root
                                                Case 1 : _selectionHelper1.SetSelectedItems({tf})
                                                Case 2 : _selectionHelper2.SetSelectedItems({tf})
                                                Case 3 : _selectionHelper3.SetSelectedItems({tf})
                                            End Select
                                        End If

                                        _isSettingSelectedFolder = False
                                    End If
                                Else
                                    If Not callback Is Nothing Then
                                        callback(Nothing)
                                    Else
                                        _selectionHelper1.SetSelectedItems({})
                                        _selectionHelper2.SetSelectedItems({})
                                        _selectionHelper3.SetSelectedItems({})
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
                        If Not callback Is Nothing Then
                            callback(folder)
                        Else
                            _view.Folder = folder
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
                Dim currentPointDown As Point = e.GetPosition(_view)
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                    Drag.Start({Me.SelectedItem}, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                End If
            End If
        End Sub

        Private Sub OnTreeViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(_view)

            If Not e.OriginalSource Is Nothing Then
                Dim treeViewItem As TreeViewItem = UIHelper.GetParentOfType(Of TreeViewItem)(e.OriginalSource)
                Dim clickedItem As Folder = treeViewItem?.DataContext
                _mouseItemDown = clickedItem
                If e.RightButton = MouseButtonState.Pressed Then
                    If Not clickedItem Is Nothing Then
                        If Me.SelectedItem Is Nothing Then Me.SetSelectedFolder(clickedItem)

                        Dim parent As Folder = clickedItem.Parent
                        If parent Is Nothing Then parent = Shell.Desktop

                        _menu = New ContextMenu()
                        AddHandler _menu.Click,
                        Sub(id As Integer, verb As String, ByRef isHandled As Boolean)
                            Select Case verb
                                Case "open"
                                    Me.SetSelectedFolder(clickedItem)
                                    isHandled = True
                            End Select
                        End Sub

                        Dim contextMenu As Controls.ContextMenu = _menu.GetContextMenu(parent, {clickedItem}, False)
                        _view.treeView1.ContextMenu = contextMenu
                        _view.treeView2.ContextMenu = contextMenu
                        _view.treeView3.ContextMenu = contextMenu
                        e.Handled = True
                    Else
                        _view.treeView1.ContextMenu = Nothing
                        _view.treeView2.ContextMenu = Nothing
                        _view.treeView3.ContextMenu = Nothing
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
    End Class
End Namespace