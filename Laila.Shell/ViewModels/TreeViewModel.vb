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
        Private _folders1 As List(Of TreeViewFolder)
        Private _folders2 As List(Of TreeViewFolder)
        Private _folders3 As List(Of TreeViewFolder)
        Private _selectionHelper1 As SelectionHelper(Of TreeViewFolder) = Nothing
        Private _selectionHelper2 As SelectionHelper(Of TreeViewFolder) = Nothing
        Private _selectionHelper3 As SelectionHelper(Of TreeViewFolder) = Nothing
        Private _isSettingSelectedFolder As Boolean
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _dropTarget As IDropTarget
        Private _menu As ContextMenu = New ContextMenu()

        Public Sub New(view As Controls.TreeView)
            _view = view

            ' home and galery
            _folders1 = New List(Of TreeViewFolder) From
            {
                TreeViewFolder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing),
                TreeViewFolder.FromParsingName("shell:::{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}", Nothing)
            }

            ' all the special folders under quick launch
            _folders2 = New List(Of TreeViewFolder)()
            Dim recentFolder As Folder = Folder.FromParsingName(Environment.GetFolderPath(Environment.SpecialFolder.Recent), Nothing, Nothing, 0)
            For Each f In CType(Folder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing, Nothing, 0), Folder) _
                .Items.Where(Function(i) TypeOf i Is Folder AndAlso Not IO.File.Exists(i.FullPath))
                If Not _folders1.Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                    Not _folders2.Exists(Function(f2) f2.FullPath = f.FullPath) Then
                    Dim tvf As TreeViewFolder = TreeViewFolder.FromParsingName(f.FullPath, Nothing)
                    'tvf.IsPinned = True
                    _folders2.Add(tvf)
                End If
            Next

            ' 5 first regular folders under quicklaunch
            For Each f In recentFolder _
                .Items.Where(Function(i) TypeOf i Is Folder).Take(5)
                If Not _folders1.Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                    Not _folders2.Exists(Function(f2) f2.FullPath = f.FullPath) Then
                    _folders2.Add(TreeViewFolder.FromParsingName(f.FullPath, Nothing))
                End If
            Next

            ' all special folders under user profile that we'rent added yet
            For Each f In CType(Folder.FromParsingName(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), Nothing, Nothing, 0), Folder) _
                .Items.Where(Function(i) TypeOf i Is Folder AndAlso
                    Not _folders1.Exists(Function(f2) f2.FullPath = i.FullPath))
                Dim fpure As Folder = Folder.FromParsingName(f.FullPath, Nothing, Nothing, 0)
                If Not _folders1.Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                    Not _folders2.Exists(Function(f2) f2.FullPath = f.FullPath) AndAlso
                    (fpure.Parent Is Nothing OrElse fpure.Parent.Parent Is Nothing) Then
                    _folders1.Insert(2, TreeViewFolder.FromParsingName(f.FullPath, Nothing))
                End If
            Next

            ' this computer & network
            _folders3 = New List(Of TreeViewFolder) From
            {
                TreeViewFolder.FromParsingName("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", Nothing),
                TreeViewFolder.FromParsingName("shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", Nothing)
            }

            Dim isWorking As Boolean
            AddHandler _view.Loaded,
                Sub(s As Object, e As EventArgs)
                    _selectionHelper1 = New SelectionHelper(Of TreeViewFolder)(_view.treeView1)
                    _selectionHelper1.SelectionChanged =
                        Sub()
                            UIHelper.OnUIThreadAsync(
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
                    _selectionHelper2 = New SelectionHelper(Of TreeViewFolder)(_view.treeView2)
                    _selectionHelper2.SelectionChanged =
                        Sub()
                            UIHelper.OnUIThreadAsync(
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
                    _selectionHelper3 = New SelectionHelper(Of TreeViewFolder)(_view.treeView3)
                    _selectionHelper3.SelectionChanged =
                        Sub()
                            UIHelper.OnUIThreadAsync(
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

                    _dropTarget = New TreeViewDropTarget(Me)
                    WpfDragTargetProxy.RegisterDragDrop(_view, _dropTarget)
                End Sub

            AddHandler System.Windows.Application.Current.MainWindow.Closed,
                Sub()
                    WpfDragTargetProxy.RevokeDragDrop(_view)
                End Sub

            AddHandler Shell.FolderNotification,
                Sub(s As Object, e As FolderNotificationEventArgs)
                    Select Case e.Event
                        Case SHCNE.RMDIR, SHCNE.DELETE, SHCNE.DRIVEREMOVED
                            If Not Me.SelectedItem Is Nothing AndAlso Me.IsSelectionWithin(Me.SelectedItem) Then
                                Me.SetSelectedItem(Me.SelectedItem.LogicalParent)
                            End If
                    End Select
                End Sub
        End Sub

        Private Function IsSelectionWithin(folder As TreeViewFolder) As Boolean
            If Me.SelectedItem.Equals(folder) Then
                Return True
            Else
                For Each f In folder.Folders
                    If Me.IsSelectionWithin(f) Then
                        Return True
                    End If
                Next
                Return False
            End If
        End Function

        Public ReadOnly Property Folders1 As List(Of TreeViewFolder)
            Get
                Return _folders1
            End Get
        End Property

        Public ReadOnly Property Folders2 As List(Of TreeViewFolder)
            Get
                Return _folders2
            End Get
        End Property

        Public ReadOnly Property Folders3 As List(Of TreeViewFolder)
            Get
                Return _folders3
            End Get
        End Property

        Protected Overridable Sub OnSelectionChanged()
            If Not Me.SelectedItem Is Nothing Then
                _view.LogicalParent = Me.SelectedItem.LogicalParent
                _view.SelectedFolderName = Me.SelectedItem.FullPath
            Else
                _view.LogicalParent = Nothing
                _view.SelectedFolderName = Nothing
            End If
        End Sub

        Public ReadOnly Property SelectedItem As TreeViewFolder
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

        Public Sub SetSelectedItem(value As TreeViewFolder)
            If value Is Nothing Then
                _selectionHelper1.SetSelectedItems({})
                _selectionHelper2.SetSelectedItems({})
                _selectionHelper3.SetSelectedItems({})
            Else
                _selectionHelper1.SetSelectedItems({value})
                If Not _selectionHelper1.SelectedItems.Count = 1 Then
                    _selectionHelper2.SetSelectedItems({value})
                    If Not _selectionHelper2.SelectedItems.Count = 1 Then
                        _selectionHelper3.SetSelectedItems({value})
                    End If
                End If
            End If
        End Sub

        Public Sub SetSelectedFolder(path As String)
            If Not _isSettingSelectedFolder Then
                _isSettingSelectedFolder = True

                Debug.WriteLine("SetSelectedFolder " & path)
                Dim list As List(Of Folder) = New List(Of Folder)()
                Dim f As Folder = Folder.FromParsingName(path, _view.LogicalParent, Nothing, 0)
                If Not f Is Nothing Then
                    While Not f.LogicalParent Is Nothing
                        list.Add(f)
                        Debug.WriteLine("SetSelectedFolder Added parent " & f.FullPath)
                        f = f.LogicalParent
                    End While

                    Dim tf As TreeViewFolder, root As Integer
                    If _folders1.Exists(Function(f1) f1.FullPath = f.FullPath) Then
                        tf = _folders1.First(Function(f1) f1.FullPath = f.FullPath)
                        root = 1
                    ElseIf _folders2.Exists(Function(f1) f1.FullPath = f.FullPath) Then
                        tf = _folders2.First(Function(f1) f1.FullPath = f.FullPath)
                        root = 2
                    ElseIf _folders3.Exists(Function(f1) f1.FullPath = f.FullPath) Then
                        tf = _folders3.First(Function(f1) f1.FullPath = f.FullPath)
                        root = 3
                    Else
                        tf = Nothing
                    End If

                    If Not tf Is Nothing Then
                        list.Reverse()

                        Dim func As Action(Of Folder, Action(Of Boolean)) =
                        Sub(item As Folder, callback As Action(Of Boolean))
                            Dim tf2 As TreeViewFolder
                            Dim thread As Thread = New Thread(New ThreadStart(
                                Sub()
                                    SyncLock tf._foldersLock
                                        If Not tf.IsExpanded Then tf._folders = Nothing
                                        tf2 = tf.Folders.FirstOrDefault(Function(f2) f2.FullPath = item.FullPath)
                                    End SyncLock
                                    If Not tf2 Is Nothing Then
                                        Debug.WriteLine("SetSelectedFolder found " & tf2.FullPath)
                                        tf.IsExpanded = True
                                        tf = tf2
                                        callback(False)
                                    Else
                                        Debug.WriteLine("SetSelectedFolder didn't find " & item.FullPath)
                                        callback(True)
                                    End If
                                End Sub))
                            thread.Start()
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

                                    Select Case root
                                        Case 1 : _selectionHelper1.SetSelectedItems({tf})
                                        Case 2 : _selectionHelper2.SetSelectedItems({tf})
                                        Case 3 : _selectionHelper3.SetSelectedItems({tf})
                                    End Select
                                    _isSettingSelectedFolder = False
                                End If
                            Else
                                Me.SetSelectedItem(Nothing)
                                _isSettingSelectedFolder = False
                            End If
                        End Sub
                        If en.MoveNext() Then
                            func(en.Current, cb)
                        Else
                            _isSettingSelectedFolder = False
                        End If
                    Else
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
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 7 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 7 Then
                    Drag.Start({Me.SelectedItem}, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                End If
            End If
        End Sub

        Private Sub OnTreeViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(_view)

            If Not e.OriginalSource Is Nothing Then
                Dim treeViewItem As TreeViewItem = UIHelper.GetParentOfType(Of TreeViewItem)(e.OriginalSource)
                Dim clickedItem As TreeViewFolder = treeViewItem?.DataContext
                _mouseItemDown = clickedItem
                If e.RightButton = MouseButtonState.Pressed Then
                    If Me.SelectedItem Is Nothing Then Me.SetSelectedItem(clickedItem)

                    Dim parent As Folder = clickedItem.Parent
                    If parent Is Nothing Then parent = Shell.Desktop

                    _menu = New ContextMenu()
                    AddHandler _menu.Click,
                        Sub(id As Integer, verb As String, ByRef isHandled As Boolean)
                            Select Case verb
                                Case "open"
                                    Me.SetSelectedItem(clickedItem)
                                    isHandled = True
                            End Select
                        End Sub

                    Dim contextMenu As System.Windows.Controls.ContextMenu = _menu.GetContextMenu(parent, {clickedItem}, False)
                    _view.treeView1.ContextMenu = contextMenu
                    _view.treeView2.ContextMenu = contextMenu
                    _view.treeView3.ContextMenu = contextMenu
                    e.Handled = True
                Else
                    _view.treeView1.ContextMenu = Nothing
                    _view.treeView2.ContextMenu = Nothing
                    _view.treeView3.ContextMenu = Nothing
                End If
            Else
                _mouseItemDown = Nothing
            End If
        End Sub
    End Class
End Namespace