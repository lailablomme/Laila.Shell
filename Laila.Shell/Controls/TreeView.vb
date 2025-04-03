Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Media
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Input
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop.Application
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Windows
Imports Laila.Shell.PinnedItems

Namespace Controls
    Public Class TreeView
        Inherits Control
        Implements IDisposable, IProcessNotifications

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ItemsProperty As DependencyProperty = DependencyProperty.Register("Items", GetType(ObservableCollection(Of Item)), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SectionsProperty As DependencyProperty = DependencyProperty.Register("Sections", GetType(ObservableCollection(Of BaseTreeViewSection)), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly RootsProperty As DependencyProperty = DependencyProperty.Register("Roots", GetType(ObservableCollection(Of Item)), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColor", GetType(Boolean), GetType(TreeView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColorOverride", GetType(Boolean?), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged))
        Public Shared ReadOnly IsCompactModeProperty As DependencyProperty = DependencyProperty.Register("IsCompactMode", GetType(Boolean), GetType(TreeView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsCompactModeOverrideProperty As DependencyProperty = DependencyProperty.Register("IsCompactModeOverride", GetType(Boolean?), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsCompactModeOverrideChanged))
        Public Shared ReadOnly DoShowAllFoldersInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowAllFoldersInTreeView", GetType(Boolean), GetType(TreeView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowAllFoldersInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowAllFoldersInTreeViewOverride", GetType(Boolean?), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowAllFoldersInTreeViewOverrideChanged))
        Public Shared ReadOnly DoShowAvailabilityStatusInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowAvailabilityStatusInTreeView", GetType(Boolean), GetType(TreeView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowAvailabilityStatusInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowAvailabilityStatusInTreeViewOverride", GetType(Boolean?), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowAvailabilityStatusInTreeViewOverrideChanged))
        Public Shared ReadOnly DoExpandTreeViewToCurrentFolderProperty As DependencyProperty = DependencyProperty.Register("DoExpandTreeViewToCurrentFolder", GetType(Boolean), GetType(TreeView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoExpandTreeViewToCurrentFolderOverrideProperty As DependencyProperty = DependencyProperty.Register("DoExpandTreeViewToCurrentFolderOverride", GetType(Boolean?), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoExpandTreeViewToCurrentFolderOverrideChanged))
        Public Shared ReadOnly DoShowLibrariesInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowLibrariesInTreeView", GetType(Boolean), GetType(TreeView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowLibrariesInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowLibrariesInTreeViewOverride", GetType(Boolean?), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowLibrariesInTreeViewOverrideChanged))
        Public Shared ReadOnly DoShowPinnedAndFrequentItemsPlaceholderProperty As DependencyProperty = DependencyProperty.Register("DoShowPinnedAndFrequentItemsPlaceholder", GetType(Boolean), GetType(TreeView), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Property IsProcessingNotifications As Boolean = True Implements IProcessNotifications.IsProcessingNotifications

        Public Event BeforeFolderOpened As EventHandler(Of FolderEventArgs)
        Public Event AfterFolderOpened As EventHandler(Of FolderEventArgs)

        Private PART_Grid As Grid
        Friend PART_ListBox As ListBox
        Friend PART_DragInsertIndicator As Grid
        Private _dropTarget As IDropTarget
        Private _isLoaded As Boolean
        Private _isSettingSelectedFolder As Boolean
        Private _menu As RightClickMenu
        Private _mouseButton As MouseButton
        Private _mouseButtonState As MouseButtonState
        Private _mouseItemDown As Item
        Private _mousePointDown As Point
        Friend _scrollViewer As ScrollViewer
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _treeViewItemDown As ListBoxItem
        Private _typeToSearchTimer As Timer
        Private _typeToSearchString As String = ""
        Private disposedValue As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(TreeView), New FrameworkPropertyMetadata(GetType(TreeView)))
        End Sub

        Public Sub New()
            ' add us to the control cache so we'll get disposed on app shutdown
            Shell.AddToControlCache(Me)

            AddHandler Me.Loaded,
                Sub(s As Object, e As RoutedEventArgs)
                    ' get scrollviewer
                    _scrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(Me.PART_ListBox)(0)

                    If Not Me.PART_ListBox Is Nothing AndAlso Not _isLoaded Then
                        _isLoaded = True

                        ' load sections
                        loadSections()

                        ' dispose of ourselves when our parent window closes
                        AddHandler Window.GetWindow(Me).Closed,
                            Sub(s2 As Object, e2 As EventArgs)
                                Me.Dispose()
                            End Sub

                        ' create drop target
                        _dropTarget = New TreeViewDropTarget(Me)
                    End If

                    ' register for drag/drop on every load (may be multiple times i.e. in drop down of combobox)
                    If Not Me.PART_ListBox Is Nothing Then
                        WpfDragTargetProxy.RegisterDragDrop(PART_ListBox, _dropTarget)
                    End If
                End Sub

            Me.Items = New ObservableCollection(Of Item)()

            Dim view As ICollectionView = CollectionViewSource.GetDefaultView(Me.Items)
            view.Filter = AddressOf filter
            CType(view, ListCollectionView).CustomSort = New TreeViewComparer("TreeSortKey")
            AddHandler Me.Items.CollectionChanged, AddressOf items_CollectionChanged

            Me.Roots = New ObservableCollection(Of Item)()
            AddHandler Me.Roots.CollectionChanged, AddressOf roots_CollectionChanged

            Me.Sections = New ObservableCollection(Of BaseTreeViewSection)()
            AddHandler Me.Sections.CollectionChanged, AddressOf sections_CollectionChanged

            AddHandler Shell.Settings.PropertyChanged,
                Sub(s As Object, e As PropertyChangedEventArgs)
                    Select Case e.PropertyName
                        Case "DoShowEncryptedOrCompressedFilesInColor"
                            setDoShowEncryptedOrCompressedFilesInColor()
                        Case "IsCompactMode"
                            setIsCompactMode()
                        Case "DoShowAllFoldersInTreeView"
                            setDoShowAllFoldersInTreeView()
                        Case "DoShowAvailabilityStatusInTreeView"
                            setDoShowAvailabilityStatusInTreeView()
                        Case "DoExpandTreeViewToCurrentFolder"
                            setDoExpandTreeViewToCurrentFolder()
                        Case "DoShowLibrariesInTreeView"
                            setDoShowLibrariesInTreeView()
                    End Select
                End Sub
            setDoShowEncryptedOrCompressedFilesInColor()
            setIsCompactMode()
            setDoShowAllFoldersInTreeView()
            setDoShowAvailabilityStatusInTreeView()
            setDoExpandTreeViewToCurrentFolder()
            setDoShowLibrariesInTreeView()

            Shell.SubscribeToNotifications(Me)

            CollectionViewSource.GetDefaultView(Me.Items).Refresh()
        End Sub

        Protected Friend Overridable Sub ProcessNotification(e As NotificationEventArgs) Implements IProcessNotifications.ProcessNotification
            If Me.PART_ListBox Is Nothing Then Return

            Select Case e.Event
                Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                    ' this is for supporting file operations within .zip files 
                    ' from explorer or 7-zip
                    Dim selectedItem As Item = Nothing
                    If selectedItem Is Nothing Then
                        UIHelper.OnUIThread(
                            Sub()
                                selectedItem = Me.SelectedItem
                            End Sub)
                    End If
                    ' is something is being renamed to the currently selected folder...
                    If Not selectedItem Is Nothing _
                        AndAlso ((Not selectedItem.Pidl Is Nothing AndAlso Not e.Item2.Pidl Is Nothing AndAlso e.Item2.Pidl.Equals(selectedItem.Pidl)) _
                            OrElse ((selectedItem.Pidl Is Nothing OrElse e.Item2.Pidl Is Nothing) _
                                AndAlso (If(e.Item2.FullPath?.Equals(selectedItem.FullPath), False) OrElse If(e.Item2.FullPath?.Equals(selectedItem.FullPath.Split("~")(0)), False)))) Then
                        Shell.GlobalThreadPool.Add(
                            Sub()
                                ' get the first available parent in case the current folder disappears
                                Dim f As Folder = selectedItem.LogicalParent
                                If Not f Is Nothing Then
                                    Thread.Sleep(300) ' wait for .zip operations/folder refresh to complete

                                    ' get the newly created .zip folder
                                    Dim replacement As Item = f.Items.ToList().LastOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                        AndAlso ((Not i.Pidl Is Nothing AndAlso Not e.Item2.Pidl Is Nothing AndAlso If(i.Pidl?.Equals(e.Item2.Pidl), False)) _
                                            OrElse (i.Pidl Is Nothing OrElse e.Item2.Pidl Is Nothing) _
                                                AndAlso If(i.FullPath?.Equals(e.Item2.FullPath), False)))
                                    If Not replacement Is Nothing AndAlso TypeOf replacement Is Folder Then
                                        ' new folder matching the current folder was found -- select it
                                        UIHelper.OnUIThread(
                                            Sub()
                                                Me.SetSelectedItem(replacement)
                                            End Sub)
                                    Else
                                        ' the current folder disappeared -- switch to parent
                                        UIHelper.OnUIThread(
                                            Sub()
                                                Me.SetSelectedItem(f)
                                            End Sub)
                                    End If
                                End If
                            End Sub)
                    End If
                Case SHCNE.DELETE, SHCNE.RMDIR, SHCNE.DRIVEREMOVED
                    ' this sets the current folder to the first available parent 
                    ' when the current folder gets deleted
                    Dim selectedItem As Item = Nothing
                    If selectedItem Is Nothing Then
                        UIHelper.OnUIThread(
                            Sub()
                                selectedItem = Me.SelectedItem
                            End Sub)
                    End If
                    ' if the current folder was deleted...
                    If Not selectedItem Is Nothing _
                        AndAlso ((Not selectedItem.Pidl Is Nothing AndAlso Not e.Item1.Pidl Is Nothing AndAlso e.Item1.Pidl.Equals(selectedItem.Pidl)) _
                            OrElse ((selectedItem.Pidl Is Nothing OrElse e.Item1.Pidl Is Nothing) _
                                AndAlso If(e.Item1.FullPath?.Equals(selectedItem.FullPath), False))) Then
                        UIHelper.OnUIThread(
                            Sub()
                                ' get the first available parent  
                                Dim f As Folder = selectedItem.LogicalParent
                                If Not f Is Nothing Then
                                    Me.Folder = f ' load it
                                End If
                            End Sub)
                    End If
            End Select
        End Sub

        Private Sub loadSections()
            If Not _isLoaded Then Return

            Shell.IsSpecialFoldersReady.WaitOne()

            For Each section In Me.Sections.ToList()
                Me.Sections.Remove(section)
            Next

            Dim currentFolder As Folder = Me.Folder
            Dim doShowLibrariesInTreeView As Boolean = Me.DoShowLibrariesInTreeView

            If Not Me.DoShowAllFoldersInTreeView Then
                Me.Sections.Add(New SystemTreeViewSection())
                Me.Sections.Add(New PinnedItemsTreeViewSection())
                Me.Sections.Add(New FrequentFoldersTreeViewSection())
                Me.Sections.Add(New EnvironmentTreeViewSection())
            Else
                Me.Sections.Add(New AllFoldersTreeViewSection())
            End If
        End Sub

        Private Sub sections_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add
                    addSections(e.NewItems, e.NewStartingIndex)
                Case NotifyCollectionChangedAction.Remove
                    removeSections(e.OldItems)
                Case NotifyCollectionChangedAction.Reset
                    Throw New NotSupportedException()
            End Select
        End Sub

        Private Function getLastTreeRootIndexFromPreviousSection(currentSectionIndex As Integer) As Long
            For i = currentSectionIndex - 1 To 0 Step -1
                If Me.Sections(i).Items.Count > 0 Then
                    Return Me.Sections(i).Items.Last().TreeRootIndex
                End If
            Next
            Return -1
        End Function

        Private Function getCurrentRootIndex(previousTreeRootIndex As Long) As Integer
            Dim currentRootIndex As Integer = 0
            If previousTreeRootIndex >= 0 Then
                currentRootIndex = Me.Roots.IndexOf(Me.Roots.FirstOrDefault(Function(r) r.TreeRootIndex = previousTreeRootIndex))
                If currentRootIndex >= 0 Then
                    currentRootIndex += 1
                End If
            End If
            Return currentRootIndex
        End Function

        Private Sub addSections(newItems As IList, newStartingIndex As Integer)
            Dim previousTreeRootIndex As Long = If(newStartingIndex > 0, getLastTreeRootIndexFromPreviousSection(newStartingIndex), -1)
            Dim currentRootIndex As Integer = getCurrentRootIndex(previousTreeRootIndex)

            For Each section In newItems
                Dim sec As BaseTreeViewSection = section
                sec.TreeView = Me
                sec.Initialize()
                For Each item In sec.Items
                    Dim i As Item = item
                    Me.Roots.Insert(currentRootIndex, i)
                    currentRootIndex += 1
                Next
                AddHandler sec.Items.CollectionChanged, AddressOf sectionItems_CollectionChanged
            Next
        End Sub

        Private Sub removeSections(oldItems As IList)
            For Each section In oldItems
                Dim sec As BaseTreeViewSection = section
                RemoveHandler sec.Items.CollectionChanged, AddressOf sectionItems_CollectionChanged
                For Each item In sec.Items
                    Me.Roots.Remove(item)
                Next
            Next
        End Sub

        Private Sub sectionItems_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add
                    Dim previousTreeRootIndex As Long =
                        getLastTreeRootIndexFromPreviousSection _
                            (Me.Sections.IndexOf(Me.Sections.FirstOrDefault(Function(sec) sec.Items.Equals(s))))
                    Dim currentRootIndex As Integer =
                        getCurrentRootIndex(previousTreeRootIndex) _
                        + If(e.NewStartingIndex > 0, e.NewStartingIndex, 0)
                    For Each item In e.NewItems
                        Dim i As Item = item
                        Me.Roots.Insert(currentRootIndex, i)
                        currentRootIndex += 1
                    Next
                Case NotifyCollectionChangedAction.Remove
                    For Each item In e.OldItems
                        Me.Roots.Remove(item)
                    Next
                Case NotifyCollectionChangedAction.Reset
                    Throw New NotSupportedException()
            End Select
        End Sub

        Private Sub roots_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add
                    Dim currentTreeRootIndex As Long = If(e.NewStartingIndex > 0, Me.Roots(e.NewStartingIndex - 1).TreeRootIndex + 1, 0)
                    For Each item In e.NewItems
                        Dim i As Item = item
                        If i.TreeRootIndex < 0 Then
                            maybeMoveRootOneUp(currentTreeRootIndex)
                            i.TreeRootIndex = currentTreeRootIndex
                            currentTreeRootIndex += 1
                        Else
                            currentTreeRootIndex = i.TreeRootIndex + 1
                        End If
                        If TypeOf i Is SeparatorFolder AndAlso TypeOf Me.Items.OrderByDescending(Function(i2) i2.TreeRootIndex).FirstOrDefault(Function(i2) i2.TreeRootIndex < i.TreeRootIndex) Is SeparatorFolder Then
                            UIHelper.OnUIThreadAsync(
                                Sub()
                                    Me.Roots.Remove(i)
                                    For Each section In Me.Sections
                                        section.Items.Remove(i)
                                    Next
                                End Sub)
                        Else
                            Me.Items.Add(item)
                        End If
                    Next
                Case NotifyCollectionChangedAction.Remove
                    For Each item In e.OldItems
                        If TypeOf item Is Folder AndAlso CType(item, Folder).IsExpanded Then
                            CType(item, Folder).IsExpanded = False
                        End If
                        Me.Items.Remove(item)
                    Next
                Case Else
                    Throw New NotSupportedException()
            End Select
        End Sub

        Private Sub maybeMoveRootOneUp(rootIndex As Long)
            Dim itemToMove As Item = Me.Items.FirstOrDefault(Function(i) i.TreeRootIndex = rootIndex)
            If Not itemToMove Is Nothing Then
                maybeMoveRootOneUp(rootIndex + 1)
                itemToMove.TreeRootIndex = rootIndex + 1
            End If
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_ListBox = Template.FindName("PART_ListBox", Me)
            PART_Grid = Template.FindName("PART_Grid", Me)
            PART_DragInsertIndicator = Template.FindName("PART_DragInsertIndicator", Me)

            _selectionHelper = New SelectionHelper(Of Item)(PART_ListBox)

            AddHandler PART_ListBox.PreviewMouseMove, AddressOf listBox_PreviewMouseMove
            AddHandler PART_ListBox.PreviewMouseDown, AddressOf listBox_PreviewMouseDown
            AddHandler PART_ListBox.PreviewMouseUp, AddressOf listBox_PreviewMouseUp
            AddHandler PART_ListBox.MouseLeave, AddressOf listBox_MouseLeave
            AddHandler Me.PreviewKeyDown, AddressOf treeView_KeyDown
            AddHandler Me.PreviewTextInput, AddressOf treeView_TextInput
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

        Public Sub SetSelectedItem(value As Item)
            If value Is Nothing Then
                _selectionHelper.SetSelectedItems(New Item() {})
            Else
                _selectionHelper.SetSelectedItems(New Item() {value})
            End If
        End Sub

        Public Async Function SetSelectedFolder(folder As Folder, Optional callback As Action(Of Folder) = Nothing) As Task
            If Not _isSettingSelectedFolder Then
                _isSettingSelectedFolder = True

                If Me.DoExpandTreeViewToCurrentFolder Then
                    Await Task.Delay(100)

                    Debug.WriteLine("SetSelectedFolder " & folder?.FullPath)
                    Dim list As List(Of Folder) = New List(Of Folder)()
                    Dim tf As Folder = Nothing
                    Dim currentFolder As Folder = folder
                    Dim noRecursive As List(Of String) = New List(Of String)()
                    Dim foldersToExpand As List(Of Folder) = New List(Of Folder)()

                    If Not currentFolder Is Nothing Then
                        While Not currentFolder Is Nothing _
                        AndAlso Not noRecursive.Contains(currentFolder.Pidl?.ToString()) _
                        AndAlso (currentFolder.TreeRootIndex <> -1 _
                                 OrElse Not currentFolder.LogicalParent Is Nothing)
                            noRecursive.Add(currentFolder.Pidl?.ToString())
                            list.Add(currentFolder)
                            Debug.WriteLine("SetSelectedFolder Added parent " & currentFolder.FullPath)
                            If currentFolder.TreeRootIndex <> -1 Then
                                tf = currentFolder
                                Exit While
                            End If
                            currentFolder = currentFolder.LogicalParent
                        End While

                        list.Reverse()

                        Dim en As IEnumerator(Of Folder) = Nothing
                        Dim en2 As IEnumerator(Of Folder) = list.GetEnumerator()
                        Dim en3 As IEnumerator(Of Item) = Nothing
                        Dim func As Func(Of Folder, Func(Of Task), Task) = Nothing
                        Dim cb As System.Func(Of Task) = Nothing
                        Dim triedDesktop As Boolean = False

                        If Not tf Is Nothing Then en2.MoveNext()

                        Dim finish As Action(Of Folder) =
                            Sub(finishFolder As Folder)
                                If Not callback Is Nothing Then
                                    callback(finishFolder)
                                End If
                            End Sub

                        Dim findNextRoot2 As Action =
                            Sub()
                                While Not en3 Is Nothing AndAlso tf Is Nothing AndAlso en3.MoveNext()
                                    If en3.Current.FullPath = en2.Current.FullPath Then
                                        tf = en3.Current
                                    End If
                                End While
                            End Sub
                        Dim findNextRoot As Func(Of Task) =
                            Async Function() As Task
                                findNextRoot2()
                                While tf Is Nothing AndAlso en2.MoveNext()
                                    en3 = Me.Roots.ToList().GetEnumerator()
                                    findNextRoot2()
                                End While

                                If Not tf Is Nothing Then
                                    If tf.FullPath = folder.FullPath Then
                                        Await cb()
                                        Return
                                    Else
                                        en = list.GetEnumerator()
                                        While Not en2.Current.Equals(en.Current) AndAlso en.MoveNext
                                        End While
                                        If en.MoveNext() Then
                                            Await func(en.Current, cb)
                                            Return
                                        End If
                                    End If
                                End If

                                If Not Me.DoShowAllFoldersInTreeView AndAlso Not triedDesktop _
                            AndAlso (list.Count = 0 OrElse list(0).FullPath <> Shell.Desktop.FullPath) Then
                                    Debug.WriteLine("SetSelectedFolder didn't find " & folder.FullPath & " -- trying Desktop")
                                    list.Insert(0, Shell.Desktop)
                                    en2 = list.GetEnumerator()
                                    foldersToExpand = New List(Of Folder)()
                                    triedDesktop = True
                                    Await findNextRoot()
                                Else
                                    Debug.WriteLine("SetSelectedFolder didn't find " & folder.FullPath & " -- giving up")
                                    'If Not Me.Folder.FullPath = folder?.FullPath Then Me.Folder = folder
                                    _selectionHelper.SetSelectedItems({})
                                    finish(Nothing)
                                    _isSettingSelectedFolder = False
                                End If
                            End Function
                        func =
                        Async Function(item As Folder, callback2 As Func(Of Task)) As Task
                            Dim tf2 = (Await tf.GetItemsAsync(,, True)).FirstOrDefault(Function(f2) f2.FullPath = item.FullPath)
                            If Not tf2 Is Nothing Then
                                foldersToExpand.Add(tf)
                                Debug.WriteLine("SetSelectedFolder found " & tf2.FullPath)
                                tf = tf2
                                Await callback2()
                            Else
                                Debug.WriteLine("SetSelectedFolder didn't find " & item.FullPath & " -- keeping trying")
                                tf = Nothing
                                Await findNextRoot()
                            End If
                        End Function
                        cb =
                        Async Function() As Task
                            If Not en Is Nothing AndAlso en.MoveNext() Then
                                Await func(en.Current, cb)
                            Else
                                For Each f In foldersToExpand
                                    f.IsExpanded = True
                                Next
                                _selectionHelper.SetSelectedItems({tf})
                                If Not Me.Folder?.FullPath = tf?.FullPath Then Me.Folder = tf
                                finish(tf)
                                _isSettingSelectedFolder = False
                            End If
                        End Function

                        Await findNextRoot()
                    Else
                        _selectionHelper.SetSelectedItems({folder})
                        If Not Me.Folder?.FullPath = folder?.FullPath Then Me.Folder = folder
                        _isSettingSelectedFolder = False
                    End If
                Else
                    _isSettingSelectedFolder = False
                End If
            End If
        End Function

        Private Sub listBox_PreviewMouseMove(sender As Object, e As MouseEventArgs)
            If Not _mouseItemDown Is Nothing AndAlso
                (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                Dim currentPointDown As Point = e.GetPosition(PART_ListBox)
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                    _mouseButtonState = MouseButtonState.Released
                    If Not _treeViewItemDown Is Nothing Then _treeViewItemDown.Tag = "Released"
                    Drag.Start({_mouseItemDown}, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                    e.Handled = True
                End If
            End If
        End Sub

        Private Sub listBox_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
            _mouseButton = e.ChangedButton
            _mouseButtonState = e.ButtonState
            _mousePointDown = e.GetPosition(Me.PART_ListBox)

            Dim element As IInputElement = Me.PART_ListBox.InputHitTest(e.GetPosition(Me.PART_ListBox))
            If Not element Is Nothing AndAlso UIHelper.GetParentOfType(Of ScrollBar)(element) Is Nothing Then
                Dim treeViewItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(element)
                Dim clickedItem As Item = TryCast(treeViewItem?.DataContext, Item)
                If Not TypeOf clickedItem Is DummyFolder Then
                    _mouseItemDown = clickedItem
                    _treeViewItemDown = treeViewItem
                    _treeViewItemDown.Tag = "Pressed"
                    Me.PART_ListBox.Focus()
                    If e.RightButton = MouseButtonState.Pressed Then
                        If Not clickedItem Is Nothing Then
                            If Not _menu Is Nothing Then
                                _menu.Dispose()
                            End If
                            _menu = New RightClickMenu() With {
                                .Folder = If(clickedItem.Parent Is Nothing, Shell.Desktop, clickedItem.Parent),
                                .SelectedItems = {clickedItem}
                            }
                            AddHandler _menu.CommandInvoked,
                                Sub(s As Object, e2 As CommandInvokedEventArgs)
                                    Select Case e2.Verb
                                        Case "open"
                                            If TypeOf clickedItem Is Folder Then
                                                e2.IsHandled = True

                                                If (If(clickedItem.Pidl?.Equals(Me.Folder?.Pidl), False) _
                                                        OrElse (clickedItem.Pidl Is Nothing AndAlso Me.Folder.Pidl Is Nothing)) _
                                                        AndAlso Not Me.Items.Contains(Me.Folder) Then Return

                                                RaiseEvent BeforeFolderOpened(Me, New FolderEventArgs(clickedItem))
                                                _selectionHelper.SetSelectedItems({clickedItem})
                                                CType(clickedItem, Folder).LastScrollOffset = New Point()
                                                Me.Folder = clickedItem
                                                RaiseEvent AfterFolderOpened(Me, New FolderEventArgs(clickedItem))
                                            End If
                                        Case "rename"
                                            Dim getCoords As Menus.GetItemNameCoordinatesDelegate =
                                                Sub(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                    ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
                                                    point = Me.PointFromScreen(listBoxItem.PointToScreen(New Point(0, 0)))
                                                    point.X += clickedItem.TreeMargin.Left + 41
                                                    point.Y -= 0
                                                    size = New Size(Me.ActualWidth - point.X - 2, listBoxItem.ActualHeight)
                                                    textAlignment = TextAlignment.Left
                                                    fontSize = Me.FontSize
                                                End Sub
                                            Menus.DoRename(getCoords, Me.PART_Grid, treeViewItem, Me.PART_ListBox)
                                            e2.IsHandled = True
                                        Case "laila.shell.(un)pin"
                                            If e2.IsChecked Then
                                                PinnedItems.PinItem(clickedItem)
                                            Else
                                                PinnedItems.UnpinItem(clickedItem.Pidl)
                                            End If
                                            e2.IsHandled = True
                                    End Select
                                End Sub

                            PART_ListBox.ContextMenu = _menu
                            e.Handled = True
                        Else
                            PART_ListBox.ContextMenu = Nothing
                        End If
                    ElseIf e.LeftButton = MouseButtonState.Pressed AndAlso e.ClickCount = 2 Then
                        If Not clickedItem Is Nothing Then
                            Using Shell.OverrideCursor(Cursors.Wait)
                                If TypeOf clickedItem Is Folder Then
                                    CType(clickedItem, Folder).IsExpanded = Not CType(clickedItem, Folder).IsExpanded
                                    e.Handled = True
                                ElseIf TypeOf clickedItem Is Link AndAlso TypeOf CType(clickedItem, Link).TargetItem Is Folder Then
                                    e.Handled = True
                                Else
                                    Dim __ = invokeDefaultCommand(clickedItem)
                                End If
                            End Using
                        End If
                    End If

                    ' this whole trickery to prevent the ListBox from selecting other items while dragging:
                    Mouse.Capture(Me.PART_ListBox)
                Else
                    _mouseItemDown = Nothing
                    e.Handled = True
                End If
            Else
                _mouseItemDown = Nothing
            End If
        End Sub

        Public Sub listBox_PreviewMouseUp(sender As Object, e As MouseButtonEventArgs)
            If Not _treeViewItemDown Is Nothing Then _treeViewItemDown.Tag = "Released"

            Dim element As IInputElement = Me.PART_ListBox.InputHitTest(e.GetPosition(Me.PART_ListBox))
            If Not element Is Nothing AndAlso UIHelper.GetParentOfType(Of ScrollBar)(element) Is Nothing Then
                If Not _mouseItemDown Is Nothing AndAlso Not TypeOf _mouseItemDown Is DummyFolder Then
                    If _mouseButton = MouseButton.Left AndAlso _mouseButtonState = MouseButtonState.Pressed Then
                        If Not _mouseItemDown Is Nothing Then
                            Using Shell.OverrideCursor(Cursors.Wait)
                                If Not UIHelper.GetParentOfType(Of ToggleButton)(element) Is Nothing Then
                                    CType(_mouseItemDown, Folder).IsExpanded = Not CType(_mouseItemDown, Folder).IsExpanded
                                Else
                                    Using Shell.OverrideCursor(Cursors.Wait)
                                        _selectionHelper.SetSelectedItems({_mouseItemDown})

                                        ' this allows for better reponse to double-clicking an unselected item
                                        UIHelper.OnUIThread(
                                            Sub()
                                                If TypeOf _mouseItemDown Is Folder Then
                                                    If (If(_mouseItemDown.Pidl?.Equals(Me.Folder?.Pidl), False) _
                                                        OrElse (_mouseItemDown.Pidl Is Nothing AndAlso Me.Folder.Pidl Is Nothing)) _
                                                        AndAlso Not Me.Items.Contains(Me.Folder) Then Return

                                                    RaiseEvent BeforeFolderOpened(Me, New FolderEventArgs(_mouseItemDown))
                                                    CType(_mouseItemDown, Folder).LastScrollOffset = New Point()
                                                    Me.Folder = _mouseItemDown
                                                    RaiseEvent AfterFolderOpened(Me, New FolderEventArgs(_mouseItemDown))
                                                ElseIf TypeOf _mouseItemDown Is Link AndAlso TypeOf CType(_mouseItemDown, Link).TargetItem Is Folder Then
                                                    If (If(CType(_mouseItemDown, Link).TargetItem.Pidl?.Equals(Me.Folder?.Pidl), False) _
                                                        OrElse (CType(_mouseItemDown, Link).TargetItem.Pidl Is Nothing AndAlso Me.Folder.Pidl Is Nothing)) _
                                                        AndAlso Not Me.Items.Contains(Me.Folder) Then Return

                                                    RaiseEvent BeforeFolderOpened(Me, New FolderEventArgs(CType(_mouseItemDown, Link).TargetItem))
                                                    CType(CType(_mouseItemDown, Link).TargetItem, Folder).LastScrollOffset = New Point()
                                                    Me.Folder = CType(_mouseItemDown, Link).TargetItem
                                                    RaiseEvent AfterFolderOpened(Me, New FolderEventArgs(CType(_mouseItemDown, Link).TargetItem))
                                                End If
                                            End Sub, Threading.DispatcherPriority.Background)
                                    End Using
                                End If
                                e.Handled = True
                            End Using
                        End If
                    End If
                Else
                    e.Handled = True
                End If
            End If

            Me.PART_ListBox.ReleaseMouseCapture()
            _mouseItemDown = Nothing
        End Sub

        Public Sub listBox_MouseLeave(sender As Object, e As MouseEventArgs)
            _mouseItemDown = Nothing
        End Sub

        Private Sub treeView_KeyDown(sender As Object, e As KeyEventArgs)
            If Not TypeOf e.OriginalSource Is TextBox _
                AndAlso (e.OriginalSource.Equals(Me.PART_ListBox) OrElse UIHelper.GetParentOfType(Of ListBox)(e.OriginalSource)?.Equals(Me.PART_ListBox)) Then
                If e.Key = Key.C AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) _
                             AndAlso Not Me.SelectedItem Is Nothing AndAlso Not TypeOf Me.SelectedItem Is DummyFolder Then
                    Clipboard.CopyFiles({Me.SelectedItem})
                    e.Handled = True
                ElseIf e.Key = Key.X AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) _
                    AndAlso Not Me.SelectedItem Is Nothing AndAlso Not TypeOf Me.SelectedItem Is DummyFolder Then
                    Clipboard.CutFiles({Me.SelectedItem})
                    e.Handled = True
                ElseIf (e.Key = Key.Space OrElse e.Key = Key.Enter) AndAlso Keyboard.Modifiers = ModifierKeys.None _
                    AndAlso Not Me.SelectedItem Is Nothing AndAlso Not TypeOf Me.SelectedItem Is DummyFolder Then
                    If TypeOf Me.SelectedItem Is Folder Then
                        If (If(Me.SelectedItem.Pidl?.Equals(Me.Folder?.Pidl), False) OrElse (Me.SelectedItem.Pidl Is Nothing AndAlso Me.Folder.Pidl Is Nothing)) _
                            AndAlso Not Me.Items.Contains(Me.Folder) Then Return

                        RaiseEvent BeforeFolderOpened(Me, New FolderEventArgs(Me.SelectedItem))
                        CType(Me.SelectedItem, Folder).LastScrollOffset = New Point()
                        Me.Folder = Me.SelectedItem
                        RaiseEvent AfterFolderOpened(Me, New FolderEventArgs(Me.SelectedItem))
                    Else
                        Dim __ = invokeDefaultCommand(Me.SelectedItem)
                    End If
                    e.Handled = True
                ElseIf e.Key = Key.Add AndAlso Keyboard.Modifiers = ModifierKeys.None _
                    AndAlso TypeOf Me.SelectedItem Is Folder AndAlso Not TypeOf Me.SelectedItem Is DummyFolder Then
                    CType(Me.SelectedItem, Folder).IsExpanded = True
                    e.Handled = True
                ElseIf e.Key = Key.Subtract AndAlso Keyboard.Modifiers = ModifierKeys.None _
                    AndAlso TypeOf Me.SelectedItem Is Folder AndAlso Not TypeOf Me.SelectedItem Is DummyFolder Then
                    CType(Me.SelectedItem, Folder).IsExpanded = False
                    e.Handled = True
                ElseIf e.Key = Key.Up Then
                    Dim view As ICollectionView = CollectionViewSource.GetDefaultView(Me.Items)
                    For x = view.Cast(Of Item).ToList().IndexOf(Me.PART_ListBox.SelectedItem) - 1 To 0 Step -1
                        If CType(view(x), Item).IsVisibleInTree AndAlso Not TypeOf view(x) Is DummyFolder Then
                            Me.SetSelectedItem(view(x))
                            Exit For
                        End If
                    Next
                    e.Handled = True
                ElseIf e.Key = Key.Down Then
                    Dim view As ICollectionView = CollectionViewSource.GetDefaultView(Me.Items)
                    For x = view.Cast(Of Item).ToList().IndexOf(Me.PART_ListBox.SelectedItem) + 1 To view.Cast(Of Item).Count - 1
                        If CType(view(x), Item).IsVisibleInTree AndAlso Not TypeOf view(x) Is DummyFolder Then
                            Me.SetSelectedItem(view(x))
                            Exit For
                        End If
                    Next
                    e.Handled = True
                End If
            End If
        End Sub

        Private Sub treeView_TextInput(sender As Object, e As TextCompositionEventArgs)
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
                    Me.Items.Skip(Me.Items.IndexOf(Me.SelectedItem) + 1).Where(Function(i) i.IsVisibleInTree) _
                            .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                If foundItem Is Nothing Then
                    foundItem =
                        Me.Items.Take(Me.Items.IndexOf(Me.SelectedItem)).Where(Function(i) i.IsVisibleInTree) _
                                .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                End If
                If Not foundItem Is Nothing Then
                    Me.SetSelectedItem(foundItem)
                    e.Handled = True
                Else
                    SystemSounds.Asterisk.Play()
                End If
            End If
        End Sub

        Private Async Function invokeDefaultCommand(item As Item) As Task
            If Not _menu Is Nothing Then
                _menu.Dispose()
            End If
            _menu = New RightClickMenu() With {
                .Folder = Me.Folder,
                .SelectedItems = {item},
                .IsDefaultOnly = True
            }
            Await _menu.Make()
            Await _menu.InvokeCommand(_menu.DefaultId)
        End Function

        Private Sub items_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add
                    For Each item In e.NewItems
                        If TypeOf item Is Folder Then
                            Dim folder As Folder = item
                            AddHandler folder.PropertyChanged, AddressOf folder_PropertyChanged
                            AddHandler folder._items.CollectionChanged, AddressOf folder_CollectionChanged
                            UIHelper.OnUIThreadAsync(
                                Sub()
                                    For Each item2 In folder._items.ToList().Where(Function(i) i.IsVisibleInTree).ToList()
                                        If Not Me.Items.Contains(item2) Then
                                            Me.Items.Add(item2)
                                        End If
                                    Next
                                End Sub)
                        End If
                    Next
                Case NotifyCollectionChangedAction.Remove
                    For Each item In e.OldItems
                        If TypeOf item Is Folder Then
                            Dim folder As Folder = item
                            RemoveHandler folder.PropertyChanged, AddressOf folder_PropertyChanged
                            RemoveHandler folder._items.CollectionChanged, AddressOf folder_CollectionChanged
                            UIHelper.OnUIThreadAsync(
                                Sub()
                                    For Each item2 In Me.Items.Where(Function(i) i.CanShowInTree _
                                        AndAlso Not i._logicalParent Is Nothing AndAlso i._logicalParent.Equals(folder)).ToList()
                                        If Me.Items.Contains(item2) Then
                                            Me.Items.Remove(item2)
                                        End If
                                    Next
                                End Sub)
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
                        If CType(item, Item).IsVisibleInTree Then
                            Me.Items.Add(item)
                        End If
                    Next
                Case NotifyCollectionChangedAction.Remove
                    UIHelper.OnUIThreadAsync(
                        Async Sub()
                            ' give it some time before removing deleted items from the ui, so we get 
                            ' to switch to the matching new folder if applicable (i.e. for .zip operations)
                            Await Task.Delay(400)

                            ' remove items
                            For Each item In e.OldItems
                                If CType(item, Item).CanShowInTree Then
                                    Me.Items.Remove(item)
                                End If
                            Next
                        End Sub)
                Case NotifyCollectionChangedAction.Replace
                    UIHelper.OnUIThreadAsync(
                        Async Sub()
                            ' give it some time before removing deleted items from the ui, so we get 
                            ' to switch to the matching new folder if applicable (i.e. for .zip operations)
                            Await Task.Delay(400)

                            ' replace items
                            For Each item In e.NewItems
                                If CType(item, Item).IsVisibleInTree Then
                                    Me.Items.Add(item)
                                End If
                            Next
                            For Each item In e.OldItems
                                If CType(item, Item).CanShowInTree Then
                                    Me.Items.Remove(item)
                                End If
                            Next
                        End Sub)
                Case NotifyCollectionChangedAction.Reset
                    Dim collection As ObservableCollection(Of Item) = s
                    Dim folder As Folder = Me.Items.FirstOrDefault(Function(i) TypeOf i Is Folder _
                        AndAlso Not CType(i, Folder).Items Is Nothing AndAlso CType(i, Folder).Items.Equals(collection))
                    If Not folder Is Nothing Then
                        UIHelper.OnUIThreadAsync(
                            Async Sub()
                                Dim selectedItem As Item = Nothing
                                ' give it some time before removing deleted items from the ui, so we get 
                                ' to switch to the matching new folder if applicable (i.e. for .zip operations)
                                Await Task.Delay(400)

                                ' refresh folder
                                For Each item In Me.Items.Where(Function(i) i.CanShowInTree _
                                    AndAlso Not i._logicalParent Is Nothing AndAlso i._logicalParent.Equals(folder)).ToList()
                                    If Me.SelectedItem?.Equals(item) Then selectedItem = item
                                    Me.Items.Remove(item)
                                Next
                                For Each item In collection.Where(Function(i) _
                                    Not i.disposedValue AndAlso i.IsVisibleInTree).ToList()
                                    Me.Items.Add(item)
                                    If Not selectedItem Is Nothing _
                                        AndAlso ((Not selectedItem.Pidl Is Nothing AndAlso Not item.Pidl Is Nothing AndAlso If(item.Pidl?.Equals(selectedItem.Pidl), False)) _
                                            OrElse ((selectedItem.Pidl Is Nothing OrElse item.Pidl Is Nothing) AndAlso If(item.FullPath?.Equals(selectedItem.FullPath), False))) Then
                                        Me.SetSelectedItem(item)
                                    End If
                                Next
                            End Sub)
                    End If
            End Select
        End Sub

        Private Sub folder_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
            Dim folder As Folder = s

            Select Case e.PropertyName
                Case "IsExpanded"
                    UIHelper.OnUIThread(
                        Sub()
                            If folder.IsExpanded Then
                                ' when a folder is expanded, add it's children to the tree
                                For Each item In folder.Items.ToList()
                                    If item.CanShowInTree Then
                                        If Not Me.Items.Contains(item) Then
                                            Me.Items.Add(item)
                                        End If
                                    End If
                                    If Me.SelectedItem Is Nothing AndAlso Not Me.Folder Is Nothing _
                                        AndAlso item.Pidl?.Equals(Me.Folder.Pidl) Then
                                        Me.SetSelectedItem(item)
                                    End If
                                Next
                            Else
                                ' when a folder is collapsed, remove it's items from the tree
                                For Each item In Me.Items.Where(Function(i) i.CanShowInTree _
                                    AndAlso Not i._logicalParent Is Nothing AndAlso i._logicalParent.Equals(folder)).ToList()
                                    If item.IsExpanded Then item.IsExpanded = False
                                    Me.Items.Remove(item)
                                Next
                            End If
                        End Sub)
                Case "TreeSortKey"
                    ' when the tree sort order of an item is updated, also update it's children
                    Dim list As List(Of Item) = Nothing
                    UIHelper.OnUIThread(
                        Sub()
                            list = Me.Items.Where(Function(i) i.CanShowInTree AndAlso Not TypeOf i Is DummyFolder _
                                AndAlso Not i._logicalParent Is Nothing AndAlso i._logicalParent.Equals(folder)).ToList()
                        End Sub)
                    For Each item2 In list
                        item2.NotifyOfPropertyChange("TreeSortKey")
                    Next
            End Select
        End Sub

        ''' <summary>
        ''' Only show items that are supposed to be visible.
        ''' </summary>
        ''' <param name="i">The item for which to evaluate visibility</param>
        ''' <returns>True if the item will be visible</returns>
        Private Function filter(i As Object) As Boolean
            Dim item As Item = i
            Dim isVisibleInTree As Boolean? = item?.IsVisibleInTree
            Return If(isVisibleInTree.HasValue, isVisibleInTree.Value, False)
        End Function

        ''' <summary>
        ''' Gets/sets the current folder.
        ''' </summary>
        ''' <returns>The current Folder object</returns>
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
            Dim __ = tv.SetSelectedFolder(e.NewValue)
        End Sub

        Public Property Items As ObservableCollection(Of Item)
            Get
                Return GetValue(ItemsProperty)
            End Get
            Protected Set(ByVal value As ObservableCollection(Of Item))
                SetCurrentValue(ItemsProperty, value)
            End Set
        End Property

        Public Property Sections As ObservableCollection(Of BaseTreeViewSection)
            Get
                Return GetValue(SectionsProperty)
            End Get
            Protected Set(ByVal value As ObservableCollection(Of BaseTreeViewSection))
                SetCurrentValue(SectionsProperty, value)
            End Set
        End Property

        Public Property Roots As ObservableCollection(Of Item)
            Get
                Return GetValue(RootsProperty)
            End Get
            Protected Set(ByVal value As ObservableCollection(Of Item))
                SetCurrentValue(RootsProperty, value)
            End Set
        End Property

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
            Dim tv As TreeView = d
            tv.setDoShowEncryptedOrCompressedFilesInColor()
        End Sub

        Public Property IsCompactMode As Boolean
            Get
                Return GetValue(IsCompactModeProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(IsCompactModeProperty, value)
            End Set
        End Property

        Private Sub setIsCompactMode()
            If Me.IsCompactModeOverride.HasValue Then
                Me.IsCompactMode = Me.IsCompactModeOverride.Value
            Else
                Me.IsCompactMode = Shell.Settings.IsCompactMode
            End If
        End Sub

        Public Property IsCompactModeOverride As Boolean?
            Get
                Return GetValue(IsCompactModeOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(IsCompactModeOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnIsCompactModeOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As TreeView = d
            bfv.setIsCompactMode()
        End Sub

        Public Property DoShowAllFoldersInTreeView As Boolean
            Get
                Return GetValue(DoShowAllFoldersInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowAllFoldersInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowAllFoldersInTreeView()
            If Me.DoShowAllFoldersInTreeViewOverride.HasValue Then
                Me.DoShowAllFoldersInTreeView = Me.DoShowAllFoldersInTreeViewOverride.Value
            Else
                Me.DoShowAllFoldersInTreeView = Shell.Settings.DoShowAllFoldersInTreeView
            End If
            loadSections()
        End Sub

        Public Property DoShowAllFoldersInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowAllFoldersInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowAllFoldersInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowAllFoldersInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As TreeView = d
            bfv.setDoShowAllFoldersInTreeView()
        End Sub

        Public Property DoShowAvailabilityStatusInTreeView As Boolean
            Get
                Return GetValue(DoShowAvailabilityStatusInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowAvailabilityStatusInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowAvailabilityStatusInTreeView()
            If Me.DoShowAvailabilityStatusInTreeViewOverride.HasValue Then
                Me.DoShowAvailabilityStatusInTreeView = Me.DoShowAvailabilityStatusInTreeViewOverride.Value
            Else
                Me.DoShowAvailabilityStatusInTreeView = Shell.Settings.DoShowAvailabilityStatusInTreeView
            End If
        End Sub

        Public Property DoShowAvailabilityStatusInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowAvailabilityStatusInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowAvailabilityStatusInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowAvailabilityStatusInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As TreeView = d
            bfv.setDoShowAvailabilityStatusInTreeView()
        End Sub

        Public Property DoExpandTreeViewToCurrentFolder As Boolean
            Get
                Return GetValue(DoExpandTreeViewToCurrentFolderProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoExpandTreeViewToCurrentFolderProperty, value)
            End Set
        End Property

        Private Sub setDoExpandTreeViewToCurrentFolder()
            If Me.DoExpandTreeViewToCurrentFolderOverride.HasValue Then
                Me.DoExpandTreeViewToCurrentFolder = Me.DoExpandTreeViewToCurrentFolderOverride.Value
            Else
                Me.DoExpandTreeViewToCurrentFolder = Shell.Settings.DoExpandTreeViewToCurrentFolder
            End If
        End Sub

        Public Property DoExpandTreeViewToCurrentFolderOverride As Boolean?
            Get
                Return GetValue(DoExpandTreeViewToCurrentFolderOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoExpandTreeViewToCurrentFolderOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoExpandTreeViewToCurrentFolderOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As TreeView = d
            bfv.setDoExpandTreeViewToCurrentFolder()
        End Sub

        Public Property DoShowLibrariesInTreeView As Boolean
            Get
                Return GetValue(DoShowLibrariesInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowLibrariesInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowLibrariesInTreeView()
            If Me.DoShowLibrariesInTreeViewOverride.HasValue Then
                Me.DoShowLibrariesInTreeView = Me.DoShowLibrariesInTreeViewOverride.Value
            Else
                Me.DoShowLibrariesInTreeView = Shell.Settings.DoShowLibrariesInTreeView
            End If
            loadSections()
        End Sub

        Public Property DoShowLibrariesInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowLibrariesInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowLibrariesInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowLibrariesInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As TreeView = d
            bfv.setDoShowLibrariesInTreeView()
        End Sub

        Public Property DoShowPinnedAndFrequentItemsPlaceholder As Boolean
            Get
                Return GetValue(DoShowPinnedAndFrequentItemsPlaceholderProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(DoShowPinnedAndFrequentItemsPlaceholderProperty, value)
            End Set
        End Property

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    Shell.UnsubscribeFromNotifications(Me)

                    ' dispose of "type to search" timer
                    If Not _typeToSearchTimer Is Nothing Then
                        _typeToSearchTimer.Dispose()
                        _typeToSearchTimer = Nothing
                    End If

                    ' dispose of sections
                    For Each section In Me.Sections
                        section.Dispose()
                    Next

                    ' neutralize items
                    For Each item In Me.Items.ToList()
                        If Not TypeOf item Is SeparatorFolder Then
                            item.TreeRootIndex = -1
                            item.IsExpanded = False
                            item._parent = Nothing
                        End If
                    Next

                    ' unsubscribe as a drop target
                    If Not Me.PART_ListBox Is Nothing Then
                        WpfDragTargetProxy.RevokeDragDrop(PART_ListBox)
                    End If

                    ' we don't need to be disposed on shutdown anymore, we're already disposed
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