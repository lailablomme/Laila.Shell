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
Imports System.Windows.Media
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.PinnedItems

Namespace Controls
    Public Class TreeView
        Inherits Control
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ItemsProperty As DependencyProperty = DependencyProperty.Register("Items", GetType(ObservableCollection(Of Item)), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
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

        Private PART_Grid As Grid
        Friend PART_ListBox As ListBox
        Friend PART_DragInsertIndicator As Grid
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _isSettingSelectedFolder As Boolean
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _dropTarget As IDropTarget
        Private _menu As RightClickMenu
        Private _frequentUpdateTimer As Timer
        Private disposedValue As Boolean
        Private _typeToSearchTimer As Timer
        Private _typeToSearchString As String = ""

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(TreeView), New FrameworkPropertyMetadata(GetType(TreeView)))
        End Sub

        Public Sub New()
            Shell.AddToControlCache(Me)

            Me.Items = New ObservableCollection(Of Item)()

            Dim view As ICollectionView = CollectionViewSource.GetDefaultView(Me.Items)
            view.Filter = AddressOf filter
            CType(view, ListCollectionView).CustomSort = New TreeViewComparer("TreeSortKey")
            AddHandler Me.Items.CollectionChanged, AddressOf items_CollectionChanged

            Me.Roots = New ObservableCollection(Of Item)()
            AddHandler Me.Roots.CollectionChanged, AddressOf roots_CollectionChanged

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

            AddHandler Shell.Notification,
                Sub(s As Object, e As NotificationEventArgs)
                    Select Case e.Event
                        Case SHCNE.RMDIR, SHCNE.DELETE, SHCNE.DRIVEREMOVED
                            UIHelper.OnUIThread(
                                Async Sub()
                                    If Not Me.SelectedItem Is Nothing AndAlso TypeOf e.Item1 Is Folder Then
                                        If Not Me.SelectedItem._parent Is Nothing AndAlso Me.SelectedItem.Pidl.Equals(e.Item1.Pidl) Then
                                            Await Me.SetSelectedFolder(Me.SelectedItem.Parent)
                                        ElseIf Not Me.SelectedItem.Pidl.Equals(e.Item1.Pidl) Then
                                            Dim f As Folder = Me.GetParentOfSelectionBefore(e.Item1)
                                            If Not f Is Nothing Then Await Me.SetSelectedFolder(f)
                                        End If
                                    End If
                                End Sub)
                    End Select
                    Select Case e.Event
                        Case SHCNE.RMDIR, SHCNE.DELETE
                            UIHelper.OnUIThread(
                                Async Sub()
                                    If Not Me.DoShowAllFoldersInTreeView AndAlso Not Me.Items.FirstOrDefault(Function(i) _
                                        Not i.disposedValue _
                                        AndAlso i.TreeRootIndex >= TreeRootSection.PINNED _
                                        AndAlso i.TreeRootIndex < TreeRootSection.ENVIRONMENT _
                                        AndAlso Not i.FullPath Is Nothing _
                                        AndAlso i.FullPath.Equals(e.Item1.FullPath)) Is Nothing Then
                                        Await updatePinnedItems()
                                        Await updateFrequentFolders()
                                    End If
                                End Sub)
                        Case SHCNE.UPDATEDIR
                            UIHelper.OnUIThread(
                                Async Sub()
                                    If Not Me.DoShowAllFoldersInTreeView AndAlso (Not Me.Items.FirstOrDefault(
                                        Function(i)
                                            If Not i.disposedValue Then
                                                Return i.TreeRootIndex >= TreeRootSection.PINNED _
                                                    AndAlso i.TreeRootIndex < TreeRootSection.ENVIRONMENT _
                                                    AndAlso (Not i.Parent Is Nothing _
                                                    AndAlso Not i.Parent.FullPath Is Nothing _
                                                    AndAlso i.Parent.FullPath.Equals(e.Item1.FullPath))
                                            Else
                                                Return False
                                            End If
                                        End Function) Is Nothing _
                                    OrElse Shell.Desktop.FullPath.Equals(e.Item1.FullPath)) Then
                                        Await updatePinnedItems()
                                        Await updateFrequentFolders()
                                    End If
                                End Sub)
                    End Select
                End Sub

            AddHandler PinnedItems.ItemPinned,
                 Async Sub(s2 As Object, e2 As PinnedItemEventArgs)
                     If Not Me.DoShowAllFoldersInTreeView Then
                         Await updatePinnedItems()
                         Await updateFrequentFolders()
                     End If
                 End Sub
            AddHandler PinnedItems.ItemUnpinned,
                 Async Sub(s2 As Object, e2 As PinnedItemEventArgs)
                     If Not Me.DoShowAllFoldersInTreeView Then
                         Await updatePinnedItems()
                         Await updateFrequentFolders()
                     End If
                 End Sub

            CollectionViewSource.GetDefaultView(Me.Items).Refresh()
        End Sub

        Private Sub loadRoots()
            If Not _frequentUpdateTimer Is Nothing Then
                _frequentUpdateTimer.Dispose()
                _frequentUpdateTimer = Nothing
            End If

            For Each item In Me.Roots.ToList()
                item.TreeRootIndex = -1
                Me.Roots.Remove(item)
            Next

            Dim currentFolder As Folder = Me.Folder
            Dim doShowLibrariesInTreeView As Boolean = Me.DoShowLibrariesInTreeView

            If Not Me.DoShowAllFoldersInTreeView Then
                Dim tcs As New TaskCompletionSource()
                Dim homeFolder As Folder, galleryFolder As Folder, librariesFolder As Folder, favoritesFolder As Folder
                Dim thisComputer As Folder, network As Folder

                Shell.STATaskQueue.Add(
                    Async Sub()
                        ' home, galery and favorites
                        If Shell.GetSpecialFolders().ContainsKey("Home") Then
                            homeFolder = Shell.GetSpecialFolder("Home").Clone()
                            homeFolder.TreeRootIndex = TreeRootSection.SYSTEM + 0
                        End If
                        If Shell.GetSpecialFolders().ContainsKey("Gallery") Then
                            galleryFolder = Shell.GetSpecialFolder("Gallery").Clone()
                            galleryFolder.TreeRootIndex = TreeRootSection.SYSTEM + 1
                        End If
                        If Shell.GetSpecialFolders().ContainsKey("Favorites") Then
                            favoritesFolder = Shell.GetSpecialFolder("Favorites").Clone()
                            favoritesFolder.TreeRootIndex = TreeRootSection.SYSTEM + 2
                        End If

                        ' this computer & network
                        thisComputer = Shell.GetSpecialFolder("This pc").Clone()
                        thisComputer.TreeRootIndex = TreeRootSection.ENVIRONMENT + 0
                        network = Shell.GetSpecialFolder("Network").Clone()
                        network.TreeRootIndex = TreeRootSection.ENVIRONMENT + 2

                        ' libraries
                        If Settings.IsWindows7OrLower OrElse doShowLibrariesInTreeView Then
                            librariesFolder = Shell.GetSpecialFolder("Libraries").Clone()
                            If Settings.IsWindows7OrLower Then
                                librariesFolder.TreeRootIndex = TreeRootSection.SYSTEM + 3
                            Else
                                librariesFolder.TreeRootIndex = TreeRootSection.ENVIRONMENT + 1
                            End If
                        End If

                        ' current
                        If Not currentFolder Is Nothing Then currentFolder = currentFolder.Clone()

                        tcs.SetResult()
                    End Sub)

                tcs.Task.Wait(Shell.ShuttingDownToken)
                ' system
                If Not Settings.IsWindows7OrLower Then
                    If Not homeFolder Is Nothing Then Me.Roots.Add(homeFolder)
                    If Not galleryFolder Is Nothing Then Me.Roots.Add(galleryFolder)
                Else
                    If Not favoritesFolder Is Nothing Then Me.Roots.Add(favoritesFolder)
                    If Not librariesFolder Is Nothing Then Me.Roots.Add(librariesFolder)
                End If
                ' separators
                Me.Roots.Add(New SeparatorFolder() With {.TreeRootIndex = TreeRootSection.PINNED - 1})
                Me.Roots.Add(New SeparatorFolder() With {.TreeRootIndex = TreeRootSection.ENVIRONMENT - 1})
                ' environment
                Me.Roots.Add(thisComputer)
                If Not Settings.IsWindows7OrLower AndAlso doShowLibrariesInTreeView Then Me.Roots.Add(librariesFolder)
                Me.Roots.Add(network)

                UIHelper.OnUIThread(
                    Async Sub()
                        Await updatePinnedItems()
                        Await updateFrequentFolders()
                    End Sub)

                ' frequent folders
                _frequentUpdateTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Async Sub()
                                Await updateFrequentFolders()
                            End Sub)
                    End Sub), Nothing, 1000 * 60, 1000 * 60)
            Else
                Dim tcs As New TaskCompletionSource()
                Dim desktopFolder As Folder

                Shell.STATaskQueue.Add(
                     Sub()
                         desktopFolder = Shell.Desktop.Clone()
                         desktopFolder.IsExpanded = True

                         ' current
                         If Not currentFolder Is Nothing Then currentFolder = currentFolder.Clone()

                         tcs.SetResult()
                     End Sub)

                tcs.Task.Wait(Shell.ShuttingDownToken)

                Me.Roots.Add(desktopFolder)
            End If

            Me.Folder = currentFolder
        End Sub

        Private Sub roots_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add
                    Dim currentRootIndex As Long = If(e.NewStartingIndex > 0, Me.Roots(e.NewStartingIndex - 1).TreeRootIndex + 1, 0)
                    For Each item In e.NewItems
                        Dim i As Item = item
                        If i.TreeRootIndex < 0 Then
                            maybeMoveRootOneUp(currentRootIndex)
                            i.TreeRootIndex = currentRootIndex
                            currentRootIndex += 1
                        Else
                            currentRootIndex = i.TreeRootIndex + 1
                        End If
                        Me.Items.Add(item)
                    Next
                Case NotifyCollectionChangedAction.Remove
                    For Each item In e.OldItems
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

        Private Function GetParentOfSelectionBefore(folder As Folder, Optional selectedItem As Folder = Nothing) As Folder
            If selectedItem Is Nothing Then selectedItem = Me.SelectedItem
            If selectedItem?.Pidl.Equals(folder.Pidl) Then
                Return selectedItem.Parent
            ElseIf Not selectedItem?.Parent Is Nothing Then
                Return Me.GetParentOfSelectionBefore(folder, selectedItem.Parent)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_ListBox = Template.FindName("PART_ListBox", Me)
            PART_Grid = Template.FindName("PART_Grid", Me)
            PART_DragInsertIndicator = Template.FindName("PART_DragInsertIndicator", Me)

            _selectionHelper = New SelectionHelper(Of Item)(PART_ListBox)

            AddHandler PART_ListBox.PreviewMouseMove, AddressOf OnTreeViewPreviewMouseMove
            AddHandler PART_ListBox.PreviewMouseLeftButtonDown, AddressOf OnTreeViewPreviewMouseButtonDown
            AddHandler PART_ListBox.PreviewMouseRightButtonDown, AddressOf OnTreeViewPreviewMouseButtonDown
            AddHandler PART_ListBox.PreviewMouseUp, AddressOf OnTreeViewPreviewMouseButtonUp
            AddHandler PART_ListBox.MouseLeave, AddressOf OnTreeViewMouseLeave
            AddHandler Me.PreviewKeyDown, AddressOf OnTreeViewKeyDown
            AddHandler Me.PreviewTextInput, AddressOf OnTreeViewTextInput

            _dropTarget = New TreeViewDropTarget(Me)
            WpfDragTargetProxy.RegisterDragDrop(PART_ListBox, _dropTarget)
        End Sub

        Private Async Function updatePinnedItems() As Task
            ' get pinned items
            Dim pinnedItemsList As IEnumerable(Of Item) = PinnedItems.GetPinnedItems()
            Dim count As Integer = 0
            For Each pinnedItem In pinnedItemsList
                Dim existingPinnedItem As Item = Me.Roots.FirstOrDefault(Function(i) _
                    i.FullPath = pinnedItem.FullPath _
                        AndAlso i.TreeRootIndex >= TreeRootSection.PINNED AndAlso i.TreeRootIndex < TreeRootSection.FREQUENT)
                If existingPinnedItem Is Nothing Then
                    ' insert new pinned item
                    pinnedItem.TreeRootIndex = TreeRootSection.PINNED + count
                    pinnedItem.IsPinned = True
                    Me.Roots.Add(pinnedItem)
                Else
                    ' update existing pinned item
                    pinnedItem.Dispose()
                    If existingPinnedItem.TreeRootIndex <> TreeRootSection.PINNED + count Then
                        existingPinnedItem.TreeRootIndex = TreeRootSection.PINNED + count
                        Me.Roots.Remove(existingPinnedItem)
                        Me.Roots.Add(existingPinnedItem)
                    End If
                End If
                count += 1
            Next
            ' remove pinned items no longer in the list
            For Each pinnedItem As Item In Me.Roots.Where(Function(i) _
                Not TypeOf i Is SeparatorFolder _
                    AndAlso i.TreeRootIndex >= TreeRootSection.PINNED AndAlso i.TreeRootIndex < TreeRootSection.FREQUENT _
                        AndAlso Not pinnedItemsList.ToList().Exists(Function(f) f.FullPath = i.FullPath)).ToList()
                If TypeOf pinnedItem Is Folder Then
                    Dim f As Folder = Me.GetParentOfSelectionBefore(pinnedItem)
                    If Not f Is Nothing Then Await Me.SetSelectedFolder(f)
                End If
                Me.Roots.Remove(pinnedItem)
            Next
            ' add/remove placeholder
            Dim placeholderFolder As PinnedAndFrequentPlaceholderFolder =
                    Me.Roots.FirstOrDefault(Function(i) TypeOf i Is PinnedAndFrequentPlaceholderFolder)
            If Me.Roots.Where(Function(i) _
                Not TypeOf i Is SeparatorFolder AndAlso Not TypeOf i Is PinnedAndFrequentPlaceholderFolder _
                    AndAlso i.TreeRootIndex >= TreeRootSection.PINNED AndAlso i.TreeRootIndex < TreeRootSection.ENVIRONMENT).Count > 0 Then
                If Not placeholderFolder Is Nothing Then Me.Roots.Remove(placeholderFolder)
            Else
                If placeholderFolder Is Nothing Then
                    Me.Roots.Add(New PinnedAndFrequentPlaceholderFolder() With {.TreeRootIndex = TreeRootSection.FREQUENT})
                End If
            End If
        End Function

        Private Async Function updateFrequentFolders() As Task
            ' get frequent folders
            Dim frequentFoldersList As IEnumerable(Of Folder) = FrequentFolders.GetMostFrequent()
            Dim count As Integer = 0
            For Each frequentFolder In frequentFoldersList
                Dim existingFrequentFolder As Folder = Me.Roots.FirstOrDefault(Function(i) _
                    i.FullPath = frequentFolder.FullPath _
                        AndAlso i.TreeRootIndex >= TreeRootSection.FREQUENT AndAlso i.TreeRootIndex < TreeRootSection.ENVIRONMENT)
                If existingFrequentFolder Is Nothing Then
                    ' insert new frequent folder
                    frequentFolder.TreeRootIndex = TreeRootSection.FREQUENT + count
                    Me.Roots.Add(frequentFolder)
                Else
                    ' update existing frequent folder
                    frequentFolder.Dispose()
                    If existingFrequentFolder.TreeRootIndex <> TreeRootSection.FREQUENT + count Then
                        existingFrequentFolder.TreeRootIndex = TreeRootSection.FREQUENT + count
                        Me.Roots.Remove(existingFrequentFolder)
                        Me.Roots.Add(existingFrequentFolder)
                    End If
                End If
                count += 1
            Next
            ' remove frequent folders no longer in the list
            For Each frequentFolder As Folder In Me.Roots.Where(Function(i) _
                Not TypeOf i Is SeparatorFolder AndAlso Not TypeOf i Is PinnedAndFrequentPlaceholderFolder _
                    AndAlso i.TreeRootIndex >= TreeRootSection.FREQUENT AndAlso i.TreeRootIndex < TreeRootSection.ENVIRONMENT _
                        AndAlso Not frequentFoldersList.ToList().Exists(Function(f) f.FullPath = i.FullPath)).ToList()
                Dim f As Folder = Me.GetParentOfSelectionBefore(frequentFolder)
                If Not f Is Nothing Then Await Me.SetSelectedFolder(f)
                Me.Roots.Remove(frequentFolder)
            Next
            ' add/remove placeholder
            Dim placeholderFolder As PinnedAndFrequentPlaceholderFolder =
                    Me.Roots.FirstOrDefault(Function(i) TypeOf i Is PinnedAndFrequentPlaceholderFolder)
            If Me.Roots.Where(Function(i) _
                Not TypeOf i Is SeparatorFolder AndAlso Not TypeOf i Is PinnedAndFrequentPlaceholderFolder _
                    AndAlso i.TreeRootIndex >= TreeRootSection.PINNED AndAlso i.TreeRootIndex < TreeRootSection.ENVIRONMENT).Count > 0 Then
                If Not placeholderFolder Is Nothing Then Me.Roots.Remove(placeholderFolder)
            Else
                If placeholderFolder Is Nothing Then
                    Me.Roots.Add(New PinnedAndFrequentPlaceholderFolder() With {.TreeRootIndex = TreeRootSection.FREQUENT})
                End If
            End If
        End Function

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
                    Dim tf As Folder
                    Dim currentFolder As Folder = folder
                    Dim noRecursive As List(Of String) = New List(Of String)()

                    If Not currentFolder Is Nothing Then
                        While Not currentFolder Is Nothing _
                        AndAlso Not noRecursive.Contains(currentFolder.Pidl.ToString()) _
                        AndAlso (currentFolder.TreeRootIndex <> -1 _
                                 OrElse Not currentFolder.Parent Is Nothing)
                            noRecursive.Add(currentFolder.Pidl.ToString())
                            list.Add(currentFolder)
                            Debug.WriteLine("SetSelectedFolder Added parent " & currentFolder.FullPath)
                            If currentFolder.TreeRootIndex <> -1 Then
                                tf = currentFolder
                                Exit While
                            End If
                            currentFolder = currentFolder.Parent
                        End While

                        list.Reverse()

                        Dim en As IEnumerator(Of Folder)
                        Dim en2 As IEnumerator(Of Folder) = list.GetEnumerator()
                        Dim en3 As IEnumerator(Of Item)
                        Dim func As Func(Of Folder, Func(Of Task), Task)
                        Dim cb As System.Func(Of Task)
                        Dim triedDesktop As Boolean = False

                        If Not tf Is Nothing Then en2.MoveNext()

                        Dim finish As Action(Of Folder) =
                            Sub(finishFolder As Folder)
                                If Not callback Is Nothing Then
                                    callback(finishFolder)
                                End If
                            End Sub

                        Dim findNextRoot2 As Action =
                            Async Sub()
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
                                    en3 = Me.Roots.GetEnumerator()
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
                                    triedDesktop = True
                                    findNextRoot()
                                Else
                                    Debug.WriteLine("SetSelectedFolder didn't find " & folder.FullPath & " -- giving up")
                                    If Not Me.Folder.FullPath = folder?.FullPath Then Me.Folder = folder
                                    _selectionHelper.SetSelectedItems({})
                                    finish(Nothing)
                                    _isSettingSelectedFolder = False
                                End If
                            End Function
                        func =
                        Async Function(item As Folder, callback2 As Func(Of Task)) As Task
                            Dim tf2 = (Await tf.GetItemsAsync()).FirstOrDefault(Function(f2) f2.FullPath = item.FullPath)
                            If Not tf2 Is Nothing Then
                                tf.IsExpanded = True
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

        Private Sub OnTreeViewPreviewMouseMove(sender As Object, e As MouseEventArgs)
            If Not _mouseItemDown Is Nothing AndAlso
                (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                Dim currentPointDown As Point = e.GetPosition(PART_ListBox)
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                    Drag.Start({_mouseItemDown}, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                    e.Handled = True
                End If
            End If
        End Sub

        Private Sub OnTreeViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(Me.PART_ListBox)

            If Not e.OriginalSource Is Nothing AndAlso UIHelper.GetParentOfType(Of ScrollBar)(e.OriginalSource) Is Nothing Then
                Dim treeViewItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
                Dim clickedItem As Item = TryCast(treeViewItem?.DataContext, Item)
                If Not TypeOf clickedItem Is SeparatorFolder Then
                    _mouseItemDown = clickedItem
                    Me.PART_ListBox.Focus()
                    If e.RightButton = MouseButtonState.Pressed Then
                        If Not clickedItem Is Nothing Then
                            If Me.SelectedItem Is Nothing Then
                                _selectionHelper.SetSelectedItems({clickedItem})
                                If TypeOf clickedItem Is Folder Then
                                    CType(clickedItem, Folder).LastScrollOffset = New Point()
                                    Me.Folder = clickedItem
                                End If
                            End If

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
                                                _selectionHelper.SetSelectedItems({clickedItem})
                                                Me.Folder = clickedItem
                                                e2.IsHandled = True
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
                            If TypeOf clickedItem Is Folder Then
                                CType(clickedItem, Folder).IsExpanded = Not CType(clickedItem, Folder).IsExpanded
                                e.Handled = True
                            Else
                                invokeDefaultCommand(clickedItem)
                            End If
                        End If
                    ElseIf e.LeftButton = MouseButtonState.Pressed Then
                        If Not clickedItem Is Nothing Then
                            If Not UIHelper.GetParentOfType(Of ToggleButton)(e.OriginalSource) Is Nothing Then
                                CType(clickedItem, Folder).IsExpanded = Not CType(clickedItem, Folder).IsExpanded
                            Else
                                Using Shell.OverrideCursor(Cursors.Wait)
                                    _selectionHelper.SetSelectedItems({clickedItem})

                                    ' this allows for better reponse to double-clicking an unselected item
                                    UIHelper.OnUIThread(
                                        Sub()
                                            If TypeOf clickedItem Is Folder Then
                                                CType(clickedItem, Folder).LastScrollOffset = New Point()
                                                Me.Folder = clickedItem
                                            End If
                                        End Sub, Threading.DispatcherPriority.Background)
                                End Using
                            End If
                            e.Handled = True
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

        Public Sub OnTreeViewPreviewMouseButtonUp(sender As Object, e As MouseButtonEventArgs)
            Me.PART_ListBox.ReleaseMouseCapture()
            _mouseItemDown = Nothing
        End Sub

        Public Sub OnTreeViewMouseLeave(sender As Object, e As MouseEventArgs)
            _mouseItemDown = Nothing
        End Sub

        Private Sub OnTreeViewKeyDown(sender As Object, e As KeyEventArgs)
            If Not TypeOf e.OriginalSource Is TextBox Then
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
                        Me.Folder = Me.SelectedItem
                    Else
                        invokeDefaultCommand(Me.SelectedItem)
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
                End If
            End If
        End Sub

        Private Sub OnTreeViewTextInput(sender As Object, e As TextCompositionEventArgs)
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
        End Sub

        Private Sub invokeDefaultCommand(item As Item)
            If Not _menu Is Nothing Then
                _menu.Dispose()
            End If
            _menu = New RightClickMenu() With {
                .Folder = Me.Folder,
                .SelectedItems = {item},
                .IsDefaultOnly = True
            }
            _menu.Make()
            _menu.InvokeCommand(_menu.DefaultId)
        End Sub

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
                                    For Each item2 In folder._items.Where(Function(i) TypeOf i Is Folder AndAlso i.IsVisibleInTree).ToList()
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
                                    For Each item2 In Me.Items.Where(Function(i) TypeOf i Is Folder _
                                        AndAlso Not i.Parent Is Nothing AndAlso i.Parent.Equals(folder)).ToList()
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
                        If TypeOf item Is Folder AndAlso CType(item, Item).IsVisibleInTree Then
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
                        If TypeOf item Is Folder AndAlso CType(item, Item).IsVisibleInTree Then
                            Me.Items.Add(item)
                        End If
                    Next
                    For Each item In e.OldItems
                        If TypeOf item Is Folder Then
                            Me.Items.Remove(item)
                        End If
                    Next
                Case NotifyCollectionChangedAction.Reset
                    Dim collection As ObservableCollection(Of Item) = s
                    Dim folder As Folder = Me.Items.FirstOrDefault(Function(i) TypeOf i Is Folder _
                        AndAlso Not CType(i, Folder).Items Is Nothing AndAlso CType(i, Folder).Items.Equals(collection))
                    If Not folder Is Nothing Then
                        For Each item In Me.Items.Where(Function(i) TypeOf i Is Folder _
                                AndAlso Not i.Parent Is Nothing AndAlso i.Parent.Equals(folder)).ToList()
                            Me.Items.Remove(item)
                        Next
                        For Each item In collection.Where(Function(i) _
                            Not i.disposedValue AndAlso i.IsVisibleInTree AndAlso TypeOf i Is Folder)
                            Me.Items.Add(item)
                        Next
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
                                For Each item In folder.Items
                                    If TypeOf item Is Folder Then
                                        If Not Me.Items.Contains(item) Then
                                            Me.Items.Add(item)
                                        End If
                                    End If
                                    If Me.SelectedItem Is Nothing AndAlso Not Me.Folder Is Nothing _
                                        AndAlso item.Pidl.Equals(Me.Folder.Pidl) Then
                                        Me.SetSelectedItem(item)
                                    End If
                                Next
                            Else
                                For Each item In Me.Items.Where(Function(i) TypeOf i Is Folder _
                                    AndAlso Not i.Parent Is Nothing AndAlso i.Parent.Equals(folder)).ToList()
                                    Me.Items.Remove(item)
                                Next
                            End If
                        End Sub)
                Case "TreeSortKey"
                    Dim list As List(Of Item)
                    UIHelper.OnUIThread(
                        Sub()
                            list = Me.Items.Where(Function(i) TypeOf i Is Folder _
                                AndAlso Not i.Parent Is Nothing AndAlso i.Parent.Equals(folder)).ToList()
                        End Sub)
                    For Each item2 In list
                        item2.NotifyOfPropertyChange("TreeSortKey")
                    Next
            End Select
        End Sub

        Private Function filter(i As Object) As Boolean
            Dim item As Item = i
            Dim isVisibleInTree As Boolean? = item?.IsVisibleInTree
            Return If(isVisibleInTree.HasValue, isVisibleInTree.Value, False)
        End Function


        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Shared Async Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Using Shell.OverrideCursor(Cursors.Wait)
                Dim tv As TreeView = TryCast(d, TreeView)
                Await tv.SetSelectedFolder(e.NewValue)
            End Using
        End Sub

        Public Property Items As ObservableCollection(Of Item)
            Get
                Return GetValue(ItemsProperty)
            End Get
            Protected Set(ByVal value As ObservableCollection(Of Item))
                SetCurrentValue(ItemsProperty, value)
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
            loadRoots()
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
            loadRoots()
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

        Public Enum TreeRootSection As Long
            SYSTEM = 0
            PINNED = 100
            FREQUENT = Long.MaxValue - 100
            ENVIRONMENT = Long.MaxValue - 10
        End Enum

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    If Not _typeToSearchTimer Is Nothing Then
                        _typeToSearchTimer.Dispose()
                        _typeToSearchTimer = Nothing
                    End If

                    If Not _frequentUpdateTimer Is Nothing Then
                        _frequentUpdateTimer.Dispose()
                        _frequentUpdateTimer = Nothing
                    End If

                    For Each item In Me.Items.ToList()
                        If Not TypeOf item Is SeparatorFolder Then
                            item.TreeRootIndex = -1
                            item.IsExpanded = False
                            item._parent = Nothing
                        End If
                    Next

                    WpfDragTargetProxy.RevokeDragDrop(PART_ListBox)

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