Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Windows

Namespace Controls.Parts
    Public Class TreeViewDropTarget
        Inherits BaseDropTarget

        Private _dataObject As ComTypes.IDataObject
        Private _treeView As Laila.Shell.Controls.TreeView
        Private _lastOverItem As Item
        Private _lastDropTarget As IDropTarget
        Private _dragOpenTimer As DispatcherTimer
        Private _scrollUpTimer As DispatcherTimer
        Private _scrollDownTimer As DispatcherTimer
        Private _scrollDirection As Boolean?
        Private _prevSelectedItem As Item
        Private _fileNameList() As String
        Private _files As List(Of Item)
        Private _dragInsertParent As ISupportDragInsert = Nothing
        Private _insertIndex As Long = -2

        Public Sub New(treeView As Laila.Shell.Controls.TreeView)
            _treeView = treeView
        End Sub

        Public Overrides Function DragEnter(pDataObj As ComTypes.IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            _dataObject = pDataObj
            _fileNameList = Clipboard.GetFileNameList(pDataObj)
            _files = ClipboardFormats.CFSTR_SHELLIDLIST.GetData(pDataObj, True)
            _prevSelectedItem = _treeView.SelectedItem
            _treeView.PART_ListBox.Focus()
            Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
        End Function

        Public Overrides Function DragOver(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
        End Function

        Public Overrides Function DragLeave() As Integer
            ' clean up
            If Not _files Is Nothing Then
                For Each f In _files
                    f.LogicalParent.Dispose()
                    f.Dispose()
                Next
            End If
            If Not _dragOpenTimer Is Nothing Then
                _dragOpenTimer.IsEnabled = False
            End If
            If Not _scrollUpTimer Is Nothing Then
                _scrollUpTimer.IsEnabled = False
            End If
            If Not _scrollDownTimer Is Nothing Then
                _scrollDownTimer.IsEnabled = False
            End If
            _scrollDirection = Nothing
            _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
            _insertIndex = -2
            _lastOverItem = Nothing
            If Not _lastDropTarget Is Nothing Then
                Try
                    Return _lastDropTarget.DragLeave()
                Finally
                    If Not _lastDropTarget Is Nothing Then
                        Marshal.ReleaseComObject(_lastDropTarget)
                        _lastDropTarget = Nothing
                    End If
                    _treeView.SetSelectedItem(_prevSelectedItem)
                End Try
            Else
                _treeView.SetSelectedItem(_prevSelectedItem)
            End If
            Return 0
        End Function

        Public Overrides Function Drop(pDataObj As ComTypes.IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            ' clean up
            If Not _dragOpenTimer Is Nothing Then
                _dragOpenTimer.IsEnabled = False
            End If
            If Not _scrollUpTimer Is Nothing Then
                _scrollUpTimer.IsEnabled = False
            End If
            If Not _scrollDownTimer Is Nothing Then
                _scrollDownTimer.IsEnabled = False
            End If
            _scrollDirection = Nothing
            _lastOverItem = Nothing

            ' we're inserting
            If Not _dragInsertParent Is Nothing Then
                CType(_dragInsertParent, ISupportDragInsert).Drop(_dataObject, _files, _insertIndex)
                _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
            End If

            ' we've got an actual drop target
            If Not _lastDropTarget Is Nothing Then
                Try
                    Dim overListBoxItem As ListBoxItem = getOverListBoxItem(ptWIN32)
                    Dim overItem As Item = overListBoxItem?.DataContext
                    If Not overItem Is Nothing AndAlso overItem.FullPath = Shell.GetSpecialFolder(SpecialFolders.RecycleBin).FullPath Then
                        Dim fo As IFileOperation = Nothing
                        Try
                            fo = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_FileOperation))
                            If grfKeyState.HasFlag(MK.MK_SHIFT) Then fo.SetOperationFlags(FOF.FOFX_WANTNUKEWARNING)
                            fo.DeleteItems(_dataObject)
                            fo.PerformOperations()
                            Return HRESULT.S_OK
                        Finally
                            If Not fo Is Nothing Then
                                Marshal.ReleaseComObject(fo)
                                fo = Nothing
                            End If
                        End Try
                    Else
                        Dim h As HRESULT = _lastDropTarget.Drop(pDataObj, grfKeyState, ptWIN32, pdwEffect)
                        Debug.WriteLine("drop=" & h.ToString())
                        Return h
                    End If
                Finally
                    If Not _lastDropTarget Is Nothing Then
                        Marshal.ReleaseComObject(_lastDropTarget)
                        _lastDropTarget = Nothing
                    End If
                    _treeView.SetSelectedItem(_prevSelectedItem)
                    If TypeOf _prevSelectedItem Is Folder Then _treeView.SetSelectedFolder(_prevSelectedItem)
                End Try
            Else
                _treeView.SetSelectedItem(_prevSelectedItem)
                If TypeOf _prevSelectedItem Is Folder Then _treeView.SetSelectedFolder(_prevSelectedItem)
            End If

            ' clean up
            If Not _files Is Nothing Then
                For Each f In _files
                    f.LogicalParent.Dispose()
                    f.Dispose()
                Next
            End If

            Return HRESULT.S_OK
        End Function

        Private Function getOverListBoxItem(ptWIN32 As WIN32POINT) As ListBoxItem
            ' translate point to listview
            Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _treeView)

            ' find which item we're over
            Dim overTreeViewItem As ListBoxItem
            Dim overObject As IInputElement = _treeView.InputHitTest(pt)
            If TypeOf overObject Is ListBoxItem Then
                overTreeViewItem = overObject
            Else
                overTreeViewItem = UIHelper.GetParentOfType(Of ListBoxItem)(overObject)
            End If
            ' apprently there's some space between the items, so try again
            If overTreeViewItem Is Nothing Then
                overObject = _treeView.InputHitTest(New Point(pt.X, pt.Y - 2))
            End If
            If TypeOf overObject Is ListBoxItem Then
                overTreeViewItem = overObject
            Else
                overTreeViewItem = UIHelper.GetParentOfType(Of ListBoxItem)(overObject)
            End If
            ' apprently there's some space between the items, so try again
            If overTreeViewItem Is Nothing Then
                overObject = _treeView.InputHitTest(New Point(pt.X, pt.Y + 2))
            End If
            If TypeOf overObject Is ListBoxItem Then
                overTreeViewItem = overObject
            Else
                overTreeViewItem = UIHelper.GetParentOfType(Of ListBoxItem)(overObject)
            End If

            Return If(Not overTreeViewItem Is Nothing,
                overTreeViewItem, Nothing)
        End Function

        Private Function dragPoint(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As UInteger) As Integer
            If Not If(_fileNameList?.Count, 0) > 0 Then
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HRESULT.S_OK
            End If

            ' scroll up and down while dragging?
            Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _treeView)
            If pt.Y < 100 Then
                If _scrollUpTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> False Then
                    _scrollDirection = False
                    If Not _scrollDownTimer Is Nothing Then _scrollDownTimer.IsEnabled = False
                    If Not _scrollUpTimer Is Nothing Then
                        _scrollUpTimer.IsEnabled = True
                    Else
                        _scrollUpTimer = New DispatcherTimer()
                        AddHandler _scrollUpTimer.Tick,
                            Sub(s2 As Object, e As EventArgs)
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_treeView.PART_ListBox)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset - 50)
                            End Sub
                        _scrollUpTimer.Interval = TimeSpan.FromMilliseconds(350)
                        _scrollUpTimer.IsEnabled = True
                    End If
                End If
            ElseIf pt.Y > _treeView.ActualHeight - 100 Then
                If _scrollDownTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> True Then
                    _scrollDirection = True
                    If Not _scrollUpTimer Is Nothing Then _scrollUpTimer.IsEnabled = False
                    If Not _scrollDownTimer Is Nothing Then
                        _scrollDownTimer.IsEnabled = True
                    Else
                        _scrollDownTimer = New DispatcherTimer()
                        AddHandler _scrollDownTimer.Tick,
                            Sub(s2 As Object, e As EventArgs)
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_treeView.PART_ListBox)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset + 50)
                            End Sub
                        _scrollDownTimer.Interval = TimeSpan.FromMilliseconds(350)
                        _scrollDownTimer.IsEnabled = True
                    End If
                End If
            Else
                If Not _scrollUpTimer Is Nothing Then
                    _scrollUpTimer.IsEnabled = False
                End If
                If Not _scrollDownTimer Is Nothing Then
                    _scrollDownTimer.IsEnabled = False
                End If
                _scrollDirection = Nothing
            End If

            ' get the item we're currently dragging over
            Dim overListBoxItem As ListBoxItem = getOverListBoxItem(ptWIN32)
            Dim overItem As Item = overListBoxItem?.DataContext

            ' possible insert index
            Dim insertIndex As Long = -1

            ' if we're actually over an item...
            If Not overItem Is Nothing Then
                ' check if we're about to insert and where
                If _fileNameList?.Count > 0 Then
                    Dim ptItem As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, overListBoxItem)
                    If ptItem.Y <= 3 AndAlso Not TypeOf overItem Is SeparatorFolder _
                        AndAlso TypeOf If(overItem.TreeViewSection, If(overItem.LogicalParent, overItem.Parent)) Is ISupportDragInsert Then
                        _dragInsertParent = If(overItem.TreeViewSection, If(overItem.LogicalParent, overItem.Parent))
                        _treeView.PART_DragInsertIndicator.Margin =
                            New Thickness(0, _treeView.PointFromScreen(overListBoxItem.PointToScreen(
                                New Point(0, 0))).Y - 3, 0, 0)
                        insertIndex = _dragInsertParent.Items.Where(Function(i) i.IsVisibleInTree AndAlso Not TypeOf i Is DummyFolder) _
                            .OrderBy(Function(i) i.TreeSortKey).ToList().IndexOf(overItem)
                        Debug.WriteLine("* Drag: insert before item")
                    ElseIf ptItem.Y >= overListBoxItem.ActualHeight - 3 AndAlso (Not overItem.IsFolder OrElse Not overItem.IsExpanded OrElse CType(overItem, Folder).Items.Count = 0) Then
                        Dim child As Object = overItem
                        Dim dip As Object = If(overItem.TreeViewSection, If(overItem.LogicalParent, overItem.Parent))
                        While Not dip Is Nothing _
                            AndAlso TypeOf dip Is Folder _
                            AndAlso Not TypeOf dip Is ISupportDragInsert _
                            AndAlso CType(dip, Folder).Items.Where(Function(i) i.IsVisibleInTree).OrderBy(Function(i) i.TreeSortKey).ToList().IndexOf(child) _
                                    = CType(dip, Folder).Items.Where(Function(i) i.IsVisibleInTree).Count - 1
                            child = dip
                            dip = If(CType(dip, Folder).TreeViewSection, If(CType(dip, Folder).LogicalParent, CType(dip, Folder).Parent))
                        End While
                        If Not dip Is Nothing AndAlso TypeOf dip Is ISupportDragInsert Then
                            _dragInsertParent = dip
                            Debug.WriteLine("* Drag: insert after not expanded item")
                            _treeView.PART_DragInsertIndicator.Margin =
                                New Thickness(0, _treeView.PointFromScreen(overListBoxItem.PointToScreen(
                                    New Point(0, overListBoxItem.ActualHeight))).Y - 3, 0, 0)
                            insertIndex = _dragInsertParent.Items.Where(Function(i) i.IsVisibleInTree AndAlso Not TypeOf i Is DummyFolder) _
                                .OrderBy(Function(i) i.TreeSortKey).ToList().IndexOf(child) + 1
                        Else
                            _dragInsertParent = Nothing
                        End If
                    ElseIf ptItem.Y >= overListBoxItem.ActualHeight - 3 Then
                        If TypeOf overItem Is Folder AndAlso TypeOf overItem Is ISupportDragInsert Then
                            Debug.WriteLine("* Drag: insert after expanded item")
                            _treeView.PART_DragInsertIndicator.Margin =
                            New Thickness(0, _treeView.PointFromScreen(overListBoxItem.PointToScreen(
                                New Point(0, overListBoxItem.ActualHeight))).Y - 3, 0, 0)
                            _dragInsertParent = overItem
                            insertIndex = 0
                        Else
                            _dragInsertParent = Nothing
                        End If
                    Else
                        Debug.WriteLine("* Full over item")
                        _dragInsertParent = Nothing
                    End If
                Else
                    Debug.WriteLine("* Not over item")
                    _dragInsertParent = Nothing
                End If

                Debug.WriteLine("_newPinnedIndex=" & insertIndex & "   overItem.TreeRootIndex=" & overItem.TreeRootIndex & "  dispn=" & overItem.DisplayName)

                ' if we're over a folder, expand it after two seconds of hovering
                If TypeOf overItem Is Folder AndAlso _dragInsertParent Is Nothing Then
                    If _lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem) OrElse insertIndex <> _insertIndex Then
                        If Not _dragOpenTimer Is Nothing Then
                            _dragOpenTimer.Tag = overItem
                            _dragOpenTimer.IsEnabled = True
                            Debug.WriteLine("Disposed dragOpenTimer")
                        Else
                            _dragOpenTimer = New DispatcherTimer()
                            AddHandler _dragOpenTimer.Tick,
                                Sub(s2 As Object, e As EventArgs)
                                    If Not CType(_dragOpenTimer.Tag, Folder).IsExpanded Then
                                        Debug.WriteLine("dragOpenTimer: expanding folder")
                                        CType(_dragOpenTimer.Tag, Folder).IsExpanded = True
                                    Else
                                        Debug.WriteLine("dragOpenTimer: folder is expanded already")
                                    End If
                                End Sub
                            _dragOpenTimer.Tag = overItem
                            _dragOpenTimer.Interval = TimeSpan.FromMilliseconds(2000)
                            _dragOpenTimer.IsEnabled = True
                        End If
                    End If
                Else
                    If Not _dragOpenTimer Is Nothing Then
                        _dragOpenTimer.IsEnabled = False
                    End If
                End If

                ' if we're not inserting not inserting and we're not over a dummy folder such as a separator...
                If _dragInsertParent Is Nothing AndAlso Not TypeOf overItem Is DummyFolder Then
                    _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
                    If _lastOverItem Is Nothing _
                        OrElse Not _lastOverItem.Equals(overItem) _
                        OrElse _insertIndex <> insertIndex Then
                        _lastOverItem = overItem

                        Dim dropTarget As IDropTarget = Nothing

                        ' first check if we're not trying to drop on ourselves or our parent
                        Dim isOurSelvesOrParent As Boolean
                        If Not _files Is Nothing Then
                            isOurSelvesOrParent = _files.Exists(Function(f) f.Pidl.Equals(overItem.Pidl))
                            If Not isOurSelvesOrParent Then
                                For Each file In _files
                                    isOurSelvesOrParent = Not file.LogicalParent Is Nothing _
                                            AndAlso file.LogicalParent.Pidl.Equals(overItem.Pidl)
                                    If isOurSelvesOrParent Then Exit For
                                Next
                            End If
                        End If
                        If Not _fileNameList Is Nothing AndAlso Not isOurSelvesOrParent Then
                            isOurSelvesOrParent = _fileNameList.ToList().Exists(Function(f) f.ToLower() = overItem.FullPath.ToLower())
                            If Not isOurSelvesOrParent Then
                                isOurSelvesOrParent = _fileNameList.ToList().Exists(Function(f) _
                                            Not IO.Path.GetDirectoryName(f) Is Nothing _
                                            AndAlso IO.Path.GetDirectoryName(f).ToLower().TrimEnd(IO.Path.DirectorySeparatorChar) _
                                                = overItem.FullPath.ToLower().TrimEnd(IO.Path.DirectorySeparatorChar))
                            End If
                        End If

                        If Not isOurSelvesOrParent Then
                            ' try get droptarget
                            If Not overItem.Parent Is Nothing Then
                                overItem.Parent.ShellFolder.GetUIObjectOf(IntPtr.Zero, 1, {overItem.Pidl.RelativePIDL}, GetType(IDropTarget).GUID, 0, dropTarget)
                            Else
                                ' desktop
                                Shell.Desktop.ShellFolder.GetUIObjectOf(IntPtr.Zero, 1, {Shell.Desktop.Pidl.AbsolutePIDL}, GetType(IDropTarget).GUID, 0, dropTarget)
                            End If
                        End If

                        If Not dropTarget Is Nothing Then
                            _treeView.SetSelectedItem(overItem)
                            If Not _lastDropTarget Is Nothing Then
                                _lastDropTarget.DragLeave()
                            End If
                            _insertIndex = insertIndex
                            Try
                                Return dropTarget.DragEnter(_dataObject, grfKeyState, ptWIN32, pdwEffect)
                            Finally
                                customizeDropDescription(overItem, grfKeyState, pdwEffect)
                                _lastDropTarget = dropTarget
                            End Try
                        Else
                            _treeView.SetSelectedItem(Nothing)
                            pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                            If Not _lastDropTarget Is Nothing Then
                                Try
                                    _lastDropTarget.DragLeave()
                                Finally
                                    If Not _lastDropTarget Is Nothing Then
                                        Marshal.ReleaseComObject(_lastDropTarget)
                                        _lastDropTarget = Nothing
                                    End If
                                End Try
                            End If
                            _insertIndex = insertIndex
                            Return HRESULT.S_OK
                        End If
                    ElseIf Not _lastDropTarget Is Nothing Then
                        _insertIndex = insertIndex
                        Try
                            Return _lastDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
                        Finally
                            customizeDropDescription(overItem, grfKeyState, pdwEffect)
                        End Try
                    Else
                        _insertIndex = insertIndex
                        pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                        _treeView.SetSelectedItem(Nothing)
                        Return HRESULT.S_OK
                    End If
                ElseIf Not _dragInsertParent Is Nothing Then
                    If Not _lastDropTarget Is Nothing Then
                        Try
                            _lastDropTarget.DragLeave()
                        Finally
                            If Not _lastDropTarget Is Nothing Then
                                Marshal.ReleaseComObject(_lastDropTarget)
                                _lastDropTarget = Nothing
                            End If
                        End Try
                    End If

                    ' we can insert, so modify user interface to reflect that
                    _treeView.SetSelectedItem(Nothing)
                    If CType(_dragInsertParent, ISupportDragInsert).DragInsertBefore(_dataObject, _files, insertIndex, overListBoxItem) = HRESULT.S_OK Then
                        pdwEffect = DROPEFFECT.DROPEFFECT_LINK
                        _treeView.PART_DragInsertIndicator.Visibility = Visibility.Visible
                    Else
                        pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                        _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
                    End If

                    _insertIndex = insertIndex
                    Return HRESULT.S_OK
                End If
            End If

            ' we're not doing anything special, cancel and unflag all
            _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
            If Not _dragOpenTimer Is Nothing Then
                _dragOpenTimer.IsEnabled = False
            End If
            _treeView.SetSelectedItem(Nothing)
            _lastOverItem = Nothing
            If Not _lastDropTarget Is Nothing Then
                Try
                    _lastDropTarget.DragLeave()
                Finally
                    If Not _lastDropTarget Is Nothing Then
                        Marshal.ReleaseComObject(_lastDropTarget)
                        _lastDropTarget = Nothing
                    End If
                End Try
            End If
            pdwEffect = DROPEFFECT.DROPEFFECT_NONE
            _insertIndex = insertIndex

            Return HRESULT.S_OK
        End Function

        Private Sub customizeDropDescription(overItem As Item, grfKeyState As MK, pdwEffect As DROPEFFECT)
            If overItem.FullPath = Shell.GetSpecialFolder(SpecialFolders.RecycleBin).FullPath And grfKeyState.HasFlag(MK.MK_SHIFT) Then
                WpfDragTargetProxy.SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_WARNING, "Delete", "")
            ElseIf pdwEffect = DROPEFFECT.DROPEFFECT_COPY AndAlso Not overItem Is Nothing Then
                WpfDragTargetProxy.SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_COPY, "Copy to %1", overItem.DisplayName)
            ElseIf pdwEffect = DROPEFFECT.DROPEFFECT_MOVE AndAlso Not overItem Is Nothing Then
                WpfDragTargetProxy.SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_MOVE, "Move to %1", overItem.DisplayName)
            ElseIf pdwEffect = DROPEFFECT.DROPEFFECT_LINK AndAlso Not overItem Is Nothing Then
                WpfDragTargetProxy.SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_LINK, "Create shortcut in %1", overItem.DisplayName)
            ElseIf pdwEffect = DROPEFFECT.DROPEFFECT_OPEN AndAlso Not overItem Is Nothing Then
                WpfDragTargetProxy.SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_COPY, "Open with %1", If(overItem.PropertiesByCanonicalName("System.FileDescription")?.Text, overItem.DisplayName))
            End If
        End Sub
    End Class
End Namespace