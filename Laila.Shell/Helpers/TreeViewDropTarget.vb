Imports Laila.Shell.Helpers
Imports Laila.Shell.ViewModels
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Controls
Imports System.Windows.Media.Imaging
Imports System.Windows.Media
Imports Laila.Shell.Controls
Imports System.Windows.Input

Public Class TreeViewDropTarget
    Inherits BaseDropTarget

    Private _dataObject As IDataObject
    Private _treeView As Laila.Shell.Controls.TreeView
    Private _lastOverItem As Item
    Private _lastDropTarget As IDropTarget
    Private _dragOpenTimer As Timer
    Private _scrollTimer As Timer
    Private _scrollDirection As Boolean?
    Private _prevSelectedItem As Item
    Private _newPinnedIndex As Long = -2
    Private _fileNameList() As String

    Public Sub New(treeView As Laila.Shell.Controls.TreeView)
        _treeView = treeView
    End Sub

    Public Overrides Function DragEnter(pDataObj As IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        _dataObject = pDataObj
        _fileNameList = Clipboard.GetFileNameList(pDataObj)
        _prevSelectedItem = _treeView.SelectedItem
        _treeView.PART_ListBox.Focus()
        Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
    End Function

    Public Overrides Function DragOver(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
    End Function

    Public Overrides Function DragLeave() As Integer
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If
        _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
        _newPinnedIndex = -2
        _lastOverItem = Nothing
        If Not _lastDropTarget Is Nothing Then
            Try
                Return _lastDropTarget.DragLeave()
            Finally
                Marshal.ReleaseComObject(_lastDropTarget)
                _lastDropTarget = Nothing
                _treeView.SetSelectedItem(_prevSelectedItem)
            End Try
        Else
            _treeView.SetSelectedItem(_prevSelectedItem)
        End If
        Return 0
    End Function

    Public Overrides Function Drop(pDataObj As IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If
        If _newPinnedIndex <> -2 Then
            For Each fileName In _fileNameList
                Dim unpinnedIndex As Integer = PinnedItems.UnpinItem(fileName)
                If unpinnedIndex <> -1 AndAlso unpinnedIndex < _newPinnedIndex Then
                    If _newPinnedIndex <> -1 Then _newPinnedIndex -= 1
                End If
            Next
            For Each fileName In _fileNameList
                PinnedItems.PinItem(fileName, _newPinnedIndex)
                If _newPinnedIndex <> -1 Then _newPinnedIndex += 1
            Next
            _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
            _newPinnedIndex = -2
            _prevSelectedItem = _treeView.Items.FirstOrDefault(Function(i) _
                i.TreeRootIndex >= Controls.TreeView.TreeRootSection.PINNED _
                AndAlso i.TreeRootIndex < Controls.TreeView.TreeRootSection.FREQUENT _
                AndAlso i.FullPath = _fileNameList(0))
        End If
        _lastOverItem = Nothing
        If Not _lastDropTarget Is Nothing Then
            Try
                Return _lastDropTarget.Drop(pDataObj, grfKeyState, ptWIN32, pdwEffect)
            Finally
                Marshal.ReleaseComObject(_lastDropTarget)
                _lastDropTarget = Nothing
                _treeView.SetSelectedItem(_prevSelectedItem)
                If TypeOf _prevSelectedItem Is Folder Then _treeView.SetSelectedFolder(_prevSelectedItem)
            End Try
        Else
            _treeView.SetSelectedItem(_prevSelectedItem)
            If TypeOf _prevSelectedItem Is Folder Then _treeView.SetSelectedFolder(_prevSelectedItem)
        End If
        Return 0
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
        Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _treeView)
        If pt.Y < 100 Then
            If _scrollTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> False Then
                _scrollDirection = False
                If Not _scrollTimer Is Nothing Then
                    _scrollTimer.Dispose()
                End If
                _scrollTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_treeView)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset - 50)
                            End Sub)
                    End Sub), Nothing, 350, 350)
            End If
        ElseIf pt.Y > _treeView.ActualHeight - 100 Then
            If _scrollTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> True Then
                _scrollDirection = True
                If Not _scrollTimer Is Nothing Then
                    _scrollTimer.Dispose()
                End If
                _scrollTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_treeView)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset + 50)
                            End Sub)
                    End Sub), Nothing, 350, 350)
            End If
        Else
            If Not _scrollTimer Is Nothing Then
                _scrollTimer.Dispose()
                _scrollDirection = Nothing
            End If
        End If

        Dim overListBoxItem As ListBoxItem = getOverListBoxItem(ptWIN32)
        Dim overItem As Item = overListBoxItem?.DataContext
        Dim newPinnedIndex As Long = -2

        If Not overItem Is Nothing Then
            If _fileNameList.Count > 0 _
                AndAlso overItem.TreeRootIndex >= Controls.TreeView.TreeRootSection.PINNED - 1 _
                AndAlso overItem.TreeRootIndex <= Controls.TreeView.TreeRootSection.FREQUENT Then
                Dim ptItem As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, overListBoxItem)
                If (ptItem.Y <= 3 AndAlso overItem.TreeRootIndex >= Controls.TreeView.TreeRootSection.PINNED _
                    AndAlso Not TypeOf overItem Is SeparatorFolder) _
                    OrElse TypeOf overItem Is PinnedAndFrequentPlaceholderFolder Then
                    _treeView.PART_DragInsertIndicator.Visibility = Visibility.Visible
                    _treeView.PART_DragInsertIndicator.Margin =
                        New Thickness(0, _treeView.PointFromScreen(overListBoxItem.PointToScreen(
                            New Point(0, 0))).Y - 3, 0, 0)
                    If overItem.TreeRootIndex >= Controls.TreeView.TreeRootSection.FREQUENT Then
                        Debug.WriteLine("* Over first frequent item")
                        newPinnedIndex = -1
                    Else
                        Debug.WriteLine("* Over pinned item")
                        newPinnedIndex = overItem.TreeRootIndex - Controls.TreeView.TreeRootSection.PINNED
                    End If
                ElseIf ptItem.Y >= overListBoxItem.ActualHeight - 3 _
                    AndAlso overItem.TreeRootIndex < Controls.TreeView.TreeRootSection.FREQUENT Then
                    Debug.WriteLine("* Over lower part &" & overListBoxItem.ActualHeight)
                    _treeView.PART_DragInsertIndicator.Visibility = Visibility.Visible
                    _treeView.PART_DragInsertIndicator.Margin =
                        New Thickness(0, _treeView.PointFromScreen(overListBoxItem.PointToScreen(
                            New Point(0, overListBoxItem.ActualHeight))).Y - 3, 0, 0)
                    newPinnedIndex = overItem.TreeRootIndex - Controls.TreeView.TreeRootSection.PINNED + 1
                    If newPinnedIndex < 0 Then _newPinnedIndex = 0
                Else
                    Debug.WriteLine("* Full over item")
                    _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
                    newPinnedIndex = -2
                End If
            Else
                Debug.WriteLine("* Not over pinned item")
                _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
                newPinnedIndex = -2
            End If

            Debug.WriteLine("_newPinnedIndex=" & newPinnedIndex & "   overItem.TreeRootIndex=" & overItem.TreeRootIndex)

            ' if we're over a folder, open it after two seconds of hovering
            If TypeOf overItem Is Folder AndAlso newPinnedIndex = -2 Then
                If (_lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem)) Then
                    If Not _dragOpenTimer Is Nothing Then
                        _dragOpenTimer.Dispose()
                    End If

                    _dragOpenTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                _treeView.SetSelectedFolder(overItem)
                                CType(overItem, Folder).IsExpanded = True
                                _prevSelectedItem = overItem
                            End Sub)
                        _dragOpenTimer.Dispose()
                        _dragOpenTimer = Nothing
                    End Sub), Nothing, 2000, 0)
                End If
            Else
                If Not _dragOpenTimer Is Nothing Then
                    _dragOpenTimer.Dispose()
                End If
            End If

            If newPinnedIndex = -2 AndAlso Not TypeOf overItem Is SeparatorFolder Then
                If _lastOverItem Is Nothing _
                    OrElse Not _lastOverItem.Equals(overItem) _
                    OrElse _newPinnedIndex <> newPinnedIndex Then
                    _lastOverItem = overItem

                    Dim dropTarget As IDropTarget, pidl As IntPtr, shellItemPtr As IntPtr, dropTargetPtr As IntPtr
                    Try
                        shellItemPtr = Marshal.GetIUnknownForObject(overItem._shellItem2)
                        Functions.SHGetIDListFromObject(shellItemPtr, pidl)
                        If Not overItem.Parent Is Nothing Then
                            Dim lastpidl As IntPtr = Functions.ILFindLastID(pidl)
                            overItem.Parent._shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {lastpidl}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
                        Else
                            Dim shellFolder As IShellFolder
                            Functions.SHGetDesktopFolder(shellFolder)
                            ' desktop
                            shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {pidl}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
                        End If
                        If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                            dropTarget = Marshal.GetTypedObjectForIUnknown(dropTargetPtr, GetType(IDropTarget))
                        Else
                            dropTarget = Nothing
                        End If
                    Finally
                        If Not IntPtr.Zero.Equals(shellItemPtr) Then
                            Marshal.Release(shellItemPtr)
                        End If
                        If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                            Marshal.Release(dropTargetPtr)
                        End If
                        If Not IntPtr.Zero.Equals(pidl) Then
                            Marshal.FreeCoTaskMem(pidl)
                        End If
                    End Try
                    'Debug.WriteLine("got dropTarget = " & If(dropTarget Is Nothing, "nothing", "something"))
                    If Not dropTarget Is Nothing Then
                        _treeView.SetSelectedItem(overItem)
                        If Not _lastDropTarget Is Nothing Then
                            'Debug.WriteLine("_lastDropTarget.DragLeave()   newPinndedIndex=" & newPinnedIndex)
                            _lastDropTarget.DragLeave()
                        End If
                        'If Drag.GetHasGlobalData(_dataObject, "DropDescription") AndAlso Not WpfDragTargetProxy._isDropDescriptionSet AndAlso Drag._isDragging Then
                        '    WpfDragTargetProxy.SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_INVALID, Nothing, Nothing)
                        'End If
                        _newPinnedIndex = newPinnedIndex
                        Try
                            'Debug.WriteLine("dropTarget.DragEnter()   newPinndedIndex=" & newPinnedIndex)
                            Return dropTarget.DragEnter(_dataObject, grfKeyState, ptWIN32, pdwEffect)
                        Finally
                            _lastDropTarget = dropTarget
                        End Try
                    Else
                        _treeView.SetSelectedItem(Nothing)
                        pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                        If Not _lastDropTarget Is Nothing Then
                            Try
                                'Debug.WriteLine("_lastDropTarget.DragLeave2()   newPinndedIndex=" & newPinnedIndex)
                                _lastDropTarget.DragLeave()
                            Finally
                                Marshal.ReleaseComObject(_lastDropTarget)
                                _lastDropTarget = Nothing
                            End Try
                        End If
                        _newPinnedIndex = newPinnedIndex
                        Return HRESULT.Ok
                    End If
                ElseIf Not _lastDropTarget Is Nothing Then
                    _newPinnedIndex = newPinnedIndex
                    'Debug.WriteLine("_lastDropTarget.DragOver()   newPinndedIndex=" & newPinnedIndex)
                    'If Drag.GetHasGlobalData(_dataObject, "DropDescription") AndAlso Not WpfDragTargetProxy._isDropDescriptionSet AndAlso Drag._isDragging Then
                    '    WpfDragTargetProxy.SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_INVALID, Nothing, Nothing)
                    'End If
                    Return _lastDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
                Else
                    'Debug.WriteLine("did nothing")
                    _newPinnedIndex = newPinnedIndex
                    pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                    _treeView.SetSelectedItem(Nothing)
                    Return HRESULT.Ok
                End If
            ElseIf newPinnedIndex <> -2 Then
                pdwEffect = DROPEFFECT.DROPEFFECT_LINK
                _treeView.SetSelectedItem(Nothing)
                If Not _lastDropTarget Is Nothing Then
                    Try
                        'Debug.WriteLine("_lastDropTarget.DragLeave()   newPinndedIndex=" & newPinnedIndex)
                        _lastDropTarget.DragLeave()
                    Finally
                        Marshal.ReleaseComObject(_lastDropTarget)
                        _lastDropTarget = Nothing
                    End Try
                End If
                WpfDragTargetProxy.SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_LINK, "Pin to %1", "Quick access")
                _newPinnedIndex = newPinnedIndex
                Return HRESULT.Ok
            End If
        End If

        _treeView.PART_DragInsertIndicator.Visibility = Visibility.Collapsed
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        _treeView.SetSelectedItem(Nothing)
        _lastOverItem = Nothing
        If Not _lastDropTarget Is Nothing Then
            Try
                'Debug.WriteLine("_lastDropTarget.DragLeave()   newPinndedIndex=" & newPinnedIndex)
                _lastDropTarget.DragLeave()
            Finally
                Marshal.ReleaseComObject(_lastDropTarget)
                _lastDropTarget = Nothing
            End Try
        End If
        pdwEffect = DROPEFFECT.DROPEFFECT_NONE
        _newPinnedIndex = newPinnedIndex
        'Debug.WriteLine("_newPinnedIndex=NO ITEM")

        Return HRESULT.Ok
    End Function
End Class
