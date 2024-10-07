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

Public Class TreeViewDropTarget
    Inherits BaseDropTarget

    Private _dataObject As ComTypes.IDataObject
    Private _treeViewModel As TreeViewModel
    Private _lastOverItem As Item
    Private _dragOpenTimer As Timer
    Private _scrollTimer As Timer
    Private _scrollDirection As Boolean?
    Private _fileList() As String

    Public Sub New(treeViewModel As TreeViewModel)
        _treeViewModel = treeViewModel
    End Sub

    Public Overrides Function DragEnter(pDataObj As ComTypes.IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        Debug.WriteLine("DragEnter")
        _dataObject = pDataObj

        _fileList = Clipboard.GetFileNameList(pDataObj)

        Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
    End Function

    Public Overrides Function DragOver(grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
    End Function

    Public Overrides Function DragLeave() As Integer
        Debug.WriteLine("DragLeave")
        If Not _prevOverTreeViewItem Is Nothing Then
            _prevOverTreeViewItem.Margin = New Thickness(0, 0, 0, 0)
            _prevOverTreeViewItem = Nothing
        End If
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If
        Return 0
    End Function

    Public Overrides Function Drop(pDataObj As ComTypes.IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If

        If Not _fileList Is Nothing Then
            Dim overItem As Item = getOverTreeViewItem(ptWIN32)?.DataContext

            If CType(pdwEffect, DROPEFFECT) <> DROPEFFECT.DROPEFFECT_NONE _
            AndAlso CType(pdwEffect, DROPEFFECT) <> DROPEFFECT.DROPEFFECT_SCROLL _
            AndAlso Not (_fileList.Count = 1 _
                AndAlso (_fileList(0) = overItem.FullPath OrElse IO.Path.GetDirectoryName(_fileList(0)) = overItem.FullPath)) Then

                If overItem.IsExecutable Then
                    overItem.Execute("""" & String.Join(""" """, _fileList) & """")
                ElseIf CType(pdwEffect, DROPEFFECT).HasFlag(DROPEFFECT.DROPEFFECT_MOVE) _
                OrElse CType(pdwEffect, DROPEFFECT).HasFlag(DROPEFFECT.DROPEFFECT_COPY) Then
                    Dim sourceItems As List(Of IShellItem) = _fileList.Select(Function(f) CType(Item.FromParsingName(f, Nothing)?._shellItem2, IShellItem)).Where(Function(i) Not i Is Nothing).ToList()
                    Dim sourcePidls As New List(Of IntPtr)()
                    Dim destPath As String = If(TypeOf overItem Is Folder, overItem.FullPath, IO.Path.GetDirectoryName(overItem.FullPath))
                    For Each item As IShellItem In sourceItems
                        Dim pidl As IntPtr = IntPtr.Zero
                        Dim punk As IntPtr = Marshal.GetIUnknownForObject(item)
                        Functions.SHGetIDListFromObject(punk, pidl)
                        sourcePidls.Add(pidl)
                        Marshal.Release(punk)
                    Next
                    Dim sourceArray As IShellItemArray
                    Functions.SHCreateShellItemArrayFromIDLists(sourcePidls.Count, sourcePidls.ToArray(), sourceArray)
                    Dim fileOperation As IFileOperation
                    Dim h As HRESULT = Functions.CoCreateInstance(Guids.CLSID_FileOperation, IntPtr.Zero, 1, GetType(IFileOperation).GUID, fileOperation)
                    Debug.WriteLine("CoCreateInstance returned " & h.ToString())
                    If CType(pdwEffect, DROPEFFECT).HasFlag(DROPEFFECT.DROPEFFECT_MOVE) Then
                        h = fileOperation.MoveItems(sourceArray, overItem._shellItem2)
                    Else
                        h = fileOperation.CopyItems(sourceArray, overItem._shellItem2)
                    End If
                    fileOperation.PerformOperations()
                    Shell.SetSelectedFolder(If(TypeOf overItem Is Folder, overItem, overItem.LogicalParent), Nothing)
                ElseIf CType(pdwEffect, DROPEFFECT).HasFlag(DROPEFFECT.DROPEFFECT_LINK) Then
                End If

                ' clean up data object
                Dim format As New FORMATETC With {
                    .cfFormat = ClipboardFormat.CF_HDROP,
                    .ptd = IntPtr.Zero,
                    .dwAspect = DVASPECT.DVASPECT_CONTENT,
                    .lindex = -1,
                    .tymed = TYMED.TYMED_HGLOBAL
                }
                Dim medium As STGMEDIUM
                _dataObject.GetData(format, medium)
                Functions.DragFinish(medium.unionmember)
            End If
        End If

        Return 0
    End Function

    Private Function getOverTreeViewItem(ptWIN32 As WIN32POINT) As TreeViewItem
        ' translate point to listview
        Dim pt As Point = UIHelper.WIN32POINTToControl(ptWIN32, _treeViewModel._view)

        ' find which item we're over
        Dim overObject As IInputElement = _treeViewModel._view.InputHitTest(pt)
        Dim overTreeViewItem As TreeViewItem
        If TypeOf overObject Is TreeViewItem Then
            overTreeViewItem = overObject
        Else
            overTreeViewItem = UIHelper.GetParentOfType(Of TreeViewItem)(overObject)
        End If
        Return overTreeViewItem
    End Function

    Private Function getDropEffect(overItem As Item) As DROPEFFECT
        If Not _fileList Is Nothing Then
            Return If(Not overItem Is Nothing _
                    AndAlso Not (_fileList.Count = 1 AndAlso (_fileList(0) = overItem.FullPath OrElse IO.Path.GetDirectoryName(_fileList(0)) = overItem.FullPath)),
                If(overItem.IsExecutable,
                    DROPEFFECT.DROPEFFECT_COPY,
                    If(Not overItem.Attributes.HasFlag(SFGAO.RDONLY),
                        DROPEFFECT.DROPEFFECT_MOVE,
                        DROPEFFECT.DROPEFFECT_NONE)),
                DROPEFFECT.DROPEFFECT_NONE)
        Else
            Return DROPEFFECT.DROPEFFECT_NONE
        End If
    End Function

    Private _prevOverTreeViewItem As TreeViewItem, _isTopOrBottom As Boolean, _offset As Double

    Private Function dragPoint(grfKeyState As UInteger, ptWIN32 As WIN32POINT, ByRef pdwEffect As UInteger) As Integer
        Dim pt As Point = UIHelper.WIN32POINTToControl(ptWIN32, _treeViewModel._view)
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
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_treeViewModel._view)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset - 50)
                            End Sub)
                    End Sub), Nothing, 350, 350)
            End If
        ElseIf pt.Y > _treeViewModel._view.ActualHeight - 100 Then
            If _scrollTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> True Then
                _scrollDirection = True
                If Not _scrollTimer Is Nothing Then
                    _scrollTimer.Dispose()
                End If
                _scrollTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_treeViewModel._view)(0)
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

        Dim overItem As Item = getOverTreeViewItem(ptWIN32)?.DataContext

        ' if we're over a folder, open it after two seconds of hovering
        If TypeOf overItem Is Folder Then
            If (_lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem)) Then
                If Not _dragOpenTimer Is Nothing Then
                    _dragOpenTimer.Dispose()
                End If

                _dragOpenTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                _treeViewModel.SetSelectedFolder(overItem)
                                CType(overItem, Folder).IsExpanded = True
                            End Sub)
                        _dragOpenTimer.Dispose()
                        _dragOpenTimer = Nothing
                    End Sub), Nothing, 1250, 0)
            End If
        Else
            If Not _dragOpenTimer Is Nothing Then
                _dragOpenTimer.Dispose()
            End If
        End If

        ' if we're over TreeView2, we might want to add pinned items
        Dim overTreeViewItem As TreeViewItem = getOverTreeViewItem(ptWIN32)
        If Not overTreeViewItem Is Nothing Then
            Dim parentTreeView As TreeView = UIHelper.GetParentOfType(Of TreeView)(overTreeViewItem)
            Const RESERVED_ITEM_SPACE As Integer = 18
            If Not parentTreeView Is Nothing AndAlso parentTreeView.Equals(_treeViewModel._view.treeView2) Then
                Dim ptOver As Point = UIHelper.WIN32POINTToControl(ptWIN32, overTreeViewItem)
                If ptOver.Y < 2 Or ptOver.Y > overTreeViewItem.ActualHeight - 2 Then
                    If _prevOverTreeViewItem Is Nothing OrElse
                    Not _prevOverTreeViewItem.Equals(overTreeViewItem) OrElse
                    _isTopOrBottom <> ptOver.Y < 2 Then

                        _isTopOrBottom = ptOver.Y < 2
                        If Not _prevOverTreeViewItem Is Nothing AndAlso
                        Not _prevOverTreeViewItem.Equals(overTreeViewItem) Then
                            _prevOverTreeViewItem.Margin = New Thickness(0, 0, 0, 0)
                            _prevOverTreeViewItem = Nothing
                        End If

                        If ptOver.Y < 2 Then
                            overTreeViewItem.Margin = New Thickness(0, RESERVED_ITEM_SPACE, 0, 0)
                            _prevOverTreeViewItem = overTreeViewItem
                        Else
                            overTreeViewItem.Margin = New Thickness(0, 0, 0, RESERVED_ITEM_SPACE)
                            _prevOverTreeViewItem = overTreeViewItem
                        End If
                    End If
                Else
                    If Not _prevOverTreeViewItem Is Nothing Then
                        _prevOverTreeViewItem.Margin = New Thickness(0, 0, 0, 0)
                        _prevOverTreeViewItem = Nothing
                    End If
                End If
            Else
                If Not _prevOverTreeViewItem Is Nothing Then
                    _prevOverTreeViewItem.Margin = New Thickness(0, 0, 0, 0)
                    _prevOverTreeViewItem = Nothing
                End If
            End If
        End If

        _lastOverItem = overItem

        pdwEffect = getDropEffect(overItem)

        Return HRESULT.Ok
    End Function
End Class
