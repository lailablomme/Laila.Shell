Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports Laila.Shell.Helpers
Imports System.Threading
Imports System.Windows
Imports System.Runtime.Serialization
Imports Laila.Shell.ViewModels
Imports System.Windows.Controls
Imports Microsoft
Imports System.Text
Imports System.IO
Imports Shell32

Public Class DropTarget
    Implements IDropTarget

    Private _dataObject As ComTypes.IDataObject
    Private _detailsListViewModel As DetailsListViewModel
    Private _lastOverItem As Item
    Private _dragOpenTimer As Timer
    Private _scrollTimer As Timer
    Private _scrollDirection As Boolean?
    Private _fileList() As String
    Private _dropTargetHelper As IDropTargetHelper

    Public Sub New(detailsListViewModel As DetailsListViewModel)
        _detailsListViewModel = detailsListViewModel
        Functions.CoCreateInstance(Guids.CLSID_DragDropHelper, IntPtr.Zero,
                &H1, GetType(IDropTargetHelper).GUID, _dropTargetHelper)
        '_dropTargetHelper = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_DragDropHelper))
    End Sub

    Public Function DragEnter(pDataObj As ComTypes.IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragEnter
        Debug.WriteLine("DragEnter")
        _dataObject = pDataObj

        _fileList = getFileList()

        Dim h As HResult = dragPoint(grfKeyState, ptWIN32, pdwEffect)
        _dropTargetHelper.DragEnter(WpfDragTargetProxy.GetHwndFromControl(_detailsListViewModel._view.listView), _dataObject, ptWIN32, pdwEffect)
        Return h
    End Function

    Public Function DragOver(grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragOver
        Dim h As HResult = dragPoint(grfKeyState, ptWIN32, pdwEffect)
        _dropTargetHelper.DragOver(ptWIN32, pdwEffect)
        Return h
    End Function

    Public Function DragLeave() As Integer Implements IDropTarget.DragLeave
        Debug.WriteLine("DragLeave")
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If
        _dropTargetHelper.DragLeave()
        Return 0
    End Function

    Public Function Drop(pDataObj As ComTypes.IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.Drop
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If

        _dropTargetHelper.Drop(_dataObject, ptWIN32, pdwEffect)

        If Not _fileList Is Nothing Then
            Dim overItem As Item = getOverItem(ptWIN32)
            If Not TypeOf overItem Is Folder Then
                overItem = _detailsListViewModel.Folder
            End If

            Dim destPath As String = If(TypeOf overItem Is Folder, overItem.FullPath, IO.Path.GetDirectoryName(overItem.FullPath))

            If CType(pdwEffect, DROPEFFECT) <> DROPEFFECT.DROPEFFECT_NONE _
            AndAlso CType(pdwEffect, DROPEFFECT) <> DROPEFFECT.DROPEFFECT_SCROLL _
            AndAlso Not (_fileList.Count = 1 _
                AndAlso (_fileList(0) = overItem.FullPath OrElse IO.Path.GetDirectoryName(_fileList(0)) = destPath)) Then

                If overItem.IsExecutable Then
                    overItem.Execute("""" & String.Join(""" """, _fileList) & """")
                ElseIf CType(pdwEffect, DROPEFFECT).HasFlag(DROPEFFECT.DROPEFFECT_MOVE) _
                OrElse CType(pdwEffect, DROPEFFECT).HasFlag(DROPEFFECT.DROPEFFECT_COPY) Then
                    Dim sourceItems As List(Of IShellItem) = _fileList.Select(Function(f) CType(Item.FromParsingName(f, Nothing, Nothing)._shellItem2, IShellItem)).ToList()
                    Dim sourcePidls As New List(Of IntPtr)()
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
                    Dim h As HResult = Functions.CoCreateInstance(Guids.CLSID_FileOperation, IntPtr.Zero, 1, GetType(IFileOperation).GUID, fileOperation)
                    Debug.WriteLine("CoCreateInstance returned" & h.ToString())
                    Dim shellItem As IShellItem = CType(overItem, Folder)._shellItem2
                    If CType(pdwEffect, DROPEFFECT).HasFlag(DROPEFFECT.DROPEFFECT_MOVE) Then
                        h = fileOperation.MoveItems(sourceArray, shellItem)
                    Else
                        h = fileOperation.CopyItems(sourceArray, shellItem)
                    End If
                    fileOperation.PerformOperations()
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

    Private Function getFileList() As String()
        ' get data from data object
        Dim format As New FORMATETC With {
            .cfFormat = Functions.RegisterClipboardFormat("Shell IDList Array"),
            .ptd = IntPtr.Zero,
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .tymed = TYMED.TYMED_HGLOBAL
        }
        If _dataObject.QueryGetData(format) = 0 Then
            Dim medium As STGMEDIUM
            _dataObject.GetData(format, medium)

            Return Pidl.GetItemsFromShellIDListArray(medium.unionmember).Select(Function(i) i.FullPath).ToArray()
        Else
            format = New FORMATETC With {
                .cfFormat = ClipboardFormat.CF_HDROP,
                .ptd = IntPtr.Zero,
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .lindex = -1,
                .tymed = TYMED.TYMED_HGLOBAL
            }
            If _dataObject.QueryGetData(format) = 0 Then
                Dim medium As STGMEDIUM
                _dataObject.GetData(format, medium)

                Dim fileCount As UInteger = Functions.DragQueryFile(medium.unionmember, UInt32.MaxValue, Nothing, 0)
                Dim fileList As New List(Of String)()
                If fileCount > 0 Then
                    For i As UInteger = 0 To fileCount - 1
                        Dim filePathBuilder As New StringBuilder(260) ' MAX_PATH
                        Functions.DragQueryFile(medium.unionmember, i, filePathBuilder, CType(filePathBuilder.Capacity, UInteger))
                        fileList.Add(filePathBuilder.ToString())
                    Next
                End If

                Return fileList.ToArray()
            End If
        End If

        Return Nothing
    End Function

    Private Function getOverItem(ptWIN32 As WIN32POINT) As Item
        ' translate point to listview
        Dim pt As Point = UIHelper.WIN32POINTToControl(ptWIN32, _detailsListViewModel._view.listView)

        ' find which item we're over
        Dim overObject As IInputElement = _detailsListViewModel._view.listView.InputHitTest(pt)
        Dim overListViewItem As ListViewItem
        If TypeOf overObject Is ListViewItem Then
            overListViewItem = overObject
        Else
            overListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(overObject)
        End If
        If Not overListViewItem Is Nothing Then
            Return overListViewItem.DataContext
        Else
            Return _detailsListViewModel.Folder
        End If
    End Function

    Private Function getDropEffect(overItem As Item) As DROPEFFECT
        Dim destPath As String = If(TypeOf overItem Is Folder, overItem.FullPath, IO.Path.GetDirectoryName(overItem.FullPath))

        If Not _fileList Is Nothing Then
            Return If(Not overItem Is Nothing _
                    AndAlso Not (_fileList.Count = 1 AndAlso (_fileList(0) = overItem.FullPath OrElse IO.Path.GetDirectoryName(_fileList(0)) = destPath)),
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

    Private Function dragPoint(grfKeyState As UInteger, ptWIN32 As WIN32POINT, ByRef pdwEffect As UInteger) As Integer
        Dim pt As Point = UIHelper.WIN32POINTToControl(ptWIN32, _detailsListViewModel._view.listView)
        If pt.Y < 100 Then
            If _scrollTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> False Then
                _scrollDirection = False
                If Not _scrollTimer Is Nothing Then
                    _scrollTimer.Dispose()
                End If
                _scrollTimer = New Timer(New TimerCallback(
                    Sub()
                        System.Windows.Application.Current.Dispatcher.Invoke(
                            Sub()
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_detailsListViewModel._view.listView)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset - 1)
                            End Sub)
                    End Sub), Nothing, 500, 500)
            End If
        ElseIf pt.Y > _detailsListViewModel._view.listView.ActualHeight - 100 Then
            If _scrollTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> True Then
                _scrollDirection = True
                If Not _scrollTimer Is Nothing Then
                    _scrollTimer.Dispose()
                End If
                _scrollTimer = New Timer(New TimerCallback(
                    Sub()
                        System.Windows.Application.Current.Dispatcher.Invoke(
                            Sub()
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_detailsListViewModel._view.listView)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset + 1)
                            End Sub)
                    End Sub), Nothing, 500, 500)
            End If
        Else
            If Not _scrollTimer Is Nothing Then
                _scrollTimer.Dispose()
                _scrollDirection = Nothing
            End If
        End If

        Dim overItem As Item = getOverItem(ptWIN32)

        ' if we're over a folder, open it after two seconds of hovering
        If TypeOf overItem Is Folder AndAlso Not overItem.Equals(_detailsListViewModel.Folder) Then
            _detailsListViewModel.SetSelectedItem(overItem)

            If (_lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem)) Then
                If Not _dragOpenTimer Is Nothing Then
                    _dragOpenTimer.Dispose()
                End If

                _dragOpenTimer = New Timer(New TimerCallback(
                    Sub()
                        System.Windows.Application.Current.Dispatcher.Invoke(
                            Sub()
                                _detailsListViewModel._view.LogicalParent = _detailsListViewModel.Folder
                                _detailsListViewModel.FolderName = overItem.FullPath
                            End Sub)
                        _dragOpenTimer.Dispose()
                        _dragOpenTimer = Nothing
                    End Sub), Nothing, 2000, 0)
            End If
        Else
            If Not _dragOpenTimer Is Nothing Then
                _dragOpenTimer.Dispose()
            End If

            If overItem.IsExecutable Then
                _detailsListViewModel.SetSelectedItem(overItem)
            Else
                _detailsListViewModel.SetSelectedItem(Nothing)
            End If
        End If

        _lastOverItem = overItem

        pdwEffect = getDropEffect(overItem)

        Return HResult.Ok
    End Function
End Class
