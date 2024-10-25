Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Controls
Imports Laila.Shell.Helpers

Public Class ListViewDropTarget
    Inherits BaseDropTarget

    Private _dataObject As IDataObject
    Private _detailsListView As DetailsListView
    Private _lastOverItem As Item
    Private _lastDropTarget As IDropTarget
    Private _dragOpenTimer As Timer
    Private _scrollTimer As Timer
    Private _scrollDirection As Boolean?

    Public Sub New(detailsListView As DetailsListView)
        _detailsListView = detailsListView
    End Sub

    Public Overrides Function DragEnter(pDataObj As IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        Debug.WriteLine("DragEnter")
        _dataObject = pDataObj
        _detailsListView.PART_ListView.Focus()
        Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
    End Function

    Public Overrides Function DragOver(grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
    End Function

    Public Overrides Function DragLeave() As Integer
        Debug.WriteLine("DragLeave")
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If
        If Not _lastDropTarget Is Nothing Then
            Try
                Return _lastDropTarget.DragLeave()
            Finally
                Marshal.ReleaseComObject(_lastDropTarget)
                _lastDropTarget = Nothing
            End Try
        End If
        Return 0
    End Function

    Public Overrides Function Drop(pDataObj As IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If
        If Not _lastDropTarget Is Nothing Then
            Try
                Return _lastDropTarget.Drop(pDataObj, grfKeyState, ptWIN32, pdwEffect)
            Finally
                Marshal.ReleaseComObject(_lastDropTarget)
                _lastDropTarget = Nothing
            End Try
        End If
        Return 0
    End Function

    Private Function getOverItem(ptWIN32 As WIN32POINT) As Item
        ' translate point to listview
        Dim pt As Point = UIHelper.WIN32POINTToControl(ptWIN32, _detailsListView.PART_ListView)

        ' find which item we're over
        Dim overObject As IInputElement = _detailsListView.PART_ListView.InputHitTest(pt)
        Dim overListViewItem As ListViewItem
        If TypeOf overObject Is ListViewItem Then
            overListViewItem = overObject
        Else
            overListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(overObject)
        End If
        If Not overListViewItem Is Nothing Then
            Return overListViewItem.DataContext
        Else
            Return _detailsListView.Folder
        End If
    End Function

    Private Function dragPoint(grfKeyState As UInteger, ptWIN32 As WIN32POINT, ByRef pdwEffect As UInteger) As Integer
        Dim pt As Point = UIHelper.WIN32POINTToControl(ptWIN32, _detailsListView.PART_ListView)
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
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_detailsListView.PART_ListView)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset - 50)
                            End Sub)
                    End Sub), Nothing, 350, 350)
            End If
        ElseIf pt.Y > _detailsListView.PART_ListView.ActualHeight - 100 Then
            If _scrollTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> True Then
                _scrollDirection = True
                If Not _scrollTimer Is Nothing Then
                    _scrollTimer.Dispose()
                End If
                _scrollTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_detailsListView.PART_ListView)(0)
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

        Dim overItem As Item = getOverItem(ptWIN32)

        ' if we're over a folder, open it after two seconds of hovering
        If TypeOf overItem Is Folder AndAlso Not overItem.Equals(_detailsListView.Folder) Then
            If (_lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem)) Then
                If Not _dragOpenTimer Is Nothing Then
                    _dragOpenTimer.Dispose()
                End If

                _dragOpenTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                If Mouse.LeftButton = MouseButtonState.Pressed _
                                    OrElse Mouse.RightButton = MouseButtonState.Pressed Then
                                    _detailsListView.Folder = overItem
                                End If
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

        If Not overItem Is Nothing Then
            If (_lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem)) Then
                _lastOverItem = overItem

                Dim dropTarget As IDropTarget, pidl As IntPtr, shellItemPtr As IntPtr, dropTargetPtr As IntPtr
                Try
                    shellItemPtr = Marshal.GetIUnknownForObject(overItem._shellItem2)
                    Functions.SHGetIDListFromObject(shellItemPtr, pidl)
                    Dim lastpidl As IntPtr = Functions.ILFindLastID(pidl)
                    overItem.Parent._shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {lastpidl}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
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

                If Not dropTarget Is Nothing Then
                    _detailsListView.SetSelectedItem(overItem)
                    If Not _lastDropTarget Is Nothing Then
                        _lastDropTarget.DragLeave()
                    End If
                    Try
                        Return dropTarget.DragEnter(_dataObject, grfKeyState, ptWIN32, pdwEffect)
                    Catch ex As Exception
                    Finally
                        _lastDropTarget = dropTarget
                    End Try
                Else
                    _detailsListView.SetSelectedItem(Nothing)
                    pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                    If Not _lastDropTarget Is Nothing Then
                        Try
                            _lastDropTarget.DragLeave()
                        Finally
                            Marshal.ReleaseComObject(_lastDropTarget)
                            _lastDropTarget = Nothing
                        End Try
                    End If
                End If
            ElseIf Not _lastDropTarget Is Nothing Then
                Return _lastDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
            End If
        Else
            _detailsListView.SetSelectedItem(Nothing)
            _lastOverItem = Nothing
            If Not _lastDropTarget Is Nothing Then
                Marshal.ReleaseComObject(_lastDropTarget)
                _lastDropTarget = Nothing
            End If
            pdwEffect = DROPEFFECT.DROPEFFECT_NONE
        End If

        Return HRESULT.Ok
    End Function
End Class
