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

    Public Sub New(treeView As Laila.Shell.Controls.TreeView)
        _treeView = treeView
    End Sub

    Public Overrides Function DragEnter(pDataObj As IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        Debug.WriteLine("DragEnter")
        _dataObject = pDataObj
        _prevSelectedItem = _treeView.SelectedItem
        _treeView.PART_ListBox.Focus()
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
                _treeView.SetSelectedItem(_prevSelectedItem)
            End Try
        Else
            _treeView.SetSelectedItem(_prevSelectedItem)
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
                _treeView.SetSelectedItem(_prevSelectedItem)
            End Try
        Else
            _treeView.SetSelectedItem(_prevSelectedItem)
        End If
        Return 0
    End Function

    Private Function getOverItem(ptWIN32 As WIN32POINT) As Item
        ' translate point to listview
        Dim pt As Point = UIHelper.WIN32POINTToControl(ptWIN32, _treeView)

        ' find which item we're over
        Dim overObject As IInputElement = _treeView.InputHitTest(pt)
        Dim overTreeViewItem As ListBoxItem
        If TypeOf overObject Is ListBoxItem Then
            overTreeViewItem = overObject
        Else
            overTreeViewItem = UIHelper.GetParentOfType(Of ListBoxItem)(overObject)
        End If
        Return If(Not overTreeViewItem Is Nothing AndAlso Not TypeOf overTreeViewItem.DataContext Is SeparatorFolder,
            overTreeViewItem.DataContext, Nothing)
    End Function

    Private Function dragPoint(grfKeyState As UInteger, ptWIN32 As WIN32POINT, ByRef pdwEffect As UInteger) As Integer
        Dim pt As Point = UIHelper.WIN32POINTToControl(ptWIN32, _treeView)
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

        Dim overItem As Item = getOverItem(ptWIN32)

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
                                If Mouse.LeftButton = MouseButtonState.Pressed _
                                    OrElse Mouse.RightButton = MouseButtonState.Pressed Then
                                    _treeView.SetSelectedFolder(overItem)
                                    CType(overItem, Folder).IsExpanded = True
                                    _prevSelectedItem = overItem
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

                If Not dropTarget Is Nothing Then
                    _treeView.SetSelectedItem(overItem)
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
                    _treeView.SetSelectedItem(Nothing)
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
            _treeView.SetSelectedItem(Nothing)
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
