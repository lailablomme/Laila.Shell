﻿Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports System.Windows.Input
Imports Laila.Shell.Controls
Imports Laila.Shell.Helpers

Public Class ListViewDropTarget
    Inherits BaseDropTarget

    Private _dataObject As IDataObject
    Private _folderView As FolderView
    Private _lastOverItem As Item
    Private _lastDropTarget As IDropTarget
    Private _dragOpenTimer As Timer
    Private _scrollTimer As Timer
    Private _scrollDirection As Boolean?

    Public Sub New(folderView As FolderView)
        _folderView = folderView
    End Sub

    Public Overrides Function DragEnter(pDataObj As IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        Debug.WriteLine("DragEnter")
        _dataObject = pDataObj
        _folderView.ActiveView.PART_ListView.Focus()
        Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
    End Function

    Public Overrides Function DragOver(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
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
        _lastOverItem = Nothing
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

    Public Overrides Function Drop(pDataObj As IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
        If Not _dragOpenTimer Is Nothing Then
            _dragOpenTimer.Dispose()
        End If
        If Not _scrollTimer Is Nothing Then
            _scrollTimer.Dispose()
            _scrollDirection = Nothing
        End If
        _lastOverItem = Nothing
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
        Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _folderView.ActiveView.PART_ListView)

        ' find which item we're over
        Dim overObject As IInputElement = _folderView.ActiveView.PART_ListView.InputHitTest(pt)
        Dim overListViewItem As ListViewItem
        If TypeOf overObject Is ListViewItem Then
            overListViewItem = overObject
        Else
            overListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(overObject)
        End If
        If Not overListViewItem Is Nothing Then
            Return overListViewItem.DataContext
        Else
            Return _folderView.Folder
        End If
    End Function

    Private Function dragPoint(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As UInteger) As Integer
        Debug.WriteLine("dragPoint")

        Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _folderView.ActiveView.PART_ListView)
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
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_folderView.ActiveView.PART_ListView)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset - 50)
                            End Sub)
                    End Sub), Nothing, 350, 350)
            End If
        ElseIf pt.Y > _folderView.ActiveView.PART_ListView.ActualHeight - 100 Then
            If _scrollTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> True Then
                _scrollDirection = True
                If Not _scrollTimer Is Nothing Then
                    _scrollTimer.Dispose()
                End If
                _scrollTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_folderView.ActiveView.PART_ListView)(0)
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
        If TypeOf overItem Is Folder AndAlso Not overItem.Equals(_folderView.Folder) Then
            If (_lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem)) Then
                If Not _dragOpenTimer Is Nothing Then
                    _dragOpenTimer.Dispose()
                End If

                _dragOpenTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                _folderView.Folder = overItem
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
                    shellItemPtr = Marshal.GetIUnknownForObject(overItem.ShellItem2)
                    Functions.SHGetIDListFromObject(shellItemPtr, pidl)
                    If Not overItem.Parent Is Nothing Then
                        Dim lastpidl As IntPtr = Functions.ILFindLastID(pidl), shellFolder As IShellFolder = overItem.Parent.ShellFolder
                        Try
                            shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {lastpidl}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
                        Finally
                            If Not shellFolder Is Nothing Then
                                Marshal.ReleaseComObject(shellFolder)
                            End If
                        End Try
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
                    Debug.WriteLine("Got dropTarget")
                    _folderView.ActiveView.SetSelectedItem(overItem)
                    If Not _lastDropTarget Is Nothing Then
                        Debug.WriteLine("      Got _lastDropTarget")
                        _lastDropTarget.DragLeave()
                    Else
                        Debug.WriteLine("      No _lastDropTarget")
                    End If
                    Try
                        Return dropTarget.DragEnter(_dataObject, grfKeyState, ptWIN32, pdwEffect)
                    Finally
                        _lastDropTarget = dropTarget
                    End Try
                Else
                    Debug.WriteLine("No dropTarget")
                    _folderView.ActiveView.SetSelectedItem(Nothing)
                    pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                    If Not _lastDropTarget Is Nothing Then
                        Try
                            Debug.WriteLine("   Got _lastDropTarget")
                            _lastDropTarget.DragLeave()
                        Finally
                            Marshal.ReleaseComObject(_lastDropTarget)
                            _lastDropTarget = Nothing
                        End Try
                    End If
                End If
            ElseIf Not _lastDropTarget Is Nothing Then
                Debug.WriteLine("DragOver")
                Return _lastDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
            Else
                Debug.WriteLine("DROPEFFECT_NONE")
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
            End If
        Else
            Debug.WriteLine("overItem=Nothing")
            _folderView.ActiveView.SetSelectedItem(Nothing)
            _lastOverItem = Nothing
            If Not _lastDropTarget Is Nothing Then
                Try
                    _lastDropTarget.DragLeave()
                Finally
                    Marshal.ReleaseComObject(_lastDropTarget)
                    _lastDropTarget = Nothing
                End Try
            End If
            pdwEffect = DROPEFFECT.DROPEFFECT_NONE
        End If

        Return HRESULT.Ok
    End Function
End Class
