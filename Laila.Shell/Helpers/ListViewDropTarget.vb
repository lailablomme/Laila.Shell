Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Controls
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Windows

Namespace Helpers
    Public Class ListViewDropTarget
        Inherits BaseDropTarget

        Private _dataObject As ComTypes.IDataObject
        Private _folderView As FolderView
        Private _lastOverItem As Item
        Private _lastDropTarget As IDropTarget
        Private _scrollTimer As Timer
        Private _scrollDirection As Boolean?
        Private _prevSelectedItems As IEnumerable(Of Item)
        Private _fileNameList() As String
        Private _files As List(Of Item)

        Public Sub New(folderView As FolderView)
            _folderView = folderView
        End Sub

        Public Overrides Function DragEnter(pDataObj As ComTypes.IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            Debug.WriteLine("DragEnter")
            _dataObject = pDataObj
            _fileNameList = Clipboard.GetFileNameList(pDataObj)
            _files = ClipboardFormats.CFSTR_SHELLIDLIST.GetData(pDataObj)
            _prevSelectedItems = _folderView.SelectedItems?.ToList()
            If _prevSelectedItems Is Nothing Then _prevSelectedItems = {}
            _folderView.ActiveView.PART_ListBox.Focus()
            Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
        End Function

        Public Overrides Function DragOver(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            Return dragPoint(grfKeyState, ptWIN32, pdwEffect)
        End Function

        Public Overrides Function DragLeave() As Integer
            Debug.WriteLine("DragLeave")
            If Not _files Is Nothing Then
                For Each f In _files
                    f.Dispose()
                Next
            End If
            _folderView.ActiveView.SetSelectedItemsSoft(_prevSelectedItems)
            If Not _scrollTimer Is Nothing Then
                _scrollTimer.Dispose()
                _scrollDirection = Nothing
            End If
            _lastOverItem = Nothing
            If Not _lastDropTarget Is Nothing Then
                Try
                    Return _lastDropTarget.DragLeave()
                Finally
                    If Not _lastDropTarget Is Nothing Then
                        Marshal.ReleaseComObject(_lastDropTarget)
                        _lastDropTarget = Nothing
                    End If
                    _lastDropTarget = Nothing
                End Try
            End If
            Return 0
        End Function

        Public Overrides Function Drop(pDataObj As ComTypes.IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            If Not _files Is Nothing Then
                For Each f In _files
                    f.Dispose()
                Next
            End If
            _folderView.ActiveView.SetSelectedItemsSoft(_prevSelectedItems)
            If Not _scrollTimer Is Nothing Then
                _scrollTimer.Dispose()
                _scrollDirection = Nothing
            End If
            _lastOverItem = Nothing
            If Not _lastDropTarget Is Nothing Then
                Try
                    Dim overItem As Item = getOverItem(ptWIN32)
                    If Not overItem Is Nothing AndAlso overItem.FullPath = "shell:::{645FF040-5081-101B-9F08-00AA002F954E}" Then
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
                End Try
            End If
            Return 0
        End Function

        Private Function getOverItem(ptWIN32 As WIN32POINT) As Item
            ' translate point to listview
            Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _folderView.ActiveView.PART_ListBox)

            ' find which item we're over
            Dim overObject As IInputElement = _folderView.ActiveView.PART_ListBox.InputHitTest(pt)
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

            Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _folderView.ActiveView.PART_ListBox)
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
                                    Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_folderView.ActiveView.PART_ListBox)(0)
                                    sv.ScrollToVerticalOffset(sv.VerticalOffset - 50)
                                End Sub, Threading.DispatcherPriority.ContextIdle)
                        End Sub), Nothing, 350, 350)
                End If
            ElseIf pt.Y > _folderView.ActiveView.PART_ListBox.ActualHeight - 100 Then
                If _scrollTimer Is Nothing OrElse Not _scrollDirection.HasValue OrElse _scrollDirection <> True Then
                    _scrollDirection = True
                    If Not _scrollTimer Is Nothing Then
                        _scrollTimer.Dispose()
                    End If
                    _scrollTimer = New Timer(New TimerCallback(
                        Sub()
                            UIHelper.OnUIThread(
                                Sub()
                                    Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_folderView.ActiveView.PART_ListBox)(0)
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

            If Not overItem Is Nothing Then
                If (_lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem)) Then
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
                        Debug.WriteLine("Got dropTarget")
                        _folderView.ActiveView.SetSelectedItemsSoft({overItem}.Union(_prevSelectedItems))
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
                            customizeDropDescription(overItem, grfKeyState, pdwEffect)
                        End Try
                    Else
                        Debug.WriteLine("No dropTarget")
                        _folderView.ActiveView.SetSelectedItemsSoft(_prevSelectedItems)
                        pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                        If Not _lastDropTarget Is Nothing Then
                            Try
                                Debug.WriteLine("   Got _lastDropTarget")
                                _lastDropTarget.DragLeave()
                            Finally
                                If Not _lastDropTarget Is Nothing Then
                                    Marshal.ReleaseComObject(_lastDropTarget)
                                    _lastDropTarget = Nothing
                                End If
                            End Try
                        End If
                    End If
                ElseIf Not _lastDropTarget Is Nothing Then
                    Debug.WriteLine("DragOver")
                    Try
                        Return _lastDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
                    Finally
                        Debug.WriteLine("pdwEffect=" & pdwEffect)
                        customizeDropDescription(overItem, grfKeyState, pdwEffect)
                    End Try
                Else
                    Debug.WriteLine("DROPEFFECT_NONE")
                    pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                End If
            Else
                Debug.WriteLine("overItem=Nothing")
                _folderView.ActiveView.SetSelectedItemsSoft(_prevSelectedItems)
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
            End If

            Return HRESULT.S_OK
        End Function

        Private Sub customizeDropDescription(overItem As Item, grfKeyState As MK, pdwEffect As DROPEFFECT)
            If overItem.FullPath = "::{645FF040-5081-101B-9F08-00AA002F954E}" And grfKeyState.HasFlag(MK.MK_SHIFT) Then
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