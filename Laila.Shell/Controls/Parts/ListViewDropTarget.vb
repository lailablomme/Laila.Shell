Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Laila.Shell.Controls
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Windows

Namespace Controls.Parts
    Public Class ListViewDropTarget
        Inherits BaseDropTarget

        Private _dataObject As ComTypes.IDataObject
        Private _folderView As FolderView
        Private _lastOverItem As Item
        Private _lastDropTarget As IDropTarget
        Private _scrollUpTimer As DispatcherTimer
        Private _scrollDownTimer As DispatcherTimer
        Private _scrollLeftTimer As DispatcherTimer
        Private _scrollRightTimer As DispatcherTimer
        Private _scrollDirectionY As Boolean?
        Private _scrollDirectionX As Boolean?
        Private _prevSelectedItems As IEnumerable(Of Item)
        Private _fileNameList() As String
        Private _files As List(Of Item)
        Private _dragInsertParent As ISupportDragInsert = Nothing
        Private _insertIndex As Long = -2

        Public Sub New(folderView As FolderView)
            _folderView = folderView
        End Sub

        Public Overrides Function DragEnter(pDataObj As ComTypes.IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            Debug.WriteLine("DragEnter")
            _dataObject = pDataObj
            _fileNameList = Clipboard.GetFileNameList(pDataObj)
            _files = ClipboardFormats.CFSTR_SHELLIDLIST.GetData(pDataObj, True)
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
                    f.LogicalParent.Dispose()
                    f.Dispose()
                Next
            End If
            _folderView.ActiveView.SetSelectedItemsSoft(_prevSelectedItems)
            If Not _scrollUpTimer Is Nothing Then
                _scrollUpTimer.IsEnabled = False
            End If
            If Not _scrollDownTimer Is Nothing Then
                _scrollDownTimer.IsEnabled = False
            End If
            If Not _scrollLeftTimer Is Nothing Then
                _scrollLeftTimer.IsEnabled = False
            End If
            If Not _scrollRightTimer Is Nothing Then
                _scrollRightTimer.IsEnabled = False
            End If
            _folderView.ActiveView.DragViewStrategy?.SetDragInsertIndicator(Nothing, Nothing, Visibility.Collapsed, -1)
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
            _folderView.ActiveView.SetSelectedItemsSoft(_prevSelectedItems)
            If Not _scrollUpTimer Is Nothing Then
                _scrollUpTimer.IsEnabled = False
            End If
            If Not _scrollDownTimer Is Nothing Then
                _scrollDownTimer.IsEnabled = False
            End If
            If Not _scrollLeftTimer Is Nothing Then
                _scrollLeftTimer.IsEnabled = False
            End If
            If Not _scrollRightTimer Is Nothing Then
                _scrollRightTimer.IsEnabled = False
            End If
            _lastOverItem = Nothing

            ' we're inserting
            If Not _dragInsertParent Is Nothing Then
                CType(_dragInsertParent, ISupportDragInsert).Drop(_dataObject, _files, _insertIndex)
                _folderView.ActiveView.DragViewStrategy?.SetDragInsertIndicator(Nothing, Nothing, Visibility.Collapsed, -1)
            End If

            If Not _lastDropTarget Is Nothing Then
                Try
                    Dim overItem As Item = _folderView.ActiveView.DragViewStrategy.GetOverListBoxItem(ptWIN32)?.DataContext
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

            If Not _files Is Nothing Then
                For Each f In _files
                    f.LogicalParent.Dispose()
                    f.Dispose()
                Next
            End If

            Return 0
        End Function

        Private Function dragPoint(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As UInteger) As Integer
            If Not If(_fileNameList?.Count, 0) > 0 Then
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HRESULT.S_OK
            End If

            ' scroll up and down while dragging?
            Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _folderView)
            If pt.Y < 100 Then
                If _scrollUpTimer Is Nothing OrElse Not _scrollDirectionY.HasValue OrElse _scrollDirectionY <> False Then
                    _scrollDirectionY = False
                    If Not _scrollDownTimer Is Nothing Then _scrollDownTimer.IsEnabled = False
                    If Not _scrollUpTimer Is Nothing Then
                        _scrollUpTimer.IsEnabled = True
                    Else
                        _scrollUpTimer = New DispatcherTimer()
                        AddHandler _scrollUpTimer.Tick,
                            Sub(s2 As Object, e As EventArgs)
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_folderView.ActiveView.PART_ListBox)(0)
                                sv.ScrollToVerticalOffset(sv.VerticalOffset - 50)
                            End Sub
                        _scrollUpTimer.Interval = TimeSpan.FromMilliseconds(350)
                        _scrollUpTimer.IsEnabled = True
                    End If
                End If
            ElseIf pt.Y > _folderView.ActualHeight - 100 Then
                If _scrollDownTimer Is Nothing OrElse Not _scrollDirectionY.HasValue OrElse _scrollDirectionY <> True Then
                    _scrollDirectionY = True
                    If Not _scrollUpTimer Is Nothing Then _scrollUpTimer.IsEnabled = False
                    If Not _scrollDownTimer Is Nothing Then
                        _scrollDownTimer.IsEnabled = True
                    Else
                        _scrollDownTimer = New DispatcherTimer()
                        AddHandler _scrollDownTimer.Tick,
                            Sub(s2 As Object, e As EventArgs)
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_folderView.ActiveView.PART_ListBox)(0)
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
                _scrollDirectionY = Nothing
            End If
            ' scroll left and right while dragging?
            If pt.X < 100 Then
                If _scrollLeftTimer Is Nothing OrElse Not _scrollDirectionX.HasValue OrElse _scrollDirectionX <> False Then
                    _scrollDirectionX = False
                    If Not _scrollRightTimer Is Nothing Then _scrollRightTimer.IsEnabled = False
                    If Not _scrollLeftTimer Is Nothing Then
                        _scrollLeftTimer.IsEnabled = True
                    Else
                        _scrollLeftTimer = New DispatcherTimer()
                        AddHandler _scrollLeftTimer.Tick,
                            Sub(s2 As Object, e As EventArgs)
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_folderView.ActiveView.PART_ListBox)(0)
                                sv.ScrollToHorizontalOffset(sv.HorizontalOffset - 25)
                            End Sub
                        _scrollLeftTimer.Interval = TimeSpan.FromMilliseconds(175)
                        _scrollLeftTimer.IsEnabled = True
                    End If
                End If
            ElseIf pt.X > _folderView.ActualWidth - 100 Then
                If _scrollRightTimer Is Nothing OrElse Not _scrollDirectionX.HasValue OrElse _scrollDirectionX <> True Then
                    _scrollDirectionX = True
                    If Not _scrollLeftTimer Is Nothing Then _scrollLeftTimer.IsEnabled = False
                    If Not _scrollRightTimer Is Nothing Then
                        _scrollRightTimer.IsEnabled = True
                    Else
                        _scrollRightTimer = New DispatcherTimer()
                        AddHandler _scrollRightTimer.Tick,
                            Sub(s2 As Object, e As EventArgs)
                                Dim sv As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(_folderView.ActiveView.PART_ListBox)(0)
                                sv.ScrollToHorizontalOffset(sv.HorizontalOffset + 25)
                                Debug.WriteLine("scroll right")
                            End Sub
                        _scrollRightTimer.Interval = TimeSpan.FromMilliseconds(175)
                        _scrollRightTimer.IsEnabled = True
                    End If
                End If
            Else
                If Not _scrollLeftTimer Is Nothing Then
                    _scrollLeftTimer.IsEnabled = False
                End If
                If Not _scrollRightTimer Is Nothing Then
                    _scrollRightTimer.IsEnabled = False
                End If
                _scrollDirectionX = Nothing
            End If

            Dim overListBoxItem As ListBoxItem = _folderView.ActiveView.DragViewStrategy.GetOverListBoxItem(ptWIN32)
            Dim overItem As Item = overListBoxItem?.DataContext

            If overItem Is Nothing Then overItem = _folderView.Folder

            Dim insertIndex As Long = -1

            If Not overItem Is Nothing Then
                If Not overItem?.Equals(_folderView.Folder) AndAlso Not _folderView.ActiveView?.DragViewStrategy Is Nothing Then
                    _folderView.ActiveView?.DragViewStrategy.GetInsertIndex(ptWIN32, overListBoxItem, overItem, _dragInsertParent, insertIndex)
                Else
                    _dragInsertParent = Nothing
                End If

                If _dragInsertParent Is Nothing Then
                    If _lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem) OrElse _insertIndex <> insertIndex Then
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
                                _folderView.ActiveView.DragViewStrategy?.SetDragInsertIndicator(Nothing, Nothing, Visibility.Collapsed, -1)
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
                        _folderView.ActiveView.DragViewStrategy?.SetDragInsertIndicator(Nothing, Nothing, Visibility.Collapsed, -1)
                    ElseIf Not _lastDropTarget Is Nothing Then
                        Debug.WriteLine("DragOver")
                        Try
                            If Not _folderView.ActiveView.SelectedItems.Contains(overItem) Then
                                _folderView.ActiveView.SetSelectedItemsSoft({overItem}.Union(_prevSelectedItems))
                            End If
                            _lastDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
                        Finally
                            Debug.WriteLine("pdwEffect=" & pdwEffect)
                            customizeDropDescription(overItem, grfKeyState, pdwEffect)
                        End Try
                    Else
                        Debug.WriteLine("DROPEFFECT_NONE")
                        pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                    End If
                Else
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
                    If CType(_dragInsertParent, ISupportDragInsert).DragInsertBefore(_dataObject, _files, insertIndex, overListBoxItem) = HRESULT.S_OK Then
                        _folderView.ActiveView.SetSelectedItemsSoft(Nothing)
                        pdwEffect = DROPEFFECT.DROPEFFECT_LINK
                        _folderView.ActiveView.DragViewStrategy?.SetDragInsertIndicator(overListBoxItem, overItem, Visibility.Visible, insertIndex)
                    Else
                        _folderView.ActiveView.SetSelectedItemsSoft(_prevSelectedItems)
                        pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                        _folderView.ActiveView.DragViewStrategy?.SetDragInsertIndicator(Nothing, Nothing, Visibility.Collapsed, -1)
                    End If

                    _insertIndex = insertIndex
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
                _folderView.ActiveView.DragViewStrategy?.SetDragInsertIndicator(Nothing, Nothing, Visibility.Collapsed, -1)
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