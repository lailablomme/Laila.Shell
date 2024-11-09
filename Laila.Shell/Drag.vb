Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Public Class Drag
    Implements IDropSource

    Public Shared ICON_SIZE As Integer = 128

    Private Shared _grid As Grid
    Private Shared _dragSourceHelper As IDragSourceHelper
    Public Shared _bitmap As System.Drawing.Bitmap
    Private Shared _dataObject As IDataObject
    Public Shared _isDragging As Boolean
    Private Shared _dragImage As SHDRAGIMAGE

    Private _initializeDragImageAction As Action
    Private _copyCursor As Cursor
    Private _moveCursor As Cursor
    Private _linkCursor As Cursor
    Private _button As MK

    Public Sub New(initializeDragImageAction As Action, button As MK)
        _button = button
        Dim copyCursorStream As Stream = New MemoryStream(My.Resources.cursor_copy)
        Dim moveCursorStream As Stream = New MemoryStream(My.Resources.cursor_move)
        Dim linkCursorStream As Stream = New MemoryStream(My.Resources.cursor_link)
        _copyCursor = New Cursor(copyCursorStream)
        _moveCursor = New Cursor(moveCursorStream)
        _linkCursor = New Cursor(linkCursorStream)
        _initializeDragImageAction = initializeDragImageAction
    End Sub

    Public Shared Sub Start(items As IEnumerable(Of Item), button As MK)
        If Not _isDragging Then
            _isDragging = True

            Try
                Debug.WriteLine("Drag.Start")

                ' make a DataObject for our list of items
                Dim shellItemPtr As IntPtr, folderpidl As IntPtr
                Dim pidls(items.Count - 1) As IntPtr, lastpidl As IntPtr, pidlsPtr As IntPtr
                Try
                    Using parent = items(0).GetParent()
                        shellItemPtr = Marshal.GetIUnknownForObject(parent.ShellItem2)
                        Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)
                    End Using
                Finally
                    If Not IntPtr.Zero.Equals(shellItemPtr) Then
                        Marshal.Release(shellItemPtr)
                    End If
                End Try
                pidlsPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of IntPtr) * items.Count)
                For i = 0 To items.Count - 1
                    Try
                        shellItemPtr = Marshal.GetIUnknownForObject(items(i).ShellItem2)
                        Functions.SHGetIDListFromObject(shellItemPtr, pidls(i))
                        lastpidl = Functions.ILFindLastID(pidls(i))
                        Marshal.WriteIntPtr(IntPtr.Add(pidlsPtr, Marshal.SizeOf(Of IntPtr) * i), lastpidl)
                    Finally
                        If Not IntPtr.Zero.Equals(shellItemPtr) Then
                            Marshal.Release(shellItemPtr)
                        End If
                    End Try
                Next
                Functions.SHCreateDataObject(folderpidl, items.Count, pidlsPtr, IntPtr.Zero, GetType(IDataObject).GUID, _dataObject)

                ' for some reason we can't properly write to our DataObject before a DropTarget initializes it,
                ' and I don't know what it's doing 
                Dim initDropTarget As IDropTarget, pidl As IntPtr, dropTargetPtr As IntPtr
                Try
                    shellItemPtr = Marshal.GetIUnknownForObject(Shell.Desktop.ShellItem2)
                    Functions.SHGetIDListFromObject(shellItemPtr, pidl)
                    Dim shellFolder As IShellFolder
                    Functions.SHGetDesktopFolder(shellFolder)
                    shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {pidl}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
                    initDropTarget = Marshal.GetTypedObjectForIUnknown(dropTargetPtr, GetType(IDropTarget))
                    initDropTarget.DragEnter(_dataObject, 0, New WIN32POINT() With {.x = 0, .y = 0}, 0)
                    initDropTarget.DragLeave()
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
                    If Not initDropTarget Is Nothing Then
                        Marshal.ReleaseComObject(initDropTarget)
                    End If
                End Try

                makeDragImageObjects(items)
                InitializeDragImage()

                Dim availableDropEffects As DROPEFFECT
                availableDropEffects = availableDropEffects Or DROPEFFECT.DROPEFFECT_LINK
                If items.All(Function(i) i.Attributes.HasFlag(SFGAO.CANCOPY)) Then
                    availableDropEffects = availableDropEffects Or DROPEFFECT.DROPEFFECT_COPY
                End If
                If items.All(Function(i) i.Attributes.HasFlag(SFGAO.CANMOVE)) Then
                    availableDropEffects = availableDropEffects Or DROPEFFECT.DROPEFFECT_MOVE
                End If
                If items.All(Function(i) i.Attributes.HasFlag(SFGAO.CANLINK)) Then
                    availableDropEffects = availableDropEffects Or DROPEFFECT.DROPEFFECT_LINK
                End If

                Dim effect As Integer
                Functions.DoDragDrop(_dataObject,
                                     New Drag(Sub() If Not _dataObject Is Nothing Then InitializeDragImage(), button),
                                     availableDropEffects, effect)

                Shell._w.Content = Nothing
                Mouse.OverrideCursor = Nothing
                If Not IntPtr.Zero.Equals(folderpidl) Then
                    Marshal.FreeCoTaskMem(folderpidl)
                End If
                For i = 0 To pidls.Count - 1
                    If Not IntPtr.Zero.Equals(pidls(i)) Then
                        Marshal.FreeCoTaskMem(pidls(i))
                    End If
                Next
                Marshal.ReleaseComObject(_dataObject)
                _dataObject = Nothing
            Finally
                _isDragging = False
            End Try
        End If
    End Sub

    Private Shared Sub makeDragImageObjects(items As IEnumerable(Of Item))
        _grid = New Grid()
        Shell._w.Content = _grid
        _grid.VerticalAlignment = VerticalAlignment.Top
        _grid.HorizontalAlignment = HorizontalAlignment.Left
        _grid.Width = ICON_SIZE + 20
        _grid.Height = ICON_SIZE + 20
        _grid.Background = System.Windows.Media.Brushes.Purple
        _grid.UseLayoutRounding = True
        _grid.SnapsToDevicePixels = True

        Dim border As Border = New Border()
        border.Background = Brushes.White
        border.CornerRadius = New CornerRadius(5)
        border.BorderThickness = New Thickness(2)
        border.BorderBrush = Brushes.Black
        _grid.Children.Add(border)

        For Each item In items
            Dim img As System.Windows.Controls.Image = New System.Windows.Controls.Image()
            img.Width = ICON_SIZE
            img.Height = ICON_SIZE
            img.Source = item.Icon(ICON_SIZE)
            img.Margin = New Thickness(10)
            img.VerticalAlignment = VerticalAlignment.Top
            img.HorizontalAlignment = HorizontalAlignment.Left
            img.UseLayoutRounding = True
            img.SnapsToDevicePixels = True
            _grid.Children.Add(img)
        Next

        If items.Count > 1 Then
            Dim tbBorder As Border = New Border()
            tbBorder.VerticalAlignment = VerticalAlignment.Center
            tbBorder.HorizontalAlignment = HorizontalAlignment.Center
            tbBorder.BorderThickness = New Thickness(2)
            tbBorder.BorderBrush = Brushes.White

            Dim textBlock As TextBlock = New TextBlock()
            textBlock.Text = items.Count
            textBlock.Padding = New Thickness(5, 1, 5, 1)
            textBlock.Background = Brushes.Blue
            textBlock.Foreground = Brushes.White
            textBlock.FontWeight = FontWeights.Bold
            textBlock.FontSize = 18

            tbBorder.Child = textBlock
            _grid.Children.Add(tbBorder)
        End If

        _grid.Measure(New Size(Double.MaxValue, Double.MaxValue))
        System.Windows.Application.Current.Dispatcher.Invoke(
            Sub()
            End Sub, Threading.DispatcherPriority.ContextIdle)
        Dim renderedImage As RenderTargetBitmap = New RenderTargetBitmap(_grid.DesiredSize.Width, _grid.DesiredSize.Height, 96, 96, PixelFormats.Pbgra32)
        renderedImage.Render(_grid)
        _bitmap = ImageHelper.GetBitmap(renderedImage)
    End Sub

    Friend Shared Sub InitializeDragImage()
        If _isDragging Then
            Debug.WriteLine("InitializeDragImage")
            Functions.CoCreateInstance(Guids.CLSID_DragDropHelper, IntPtr.Zero,
                    &H1, GetType(IDragSourceHelper).GUID, _dragSourceHelper)
            _dragImage.sizeDragImage.Width = _bitmap.Width
            _dragImage.sizeDragImage.Height = _bitmap.Height
            _dragImage.ptOffset.x = (ICON_SIZE + 20) / 2
            _dragImage.ptOffset.y = ICON_SIZE + 10
            _dragImage.hbmpDragImage = _bitmap.GetHbitmap()
            _dragImage.crColorKey = System.Drawing.Color.Purple.ToArgb()
            InitializeDragSourceHelper2()
            Dim h As HRESULT = _dragSourceHelper.InitializeFromBitmap(_dragImage, _dataObject)
            'MsgBox("h=" & h.ToString())
            Debug.WriteLine("InitializeFromBitmap returned " & h.ToString())
        End If
    End Sub

    Friend Shared Sub InitializeDragSourceHelper2()
        Dim dragSourceHelper2 As IDragSourceHelper2 = _dragSourceHelper
        dragSourceHelper2.SetFlags(1)
    End Sub

    Public Function QueryContinueDrag(<[In]> <MarshalAs(UnmanagedType.Bool)> fEscapePressed As Boolean, <[In]> grfKeyState As Integer) As Integer Implements IDropSource.QueryContinueDrag
        If fEscapePressed Then
            Mouse.OverrideCursor = Nothing
            Return DragDropResult.DRAGDROP_S_CANCEL
        End If
        If (grfKeyState And _button) = 0 Then
            Mouse.OverrideCursor = Nothing
            Return DragDropResult.DRAGDROP_S_DROP
        End If
        Return DragDropResult.S_OK
    End Function

    Public Function GiveFeedback(<[In]> dwEffect As DROPEFFECT) As Integer Implements IDropSource.GiveFeedback
        Dim isShowingLayered As Integer? = getGlobalDataDWord("IsShowingLayered")
        If isShowingLayered.HasValue AndAlso isShowingLayered.Value = 1 Then
            Mouse.OverrideCursor = Nothing

            Dim wParam As IntPtr
            If dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_MOVE) Then
                wParam = DROPIMAGETYPE.DROPIMAGE_MOVE
            ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_COPY) Then
                wParam = DROPIMAGETYPE.DROPIMAGE_COPY
            ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_LINK) Then
                wParam = DROPIMAGETYPE.DROPIMAGE_LINK
            Else
                wParam = DROPIMAGETYPE.DROPIMAGE_NONE
            End If

            Dim hwnd As IntPtr? = getGlobalDataDWord("DragWindow")
            If hwnd.HasValue Then
                Functions.SendMessage(hwnd.Value, WM.USER + 2, wParam, 0)
            End If
        Else
            If dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_MOVE) Then
                Mouse.OverrideCursor = _moveCursor
            ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_COPY) Then
                Mouse.OverrideCursor = _copyCursor
            ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_LINK) Then
                Mouse.OverrideCursor = _linkCursor
            Else
                Mouse.OverrideCursor = Cursors.No
            End If
        End If

        Return DragDropResult.S_OK
    End Function

    Private Function getGlobalDataDWord(clipboardFormat As String) As Integer?
        Dim format As FORMATETC = New FORMATETC() With {
            .cfFormat = Functions.RegisterClipboardFormat(clipboardFormat),
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .ptd = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL
        }
        Dim medium As STGMEDIUM
        Dim result As Integer = _dataObject.GetData(format, medium)
        If result = 0 Then
            Return Marshal.ReadInt32(medium.unionmember)
        Else
            Return Nothing
        End If
    End Function

    Private Shared Sub setGlobalDataDWord(dataObject As IDataObject, clipboardFormat As String, val As Integer)
        Dim ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of Integer))
        Marshal.WriteInt32(ptr, val)
        Dim format As FORMATETC = New FORMATETC() With {
            .cfFormat = Functions.RegisterClipboardFormat(clipboardFormat),
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .ptd = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL
        }
        Dim medium As STGMEDIUM = New STGMEDIUM() With {
            .pUnkForRelease = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL,
            .unionmember = ptr
        }
        dataObject.SetData(format, medium, True)
    End Sub
End Class
