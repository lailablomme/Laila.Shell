Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers

Public Class DragDrop
    Implements IDropSource

    Private Const MAX_COL As Integer = 5
    Private Const MAX_ROW As Integer = 5
    Private Const GMEM_MOVEABLE As Integer = &H2
    Private Const GMEM_ZEROINIT As Integer = &H40
    Private Const GMEM_SHARE As Integer = &H2000
    Private Const GHND As Integer = GMEM_MOVEABLE Or GMEM_ZEROINIT Or GMEM_SHARE

    Public Shared ICON_SIZE As Integer = 128

    Private Shared _grid As Grid
    Private Shared _dragSourceHelper As IDragSourceHelper
    Public Shared _bitmap As Bitmap
    Private Shared _dataObject As ComTypes.IDataObject
    Private Shared _isDragging As Boolean
    Private Shared _dragImage As SHDRAGIMAGE

    Private _handleDrag As Action
    Private _copyCursor As Cursor
    Private _moveCursor As Cursor
    Private _linkCursor As Cursor
    Private _button As MK

    Public Sub New(addDragImageAction As Action, button As MK)
        _button = button
        Dim copyCursorStream As Stream = New MemoryStream(My.Resources.cursor_copy)
        Dim moveCursorStream As Stream = New MemoryStream(My.Resources.cursor_move)
        Dim linkCursorStream As Stream = New MemoryStream(My.Resources.cursor_link)
        _copyCursor = New Cursor(copyCursorStream)
        _moveCursor = New Cursor(moveCursorStream)
        _linkCursor = New Cursor(linkCursorStream)
        _handleDrag = addDragImageAction
        _lastEffect = -1
    End Sub

    Public Shared Sub DoDragDrop(items As IEnumerable(Of Item), button As MK)
        If Not _isDragging Then
            _isDragging = True

            Try
                Debug.WriteLine("DoDragDrop")

                _dataObject = New DragDataObject()

                Dim format As New FORMATETC With {
                    .cfFormat = ClipboardFormat.CF_HDROP,
                    .ptd = IntPtr.Zero,
                    .dwAspect = DVASPECT.DVASPECT_CONTENT,
                    .lindex = -1,
                    .tymed = TYMED.TYMED_HGLOBAL
                }
                Dim medium As New STGMEDIUM With {
                    .tymed = TYMED.TYMED_HGLOBAL,
                    .unionmember = createCFHDrop(items.Select(Function(i) i.FullPath).ToArray()),
                    .pUnkForRelease = IntPtr.Zero
                }
                _dataObject.SetData(format, medium, False)

                format = New FORMATETC With {
                    .cfFormat = Functions.RegisterClipboardFormat("Shell ID List"),
                    .ptd = IntPtr.Zero,
                    .dwAspect = DVASPECT.DVASPECT_CONTENT,
                    .lindex = -1,
                    .tymed = TYMED.TYMED_HGLOBAL
                }
                Dim pidls As List(Of IntPtr) = New List(Of IntPtr)()
                For Each item In items
                    Dim pidl As IntPtr = IntPtr.Zero
                    Dim punk As IntPtr = Marshal.GetIUnknownForObject(items(0)._shellItem2)
                    Functions.SHGetIDListFromObject(punk, pidl)
                    pidls.Add(pidl)
                    Marshal.Release(punk)
                Next
                medium = New STGMEDIUM With {
                    .tymed = TYMED.TYMED_HGLOBAL,
                    .unionmember = createShellIDList(pidls),
                    .pUnkForRelease = IntPtr.Zero
                }

                makeDragImageObjects(items)
                handleDrag(items)

                Dim availableDropEffects As DROPEFFECT
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
                Functions.DoDragDrop(_dataObject, New DragDrop(Sub() If Not _dataObject Is Nothing Then handleDrag(items), button),
                             availableDropEffects, effect)

                Functions.ReleaseStgMedium(medium)
                Shell._w.Content = Nothing
                Mouse.OverrideCursor = Nothing
                CType(_dataObject, IDisposable).Dispose()
                _dataObject = Nothing
            Finally
                _isDragging = False
            End Try
        End If
    End Sub

    Private Shared Function createShellIDList(pidls As List(Of IntPtr)) As IntPtr
        Dim totalSize As Integer = 0
        For Each pid In pidls
            totalSize += Marshal.ReadInt16(pid) ' Get size of each PIDL
        Next

        Dim idListPtr As IntPtr = Marshal.AllocHGlobal(totalSize + 4) ' +4 for the header size
        Marshal.WriteInt32(idListPtr, totalSize) ' Write the total size at the start

        Dim offset As Integer = 4 ' Start writing after the size header
        For Each pid In pidls
            Dim pidSize As Integer = Marshal.ReadInt16(pid) ' Get size of the current PIDL
            Functions.CopyMemory(idListPtr + offset, pid, pidSize) ' Copy the PIDL data
            offset += pidSize ' Update offset for the next PIDL
        Next

        Return idListPtr ' Return pointer to the allocated ID list
    End Function

    Private Shared Function createCFHDrop(files As String()) As IntPtr
        Dim sb As StringBuilder = New StringBuilder()
        For Each file As String In files
            sb.Append(file & Chr(0))
        Next
        sb.Append(Chr(0))

        Dim strPtr As IntPtr = Marshal.StringToHGlobalUni(sb.ToString())
        Dim strSize As Integer = Functions.GlobalSize(strPtr)

        Dim hGlobal As IntPtr = Functions.GlobalAlloc(GMEM_MOVEABLE, Marshal.SizeOf(Of DROPFILES) + strSize)
        Dim pGlobal As IntPtr = Functions.GlobalLock(hGlobal)

        ' Write the DROPFILES structure
        Dim dropfiles As New DROPFILES()
        dropfiles.pFiles = CType(Marshal.SizeOf(GetType(DROPFILES)), UInteger)
        dropfiles.fWide = True ' Using Unicode (WCHAR)

        Marshal.StructureToPtr(dropfiles, pGlobal, False)
        Functions.RtlMoveMemory(IntPtr.Add(pGlobal, dropfiles.pFiles), strPtr, strSize)

        Marshal.FreeHGlobal(strPtr)
        Functions.GlobalUnlock(pGlobal)

        Return hGlobal
    End Function

    Private Shared Sub makeDragImageObjects(items As IEnumerable(Of Item))
        If items.Count > 5 Then
            ICON_SIZE = 64
        Else
            ICON_SIZE = 128
        End If

        _grid = New Grid()
        Shell._w.Content = _grid
        _grid.VerticalAlignment = VerticalAlignment.Top
        _grid.HorizontalAlignment = Windows.HorizontalAlignment.Left
        _grid.Width = MAX_ROW * ICON_SIZE
        _grid.Height = MAX_COL * ICON_SIZE
        _grid.Background = System.Windows.Media.Brushes.Purple
        _grid.UseLayoutRounding = True
        _grid.SnapsToDevicePixels = True
        'Dim brush As LinearGradientBrush = New LinearGradientBrush(Colors.White, Colors.Transparent, 45)
        'grid.OpacityMask = brush

        Dim left As Double, top As Double, count = 1
        Dim images As List(Of System.Windows.Controls.Image) = New List(Of System.Windows.Controls.Image)()
        For Each item In items
            Dim img As System.Windows.Controls.Image = New System.Windows.Controls.Image()
            img.Width = ICON_SIZE
            img.Height = ICON_SIZE
            img.Source = item.Icon(ICON_SIZE)
            img.Margin = New Windows.Thickness(left, top, 0, 0)
            img.VerticalAlignment = VerticalAlignment.Top
            img.HorizontalAlignment = Windows.HorizontalAlignment.Left
            img.UseLayoutRounding = True
            img.SnapsToDevicePixels = True

            left += ICON_SIZE
            If count Mod MAX_COL = 0 Then
                top += ICON_SIZE
                left = 0
            End If
            count += 1
            If count > MAX_COL * MAX_ROW Then
                Exit For
            End If

            images.Add(img)
        Next

        images.Reverse()
        For Each img In images
            _grid.Children.Add(img)
        Next

        _grid.Measure(New Windows.Size(ICON_SIZE * MAX_COL, ICON_SIZE * MAX_ROW))
        System.Windows.Application.Current.Dispatcher.Invoke(
            Sub()
            End Sub, Threading.DispatcherPriority.ContextIdle)
        Dim renderedImage As RenderTargetBitmap = New RenderTargetBitmap(_grid.DesiredSize.Width, _grid.DesiredSize.Height, 96, 96, PixelFormats.Pbgra32)
        renderedImage.Render(_grid)
        _bitmap = ImageHelper.GetBitmap(renderedImage)
    End Sub

    Private Shared Sub addDropDescription(dataObject As ComTypes.IDataObject, dwEffect As DROPEFFECT)
        Dim i As Integer = Functions.RegisterClipboardFormat("DropDescription")
        While i > Short.MaxValue
            i -= 65536
        End While

        Dim formatEtc As New FORMATETC With {
            .cfFormat = i,
            .ptd = IntPtr.Zero,
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .tymed = TYMED.TYMED_HGLOBAL
        }

        Dim dropDescription As DROPDESCRIPTION
        If dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_COPY) Then
            dropDescription.szMessage = "Copy here"
            dropDescription.type = DropImageType.DROPIMAGE_COPY
        ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_MOVE) Then
            dropDescription.szMessage = "Move here"
            dropDescription.type = DropImageType.DROPIMAGE_MOVE
        ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_LINK) Then
            dropDescription.szMessage = "Create shortcut here"
            dropDescription.type = DropImageType.DROPIMAGE_LINK
        Else
            dropDescription.szMessage = ""
            dropDescription.type = DropImageType.DROPIMAGE_NONE
        End If
        dropDescription.szInsert = ""

        Dim ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(DROPDESCRIPTION)))
        Marshal.StructureToPtr(dropDescription, ptr, False)

        Dim medium As System.Runtime.InteropServices.ComTypes.STGMEDIUM
        medium.pUnkForRelease = IntPtr.Zero
        medium.tymed = ComTypes.TYMED.TYMED_HGLOBAL
        medium.unionmember = ptr

        dataObject.SetData(formatEtc, medium, True)
    End Sub

    Private Shared Sub handleDrag(items As IEnumerable(Of Item))
        Debug.WriteLine("HANDLEDRAG")
        Functions.CoCreateInstance(Guids.CLSID_DragDropHelper, IntPtr.Zero,
                &H1, GetType(IDragSourceHelper).GUID, _dragSourceHelper)
        Dim dragSourceHelper2 As IDragSourceHelper2 = _dragSourceHelper
        dragSourceHelper2.SetFlags(1)
        _dragImage.sizeDragImage.Width = _bitmap.Width
        _dragImage.sizeDragImage.Height = _bitmap.Height
        _dragImage.ptOffset.x = ICON_SIZE / 2
        _dragImage.ptOffset.y = ICON_SIZE / 2
        _dragImage.hbmpDragImage = _bitmap.GetHbitmap()
        _dragImage.crColorKey = System.Drawing.Color.Purple.ToArgb()
        Debug.WriteLine("InitializeFromBitmap returned " & _dragSourceHelper.InitializeFromBitmap(_dragImage, _dataObject))
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

    Private _lastEffect As DROPEFFECT
    Public Function GiveFeedback(<[In]> dwEffect As DROPEFFECT) As Integer Implements IDropSource.GiveFeedback
        If _lastEffect <> dwEffect Then
            _lastEffect = dwEffect
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

        Debug.WriteLine(CType(dwEffect, DROPEFFECT).ToString())

        Return DragDropResult.S_OK
    End Function
End Class
