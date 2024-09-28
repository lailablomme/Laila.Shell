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

Public Class DragDrop
    Implements IDropSource

    Private Const MAX_COL As Integer = 5
    Private Const MAX_ROW As Integer = 5
    Private Const GMEM_MOVEABLE As Integer = &H2
    Private Const GMEM_ZEROINIT As Integer = &H40
    Private Const GMEM_SHARE As Integer = &H2000
    Private Const GHND As Integer = GMEM_MOVEABLE Or GMEM_ZEROINIT Or GMEM_SHARE

    Private Shared ICON_SIZE As Integer = 128

    Private Shared _grid As Grid
    Private Shared _dragSourceHelper As IDragSourceHelper
    Private Shared _bitmap As Bitmap
    Private Shared _dataObject As DragDataObject

    Private _addDragImageAction As Action
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
        _addDragImageAction = addDragImageAction
    End Sub

    Public Shared Sub DoDragDrop(items As IEnumerable(Of Item), button As MK)
        Dim dropFiles As New DROPFILES()
        dropFiles.pFiles = Marshal.SizeOf(dropFiles)
        dropFiles.fWide = False

        Dim fileNames As New StringBuilder()
        For Each item In items
            fileNames.Append(item.FullPath)
            fileNames.Append(Chr(0)) ' Null character
        Next
        fileNames.Append(Chr(0)) ' Double null character to end the list

        Dim fileNamesBytes As Byte() = Encoding.ASCII.GetBytes(fileNames.ToString())
        Dim globalPtr As IntPtr = Functions.GlobalAlloc(GHND, Marshal.SizeOf(dropFiles) + fileNamesBytes.Length)
        Dim lockPtr As IntPtr = Functions.GlobalLock(globalPtr)
        Marshal.StructureToPtr(dropFiles, lockPtr, False)
        Marshal.Copy(fileNamesBytes, 0, IntPtr.Add(lockPtr, Marshal.SizeOf(dropFiles)), fileNamesBytes.Length)
        Functions.GlobalUnlock(lockPtr)

        Dim format As New FORMATETC With {
            .cfFormat = ClipboardFormat.CF_HDROP,
            .ptd = IntPtr.Zero,
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .tymed = TYMED.TYMED_HGLOBAL
        }

        Dim medium As New STGMEDIUM With {
            .tymed = TYMED.TYMED_HGLOBAL,
            .unionmember = globalPtr,
            .pUnkForRelease = IntPtr.Zero
        }

        If Not _dataObject Is Nothing Then _dataObject.Dispose()

        _dataObject = New DragDataObject()
        _dataObject.SetData(format, medium, False)

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
        Functions.DoDragDrop(_dataObject, New DragDrop(Sub() handleDrag(items), button),
                             availableDropEffects, effect)

        Functions.ReleaseStgMedium(medium)
        Shell._w.Content = Nothing
        Mouse.OverrideCursor = Nothing
    End Sub

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
        If _dragSourceHelper Is Nothing Then
            _dragSourceHelper = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_DragDropHelper))
        End If
        Dim dragSourceHelper2 As IDragSourceHelper2 = _dragSourceHelper
        dragSourceHelper2.SetFlags(1)
        Dim dragImage As SHDRAGIMAGE
        dragImage.sizeDragImage.Width = _bitmap.Width
        dragImage.sizeDragImage.Height = _bitmap.Height
        dragImage.ptOffset.x = ICON_SIZE / 2
        dragImage.ptOffset.y = ICON_SIZE / 2
        dragImage.hbmpDragImage = _bitmap.GetHbitmap()
        dragImage.crColorKey = System.Drawing.Color.Purple.ToArgb()
        Debug.WriteLine("InitializeFromBitmap returned " & _dragSourceHelper.InitializeFromBitmap(dragImage, _dataObject))
        Functions.DeleteObject(dragImage.hbmpDragImage)

        Dim format2 As FORMATETC, medium2 As STGMEDIUM
        format2.cfFormat = Functions.RegisterClipboardFormat("IsShowingLayered")
        If _dataObject.QueryGetData(format2) = 0 Then
            _dataObject.GetData(format2, medium2)
            If Not IntPtr.Zero.Equals(medium2.unionmember) Then
                Dim format As FORMATETC, medium As STGMEDIUM
                format.cfFormat = Functions.RegisterClipboardFormat("DragWindow")
                If _dataObject.QueryGetData(format) = 0 Then
                    addDropDescription(_dataObject, DROPEFFECT.DROPEFFECT_COPY)
                    _dataObject.GetData(format, medium)
                    Dim hwnd As Integer = Marshal.ReadInt32(medium.unionmember)
                    Functions.SendMessage(hwnd, WM.USER + 2, 0, 0)
                End If
            End If
        End If
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
        _addDragImageAction.Invoke()

        If dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_MOVE) Then
            Mouse.OverrideCursor = _moveCursor
        ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_COPY) Then
            Mouse.OverrideCursor = _copyCursor
        ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_LINK) Then
            Mouse.OverrideCursor = _linkCursor
        Else
            Mouse.OverrideCursor = Cursors.No
        End If

        Debug.WriteLine(CType(dwEffect, DROPEFFECT).ToString())

        Return DragDropResult.S_OK
    End Function
End Class
