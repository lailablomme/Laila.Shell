﻿Imports System.Drawing
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

Public Class Drag
    Implements IDropSource

    Private Const MAX_COL As Integer = 5
    Private Const MAX_ROW As Integer = 5

    Public Shared ICON_SIZE As Integer = 128

    Private Shared _grid As Grid
    Private Shared _dragSourceHelper As IDragSourceHelper
    Public Shared _bitmap As Bitmap
    Private Shared _dataObject As IDataObject
    Private Shared _isDragging As Boolean
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
        _lastEffect = -1
    End Sub

    Public Shared Sub Start(items As IEnumerable(Of Item), button As MK)
        If Not _isDragging Then
            _isDragging = True

            Try
                Debug.WriteLine("Drag.Start")

                Dim shellItemPtr As IntPtr, folderpidl As IntPtr
                Dim pidls(items.Count - 1) As IntPtr, lastpidl As IntPtr, pidlsPtr As IntPtr
                Try
                    shellItemPtr = Marshal.GetIUnknownForObject(items(0).Parent._shellItem2)
                    Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)
                Finally
                    If Not IntPtr.Zero.Equals(shellItemPtr) Then
                        Marshal.Release(shellItemPtr)
                    End If
                End Try
                pidlsPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of IntPtr) * items.Count)
                For i = 0 To items.Count - 1
                    Try
                        shellItemPtr = Marshal.GetIUnknownForObject(items(i)._shellItem2)
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

    Friend Shared Sub InitializeDragImage()
        If _isDragging Then
            Debug.WriteLine("InitializeDragImage")
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

        Return DragDropResult.S_OK
    End Function
End Class
