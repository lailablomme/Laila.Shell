Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Media
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class Previewer
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(Previewer), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(Previewer), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))

        Private Const ERROR_MESSAGE As String = "There was an error displaying the preview ({0})."

        Private _handler As IPreviewHandler
        Private _stream As IStream
        Private _window As Window
        Private _isThumbnail As Boolean
        Private _errorText As String
        Private _timer As Timer
        Private _isMade As Boolean
        Private _previewItem As Item
        Private PART_Message As TextBlock
        Private PART_Thumbnail As Image

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(Previewer), New FrameworkPropertyMetadata(GetType(Previewer)))
        End Sub

        Public Sub New()
            AddHandler Shell.Notification, AddressOf shell_Notification

            AddHandler Me.SizeChanged,
                Sub(s As Object, e As SizeChangedEventArgs)
                    If Not _window Is Nothing Then
                        'updateWindowCoords(Me)
                        UIHelper.OnUIThreadAsync(
                            Sub()
                                updateWindowCoords(Me)
                            End Sub)
                    ElseIf _isThumbnail Then
                        setThumbnail()
                    End If
                End Sub
            AddHandler Me.IsVisibleChanged,
                Sub(s As Object, e As DependencyPropertyChangedEventArgs)
                    If Me.IsVisible Then
                        showPreview(Me)
                    Else
                        hidePreview(Me)
                    End If
                End Sub
            AddHandler Me.Loaded,
                Sub(s As Object, e As EventArgs)
                    setMessage()

                    Dim owner As Window = Window.GetWindow(Me)
                    AddHandler owner.LocationChanged,
                        Sub(s2 As Object, e2 As EventArgs)
                            If Not _window Is Nothing Then
                                updateWindowCoords(Me)
                            ElseIf _isThumbnail Then
                                setThumbnail()
                            End If
                        End Sub
                End Sub
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_Message = Me.Template.FindName("PART_Message", Me)
            PART_Thumbnail = Me.Template.FindName("PART_Thumbnail", Me)
        End Sub

        Private Sub setMessage()
            If Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0 Then
                PART_Message.Text = "Select the file for which you want to display a preview."
                PART_Message.Visibility = Visibility.Visible
            ElseIf Not String.IsNullOrWhiteSpace(_errorText) Then
                PART_Message.Text = _errorText
                PART_Message.Visibility = Visibility.Visible
            ElseIf _handler Is Nothing AndAlso Not _isThumbnail AndAlso _isMade Then
                PART_Message.Text = "Preview is not available."
                PART_Message.Visibility = Visibility.Visible
            Else
                PART_Message.Visibility = Visibility.Collapsed
            End If
        End Sub

        Private Sub setThumbnail()
            If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 Then
                Dim previewItem As Item = Me.SelectedItems(Me.SelectedItems.Count - 1)
                Dim dpi As DpiScale = VisualTreeHelper.GetDpi(Me)
                Dim size As Integer
                If Me.ActualHeight < Me.ActualWidth Then
                    size = Me.ActualWidth / (dpi.PixelsPerInchX / 96.0)
                Else
                    size = Me.ActualHeight / (dpi.PixelsPerInchY / 96.0)
                End If
                PART_Thumbnail.Source = previewItem.Image(size - 10)
                PART_Thumbnail.Visibility = Visibility.Visible
            End If
        End Sub

        Public Overridable Property SelectedItems As IEnumerable(Of Item)
            Get
                Return GetValue(SelectedItemsProperty)
            End Get
            Set(value As IEnumerable(Of Item))
                SetCurrentValue(SelectedItemsProperty, value)
            End Set
        End Property

        Shared Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            hidePreview(d)

            Dim previewer As Previewer = d
            If Not previewer._timer Is Nothing Then
                previewer._timer.Dispose()
                previewer._timer = Nothing
            End If
            previewer._timer = New Timer(New TimerCallback(
                Sub()
                    If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                        UIHelper.OnUIThread(
                            Sub()
                                If Not previewer._timer Is Nothing Then
                                    previewer._timer.Dispose()
                                    previewer._timer = Nothing
                                End If

                                showPreview(d)
                            End Sub)
                    End If
                End Sub), Nothing, 500, Timeout.Infinite)
        End Sub

        Private Shared Sub showPreview(previewer As Previewer)
            previewer._errorText = Nothing

            If Not previewer.SelectedItems Is Nothing AndAlso previewer.SelectedItems.Count > 0 AndAlso Not previewer._isMade Then
                previewer._previewItem = previewer.SelectedItems(previewer.SelectedItems.Count - 1)
                Debug.WriteLine("PreviewItem=" & previewer._previewItem.FullPath)

                previewer._isThumbnail = ImageHelper.IsImage(previewer._previewItem.FullPath)
                If previewer._isThumbnail Then
                    previewer.setThumbnail()
                Else
                    Dim clsid As Guid = getHandlerCLSID(IO.Path.GetExtension(previewer._previewItem.FullPath))
                    If Not Guid.Empty.Equals(clsid) Then
                        Debug.WriteLine("IPreviewHandler=" & clsid.ToString())
                        Dim cc As ClassContext = ClassContext.LocalServer
                        If Debugger.IsAttached Then cc = cc Or ClassContext.InProcServer
                        Dim h As HRESULT = Functions.CoCreateInstance(clsid, IntPtr.Zero, cc, GetType(IPreviewHandler).GUID, previewer._handler)
                        If h = HRESULT.S_OK Then
                            h = HRESULT.S_FALSE
                            If Not previewer._handler Is Nothing Then
                                If h <> HRESULT.S_OK AndAlso TypeOf previewer._handler Is IInitializeWithStream Then
                                    If IO.File.Exists(previewer._previewItem.FullPath) Then
                                        h = Functions.SHCreateStreamOnFileEx(previewer._previewItem.FullPath, STGM.STGM_READ Or STGM.STGM_SHARE_DENY_NONE, 0, 0, IntPtr.Zero, previewer._stream)
                                        Debug.WriteLine("SHCreateStreamOnFileEx=" & h.ToString())
                                    Else
                                        SyncLock previewer._previewItem._shellItemLock
                                            h = previewer._previewItem.ShellItem2.BindToHandler(Nothing, Guids.BHID_Stream, GetType(IStream).GUID, previewer._stream)
                                        End SyncLock
                                        Debug.WriteLine("BHID_Stream=" & h.ToString())
                                    End If
                                    If h = HRESULT.S_OK Then
                                        h = CType(previewer._handler, IInitializeWithStream).Initialize(previewer._stream, STGM.STGM_READ)
                                        Debug.WriteLine("IPreviewHandler.IInitializeWithStream=" & h.ToString())
                                    End If
                                End If
                                If h <> HRESULT.S_OK AndAlso TypeOf previewer._handler Is IInitializeWithFile Then
                                    h = CType(previewer._handler, IInitializeWithFile).Initialize(previewer._previewItem.FullPath, STGM.STGM_READ)
                                    Debug.WriteLine("IPreviewHandler.IInitializeWithFile=" & h.ToString())
                                End If
                                If h <> HRESULT.S_OK AndAlso TypeOf previewer._handler Is IInitializeWithItem Then
                                    h = CType(previewer._handler, IInitializeWithItem).Initialize(previewer._previewItem, STGM.STGM_READ)
                                    Debug.WriteLine("IPreviewHandler.IInitializeWithItem=" & h.ToString())
                                End If
                                If h <> HRESULT.S_OK Then
                                    previewer._errorText = String.Format(Previewer.ERROR_MESSAGE, h)
                                    hidePreview(previewer)
                                End If
                            End If
                        Else
                            previewer._errorText = String.Format(Previewer.ERROR_MESSAGE, h)
                        End If

                        If Not previewer._handler Is Nothing AndAlso previewer.IsVisible Then
                            makeWindow(previewer)
                            Dim ih As New WindowInteropHelper(previewer._window)
                            ih.EnsureHandle()
                            Dim hwnd As IntPtr = ih.Handle
                            'Dim color As Color = Colors.Red
                            'Dim c As UInteger = (CUInt(color.B) << 16) Or (CUInt(color.G) << 8) Or CUInt(color.R)
                            'CType(previewer._handler, IPreviewHandlerVisuals).SetTextColor(c)
                            Dim el As IInputElement = Keyboard.FocusedElement
                            previewer._handler.SetWindow(hwnd, getRect(previewer))
                            h = previewer._handler.DoPreview()
                            Debug.WriteLine("IPreviewHandler.DoPreview=" & h.ToString())
                            If h <> HRESULT.S_OK Then
                                previewer._errorText = String.Format(Previewer.ERROR_MESSAGE, h)
                                hidePreview(previewer)
                            Else
                                previewer._handler.SetRect(getRect(previewer))
                            End If
                            If Not el Is Nothing AndAlso Not previewer._window Is Nothing _
                                AndAlso previewer._window.IsKeyboardFocusWithin Then el.Focus()
                        End If
                    End If
                End If

                previewer._isMade = True
                previewer.setMessage()
            End If
        End Sub

        Private Shared Sub hidePreview(previewer As Previewer)
            If Not previewer._handler Is Nothing Then
                Debug.WriteLine("Unloading IPreviewHandler")
                previewer._handler.Unload()
                Marshal.ReleaseComObject(previewer._handler)
                previewer._handler = Nothing
            End If
            If Not previewer._stream Is Nothing Then
                Marshal.ReleaseComObject(previewer._stream)
                previewer._stream = Nothing
            End If
            If Not previewer._window Is Nothing Then
                RemoveHandler previewer._window.Owner.SizeChanged, AddressOf previewer.owner_SizeChanged
                previewer._window.Close()
                previewer._window = Nothing
            End If

            If previewer._isThumbnail Then
                previewer._isThumbnail = False
                previewer.PART_Thumbnail.Visibility = Visibility.Collapsed
            End If

            previewer._previewItem = Nothing
            previewer._isMade = False

            previewer.setMessage()
        End Sub

        Private Shared Function makeWindow(previewer As Previewer)
            previewer._window = New Window()
            previewer._window.SetValue(System.Windows.Shell.WindowChrome.WindowChromeProperty,
                New System.Windows.Shell.WindowChrome() With {
                    .CaptionHeight = 0,
                    .ResizeBorderThickness = New Thickness(0),
                    .CornerRadius = New CornerRadius(0),
                    .GlassFrameThickness = New Thickness(0)
                })
            previewer._window.WindowStyle = WindowStyle.None
            previewer._window.Background = Brushes.White
            previewer._window.ResizeMode = ResizeMode.NoResize
            previewer._window.ShowInTaskbar = False
            previewer._window.Owner = Window.GetWindow(previewer)
            AddHandler previewer._window.Owner.SizeChanged, AddressOf previewer.owner_SizeChanged
            AddHandler previewer._window.Owner.IsVisibleChanged, AddressOf previewer.owner_VisibilityChanged
            Dim descriptor As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(Window.OpacityProperty, GetType(Window))
            descriptor.AddValueChanged(previewer._window.Owner, AddressOf previewer.owner_OpacityChanged)
            previewer._window.ShowActivated = False
            updateWindowCoords(previewer)
            previewer._window.Show()
        End Function

        Private Sub owner_OpacityChanged(s As Object, e As EventArgs)
            If CType(s, Window).Opacity > 0 Then
                showPreview(Me)
            Else
                hidePreview(Me)
            End If
        End Sub

        Private Sub owner_VisibilityChanged(s As Object, e As DependencyPropertyChangedEventArgs)
            If CType(s, Window).IsVisible Then
                showPreview(Me)
            Else
                hidePreview(Me)
            End If
        End Sub

        Private Sub owner_SizeChanged(s As Object, e As SizeChangedEventArgs)
            If Not _window Is Nothing Then
                updateWindowCoords(Me)
                UIHelper.OnUIThreadAsync(
                    Sub()
                        updateWindowCoords(Me)
                    End Sub)
            ElseIf _isThumbnail Then
                setThumbnail()
            End If
        End Sub

        Private Shared Sub updateWindowCoords(previewer As Previewer)
            Dim leftTop As Point = previewer.PointToScreen(New Point(0, 0))
            Dim dpi As DpiScale = VisualTreeHelper.GetDpi(previewer)
            If Not previewer._window Is Nothing Then
                previewer._window.Left = leftTop.X / (dpi.PixelsPerInchX / 96.0)
                previewer._window.Top = leftTop.Y / (dpi.PixelsPerInchY / 96.0)
                previewer._window.Width = previewer.ActualWidth
                previewer._window.Height = previewer.ActualHeight
            End If
            If Not previewer._handler Is Nothing Then
                previewer._handler.SetRect(getRect(previewer))
            End If
        End Sub

        Private Shared Function getRect(previewer As Previewer) As WIN32RECT
            Dim result As WIN32RECT = New WIN32RECT()
            Dim dpi As DpiScale = VisualTreeHelper.GetDpi(previewer)
            result.Right = previewer.ActualWidth * (dpi.PixelsPerInchX / 96.0)
            result.Bottom = previewer.ActualHeight * (dpi.PixelsPerInchY / 96.0)
            Return result
        End Function

        Private Shared Function getHandlerCLSID(ext As String) As Guid
            Dim output As String = New String(ChrW(0), 260) ' Initial buffer
            Dim bufferSize As UInteger = CUInt(output.Length)

            Functions.AssocQueryStringW(AssocF.None, AssocStr.ShellExtension, ext,
                                        GetType(IPreviewHandler).GUID.ToString("B"), output, bufferSize)
            output = output.Substring(0, CInt(bufferSize - 1))

            If output.Length = 38 Then
                Return New Guid(output)
            Else
                output = New String(ChrW(0), 260)
                Functions.AssocQueryStringW(AssocF.None, AssocStr.ShellExtension, ext & "file",
                                            GetType(IPreviewHandler).GUID.ToString("B"), output, bufferSize)
                output = output.Substring(0, CInt(bufferSize - 1))

                If output.Length = 38 Then
                    Return New Guid(output)
                Else
                    Return Guid.Empty
                End If
            End If
        End Function

        Protected Overridable Sub shell_Notification(sender As Object, e As NotificationEventArgs)
            Dim previewItem As Item = _previewItem

            Select Case e.Event
                Case SHCNE.UPDATEDIR
                    If Not previewItem Is Nothing AndAlso previewItem.Parent.FullPath = e.Item1.FullPath Then
                        UIHelper.OnUIThread(
                            Sub()
                                hidePreview(Me)
                                showPreview(Me)
                            End Sub)
                    End If
                Case SHCNE.UPDATEITEM
                    If Not previewItem Is Nothing _
                        AndAlso (previewItem.Parent.FullPath = e.Item1.FullPath OrElse previewItem.FullPath = e.Item1.FullPath) Then
                        UIHelper.OnUIThread(
                            Sub()
                                hidePreview(Me)
                                showPreview(Me)
                            End Sub)
                    End If
            End Select
        End Sub
    End Class
End Namespace