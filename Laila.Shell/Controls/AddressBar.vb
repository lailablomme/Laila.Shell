Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Media
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Converters
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop.Windows

Namespace Controls
    Public Class AddressBar
        Inherits Laila.AutoCompleteTextBox.AutoCompleteTextBox
        Implements IDisposable

        Private Const INVALID_VALUE As String = "5e979b53-746b-4a0c-9f5f-00fdd22c91d8"
        Private Const CLICKABLE_SPACE_WIDTH As Double = 60

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(AddressBar), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColor", GetType(Boolean), GetType(AddressBar), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColorOverride", GetType(Boolean?), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged))
        Public Shared Shadows ReadOnly IsTabStopProperty As DependencyProperty = DependencyProperty.Register("IsTabStop", GetType(Boolean), GetType(AddressBar), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoUseLightThemeProperty As DependencyProperty = DependencyProperty.Register("DoUseLightTheme", GetType(Boolean), GetType(AddressBar), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoUseLightThemeOverrideProperty As DependencyProperty = DependencyProperty.Register("DoUseLightThemeOverride", GetType(Boolean?), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoUseLightThemeOverrideChanged))
        Public Shared ReadOnly AccentProperty As DependencyProperty = DependencyProperty.Register("Accent", GetType(Brush), GetType(AddressBar), New FrameworkPropertyMetadata(Brushes.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentColorProperty As DependencyProperty = DependencyProperty.Register("AccentColor", GetType(Color), GetType(AddressBar), New FrameworkPropertyMetadata(Colors.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentColorOverrideProperty As DependencyProperty = DependencyProperty.Register("AccentColorOverride", GetType(Color?), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnAccentColorOverrideChanged))
        Public Shared ReadOnly AccentDarkerProperty As DependencyProperty = DependencyProperty.Register("AccentDarker", GetType(Brush), GetType(AddressBar), New FrameworkPropertyMetadata(Brushes.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentDarkerColorProperty As DependencyProperty = DependencyProperty.Register("AccentDarkerColor", GetType(Color), GetType(AddressBar), New FrameworkPropertyMetadata(Colors.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentLighterProperty As DependencyProperty = DependencyProperty.Register("AccentLighter", GetType(Brush), GetType(AddressBar), New FrameworkPropertyMetadata(Brushes.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly AccentLighterColorProperty As DependencyProperty = DependencyProperty.Register("AccentLighterColor", GetType(Color), GetType(AddressBar), New FrameworkPropertyMetadata(Colors.Blue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ForegroundColorProperty As DependencyProperty = DependencyProperty.Register("ForegroundColor", GetType(Color), GetType(AddressBar), New FrameworkPropertyMetadata(Colors.Black, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ForegroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("ForegroundColorOverride", GetType(Color?), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnForegroundColorOverrideChanged))
        Public Shared ReadOnly BackgroundColorProperty As DependencyProperty = DependencyProperty.Register("BackgroundColor", GetType(Color), GetType(AddressBar), New FrameworkPropertyMetadata(Colors.White, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly BackgroundColorOverrideProperty As DependencyProperty = DependencyProperty.Register("BackgroundColorOverride", GetType(Color?), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnBackgroundColorOverrideChanged))

        Private PART_NavigationButtonsPanel As SelectedFolderControl
        Private PART_NavigationButtons As Border
        Private PART_ClickToEdit As Border
        Private disposedValue As Boolean
        Private _source As HwndSource

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(AddressBar), New FrameworkPropertyMetadata(GetType(AddressBar)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            MyBase.IsTabStop = False

            Shell.AddToControlCache(Me)

            Me.PART_NavigationButtonsPanel = Template.FindName("PART_NavigationButtonsPanel", Me)
            Me.PART_NavigationButtons = Template.FindName("PART_NavigationButtons", Me)
            Me.PART_ClickToEdit = Template.FindName("PART_ClickToEdit", Me)

            AddHandler Me.Loaded,
                Sub(s As Object, e As EventArgs)
                    Me.PART_NavigationButtonsPanel.MaxWidth = If(Me.ActualWidth - CLICKABLE_SPACE_WIDTH <= 0, Me.ActualWidth, Me.ActualWidth - CLICKABLE_SPACE_WIDTH)
                End Sub

            AddHandler Me.SizeChanged,
                Sub(s As Object, e As SizeChangedEventArgs)
                    Me.PART_NavigationButtonsPanel.MaxWidth = If(Me.ActualWidth - CLICKABLE_SPACE_WIDTH <= 0, Me.ActualWidth, Me.ActualWidth - CLICKABLE_SPACE_WIDTH)
                End Sub

            AddHandler Window.GetWindow(Me).PreviewMouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    Dim parent As AddressBar = UIHelper.GetParentOfType(Of AddressBar)(e.OriginalSource)
                    If (parent Is Nothing OrElse Not parent.Equals(Me)) _
                        AndAlso Me.PART_TextBox.IsKeyboardFocusWithin Then
                        Me.Cancel()
                    End If
                End Sub

            AddHandler Window.GetWindow(Me).Deactivated,
                Sub(s As Object, e As EventArgs)
                    If Me.PART_TextBox.IsKeyboardFocusWithin Then
                        Me.Cancel()
                    End If
                End Sub

            Dim hWnd As IntPtr = New WindowInteropHelper(Window.GetWindow(Me)).Handle
            _source = HwndSource.FromHwnd(hWnd)
            _source.AddHook(AddressOf HwndHook)

            AddHandler Shell.Settings.PropertyChanged,
                Sub(s As Object, e As PropertyChangedEventArgs)
                    Select Case e.PropertyName
                        Case "DoShowEncryptedOrCompressedFilesInColor"
                            setDoShowEncryptedOrCompressedFilesInColor()
                    End Select
                End Sub
            setDoShowEncryptedOrCompressedFilesInColor()

            AddHandler Me.PART_TextBox.PreviewGotKeyboardFocus,
                Sub(s As Object, e As KeyboardFocusChangedEventArgs)
                    _isSettingTextInternally = True
                    Me.SelectedItem = Me.Folder
                    Me.Text = Me.Folder.AddressBarDisplayPath
                    _isSettingTextInternally = False
                    PART_NavigationButtons.Visibility = Visibility.Hidden
                    Me.PART_TextBox.SelectionStart = 0
                    Me.PART_TextBox.SelectionLength = Me.Text.Length
                    Me.PART_DropDownButton.IsChecked = True
                End Sub
            AddHandler Me.PART_ClickToEdit.MouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    Me.PART_TextBox.Focus()
                End Sub
            AddHandler Me.PART_TextBox.PreviewKeyDown,
                Sub(s As Object, e As KeyEventArgs)
                    If Not Me.IsDropDownOpen OrElse Me.IsLoadingSuggestions Then
                        Select Case e.Key
                            Case Key.Enter
                                If Not Me.IsDirty Then
                                    Me.Cancel()
                                End If
                            Case Key.Escape
                                If Not Me.IsDirty Then
                                    Me.Cancel()
                                End If
                        End Select
                    End If
                End Sub
        End Sub

        Public Overrides Sub FocusChild()
        End Sub

        Private Function HwndHook(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
            Select Case CType(msg, WM)
                Case WM.NCLBUTTONDOWN, WM.NCMBUTTONDOWN, WM.NCRBUTTONDOWN
                    If Me.PART_TextBox.IsKeyboardFocusWithin Then
                        Me.Cancel()
                    End If
            End Select

            Return IntPtr.Zero
        End Function

        Protected Overrides Sub OnItemSelected()
            Using Shell.OverrideCursor(Cursors.Wait)
                If INVALID_VALUE.Equals(Me.SelectedValue) Then
                    Dim text As String = Me.PART_TextBox.Text
                    Dim item As Item = Item.FromParsingName(text, Nothing, False)
                    If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                        CType(item, Folder).LastScrollOffset = New Point()
                        Me.Folder = item
                        Me.SelectedItem = item
                        Me.Folder.IsVisibleInAddressBar = True
                        AddressBarHistory.Track(Me.Folder)
                    ElseIf Not item Is Nothing AndAlso TypeOf item Is Item Then
                        AddressBarHistory.Track(item)
                        Dim __ = Menus.InvokeDefaultCommand(item)
                        Me.Cancel()
                    Else
                        System.Media.SystemSounds.Asterisk.Play()
                        Me.Cancel()
                    End If
                ElseIf Not Me.SelectedItem Is Nothing Then
                    If TypeOf Me.SelectedItem Is Folder Then
                        Dim selectedItem As Folder = Me.SelectedItem
                        If (Not selectedItem.Pidl Is Nothing AndAlso Not selectedItem.Pidl.Equals(Me.Folder?.Pidl)) _
                            OrElse (selectedItem.Pidl Is Nothing AndAlso Not selectedItem.FullPath?.Equals(Me.Folder.FullPath)) Then
                            Dim folder As Folder = selectedItem.Clone()
                            Me.Folder = folder
                            AddressBarHistory.Track(folder)
                        Else
                            Me.Cancel()
                        End If
                    ElseIf TypeOf Me.SelectedItem Is Item Then
                        AddressBarHistory.Track(Me.SelectedItem)
                        Dim __ = Menus.InvokeDefaultCommand(Me.SelectedItem)
                        Me.Cancel()
                    End If
                Else
                    Me.Cancel()
                End If

                releaseItems()
            End Using
        End Sub

        Protected Overrides Sub Cancel()
            MyBase.Cancel()
            Me.PART_NavigationButtons.Visibility = Visibility.Visible
            releaseItems()
        End Sub

        Private Sub releaseItems()
            Dim provider As AddressBarSuggestionProvider = Me.Provider
            SyncLock provider._lock
                ' release folder
                If Not provider._folder Is Nothing AndAlso Not Me.PART_NavigationButtonsPanel.VisibleFolders.Contains(provider._folder) Then
                    provider._folder.IsVisibleInAddressBar = False
                    provider._folder = Nothing
                End If

                ' release items
                If Not provider._items Is Nothing Then
                    For Each item In provider._items.Where(Function(i) Not TypeOf i Is Folder OrElse Not Me.PART_NavigationButtonsPanel.VisibleFolders.Contains(i))
                        item.IsVisibleInAddressBar = False
                    Next
                    provider._items = Nothing
                End If
            End SyncLock
        End Sub

        Protected Overrides Sub TextBox_LostFocus(s As Object, e As RoutedEventArgs)
            Me.Cancel()
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                Me.IsLoading = True
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim ab As AddressBar = TryCast(d, AddressBar)
            If Not e.NewValue Is Nothing Then
                Dim f As Folder = e.NewValue
                ab.SelectedItem = f
                f.AddressBarRoot = Nothing
                f.AddressBarDisplayName = Nothing
                ab.IsLoading = False
            End If
        End Sub

        Public Property IsLoading As Boolean
            Get
                Return GetValue(IsLoadingProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsLoadingProperty, value)
            End Set
        End Property

        Public Property DoShowEncryptedOrCompressedFilesInColor As Boolean
            Get
                Return GetValue(DoShowEncryptedOrCompressedFilesInColorProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowEncryptedOrCompressedFilesInColorProperty, value)
            End Set
        End Property

        Private Sub setDoShowEncryptedOrCompressedFilesInColor()
            If Me.DoShowEncryptedOrCompressedFilesInColorOverride.HasValue Then
                Me.DoShowEncryptedOrCompressedFilesInColor = Me.DoShowEncryptedOrCompressedFilesInColorOverride.Value
            Else
                Me.DoShowEncryptedOrCompressedFilesInColor = Shell.Settings.DoShowEncryptedOrCompressedFilesInColor
            End If
            Me.PART_NavigationButtonsPanel.DoShowEncryptedOrCompressedFilesInColorOverride = Me.DoShowEncryptedOrCompressedFilesInColor
        End Sub

        Public Property DoShowEncryptedOrCompressedFilesInColorOverride As Boolean?
            Get
                Return GetValue(DoShowEncryptedOrCompressedFilesInColorOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowEncryptedOrCompressedFilesInColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim ab As AddressBar = d
            ab.setDoShowEncryptedOrCompressedFilesInColor()
        End Sub

        Public Property DoUseLightTheme As Boolean
            Get
                Return GetValue(DoUseLightThemeProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoUseLightThemeProperty, value)
            End Set
        End Property

        Private Sub setDoUseLightTheme()
            If Me.DoUseLightThemeOverride.HasValue Then
                Me.DoUseLightTheme = Me.DoUseLightThemeOverride.Value
            Else
                Me.DoUseLightTheme = Shell.Settings.DoUseLightTheme
            End If
            setForegroundColor()
            setBackgroundColor()
        End Sub

        Public Property DoUseLightThemeOverride As Boolean?
            Get
                Return GetValue(DoUseLightThemeOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoUseLightThemeOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoUseLightThemeOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As AddressBar = d
            bfv.setDoUseLightTheme()
        End Sub

        Public Property Accent As Brush
            Get
                Return GetValue(AccentProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(AccentProperty, value)
            End Set
        End Property

        Public Property AccentColor As Color
            Get
                Return GetValue(AccentColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(AccentColorProperty, value)
            End Set
        End Property

        Public Property AccentDarker As Brush
            Get
                Return GetValue(AccentDarkerProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(AccentDarkerProperty, value)
            End Set
        End Property

        Public Property AccentDarkerColor As Color
            Get
                Return GetValue(AccentDarkerColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(AccentDarkerColorProperty, value)
            End Set
        End Property

        Public Property AccentLighter As Brush
            Get
                Return GetValue(AccentLighterProperty)
            End Get
            Protected Set(ByVal value As Brush)
                SetCurrentValue(AccentLighterProperty, value)
            End Set
        End Property

        Public Property AccentLighterColor As Color
            Get
                Return GetValue(AccentLighterColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(AccentLighterColorProperty, value)
            End Set
        End Property

        Private Sub setAccentColor()
            If Me.AccentColorOverride.HasValue Then
                Me.AccentColor = Me.AccentColorOverride.Value
            Else
                Me.AccentColor = Shell.Settings.WindowsAccentColor
            End If
            Me.Accent = New SolidColorBrush(Me.AccentColor)
            Me.AccentDarkerColor = New LightnessColorConverter().Convert(Me.AccentColor, Nothing, 0.8, Nothing)
            Me.AccentDarker = New SolidColorBrush(Me.AccentDarkerColor)
            Me.AccentLighterColor = New LightnessColorConverter().Convert(Me.AccentColor, Nothing, 1.2, Nothing)
            Me.AccentLighter = New SolidColorBrush(Me.AccentLighterColor)
        End Sub

        Public Property AccentColorOverride As Color?
            Get
                Return GetValue(AccentColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(AccentColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnAccentColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As AddressBar = d
            bfv.setAccentColor()
        End Sub

        Public Property ForegroundColor As Color
            Get
                Return GetValue(ForegroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(ForegroundColorProperty, value)
            End Set
        End Property

        Private Sub setForegroundColor()
            If Me.ForegroundColorOverride.HasValue Then
                Me.ForegroundColor = Me.ForegroundColorOverride.Value
            Else
                Me.ForegroundColor = If(Me.DoUseLightTheme, Colors.Black, Colors.White)
            End If
            Me.Foreground = New SolidColorBrush(Me.ForegroundColor)
        End Sub

        Public Property ForegroundColorOverride As Color?
            Get
                Return GetValue(ForegroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(ForegroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnForegroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As AddressBar = d
            bfv.setForegroundColor()
        End Sub

        Public Property BackgroundColor As Color
            Get
                Return GetValue(BackgroundColorProperty)
            End Get
            Protected Set(ByVal value As Color)
                SetCurrentValue(BackgroundColorProperty, value)
            End Set
        End Property

        Private Sub setBackgroundColor()
            If Me.BackgroundColorOverride.HasValue Then
                Me.BackgroundColor = Me.BackgroundColorOverride.Value
            Else
                Me.BackgroundColor = If(Me.DoUseLightTheme, Colors.White, Colors.Black)
            End If
            Me.Background = New SolidColorBrush(Me.BackgroundColor)
        End Sub

        Public Property BackgroundColorOverride As Color?
            Get
                Return GetValue(BackgroundColorOverrideProperty)
            End Get
            Set(ByVal value As Color?)
                SetCurrentValue(BackgroundColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnBackgroundColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As AddressBar = d
            bfv.setBackgroundColor()
        End Sub

        Public Overloads Property IsTabStop As Boolean
            Get
                Return GetValue(IsTabStopProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsTabStopProperty, value)
            End Set
        End Property

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    Me.PART_NavigationButtonsPanel.Dispose()

                    _source.RemoveHook(AddressOf HwndHook)

                    Shell.RemoveFromControlCache(Me)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace