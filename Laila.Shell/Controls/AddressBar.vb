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
        Inherits AutoCompleteTextBox
        Implements IDisposable

        Private Const INVALID_VALUE As String = "5e979b53-746b-4a0c-9f5f-00fdd22c91d8"
        Private Const CLICKABLE_SPACE_WIDTH As Double = 60

        Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(AddressBar), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared Shadows ReadOnly IsTabStopProperty As DependencyProperty = DependencyProperty.Register("IsTabStop", GetType(Boolean), GetType(AddressBar), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

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
                    Me.PART_NavigationButtons.Visibility = Visibility.Hidden
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
                            OrElse (selectedItem.Pidl Is Nothing AndAlso Not Item.ArePathsEqual(selectedItem.FullPath, Me.Folder.FullPath)) Then
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

        Protected Overrides Sub OnFolderChanged(ByVal e As DependencyPropertyChangedEventArgs)
            If Not e.NewValue Is Nothing Then
                Dim f As Folder = e.NewValue
                Me.SelectedItem = f
                f.AddressBarRoot = Nothing
                f.AddressBarDisplayName = Nothing
                Me.IsLoading = False
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