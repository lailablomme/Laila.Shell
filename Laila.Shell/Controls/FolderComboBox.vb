Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports Laila.AutoCompleteTextBox
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class FolderComboBox
        Inherits Control

        Public Shared ReadOnly ShowDropDownButtonProperty As DependencyProperty = DependencyProperty.Register("ShowDropDownButton", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly WatermarkProperty As DependencyProperty = DependencyProperty.Register("Watermark", GetType(String), GetType(FolderComboBox), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsDropDownOpenProperty As DependencyProperty = DependencyProperty.Register("IsDropDownOpen", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(FolderComboBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColor", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColorOverride", GetType(Boolean?), GetType(FolderComboBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged))
        Public Shared ReadOnly IsCompactModeProperty As DependencyProperty = DependencyProperty.Register("IsCompactMode", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsCompactModeOverrideProperty As DependencyProperty = DependencyProperty.Register("IsCompactModeOverride", GetType(Boolean?), GetType(FolderComboBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsCompactModeOverrideChanged))
        Public Shared ReadOnly DoShowAllFoldersInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowAllFoldersInTreeView", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowAllFoldersInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowAllFoldersInTreeViewOverride", GetType(Boolean?), GetType(FolderComboBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowAllFoldersInTreeViewOverrideChanged))
        Public Shared ReadOnly DoShowAvailabilityStatusInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowAvailabilityStatusInTreeView", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowAvailabilityStatusInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowAvailabilityStatusInTreeViewOverride", GetType(Boolean?), GetType(FolderComboBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowAvailabilityStatusInTreeViewOverrideChanged))
        Public Shared ReadOnly DoExpandTreeViewToCurrentFolderProperty As DependencyProperty = DependencyProperty.Register("DoExpandTreeViewToCurrentFolder", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoExpandTreeViewToCurrentFolderOverrideProperty As DependencyProperty = DependencyProperty.Register("DoExpandTreeViewToCurrentFolderOverride", GetType(Boolean?), GetType(FolderComboBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoExpandTreeViewToCurrentFolderOverrideChanged))
        Public Shared ReadOnly DoShowLibrariesInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowLibrariesInTreeView", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowLibrariesInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowLibrariesInTreeViewOverride", GetType(Boolean?), GetType(FolderComboBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowLibrariesInTreeViewOverrideChanged))

        Private PART_Popup As Popup
        Private PART_TreeView As TreeView
        Private PART_DropDownButton As ToggleButton
        Private PART_ClickToEdit As Border
        Private PART_SelectedFolderControl As SelectedFolderControl
        Private _overrideCursor As IDisposable

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(FolderComboBox), New FrameworkPropertyMetadata(GetType(FolderComboBox)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_Popup = Me.Template.FindName("PART_Popup", Me)
            Me.PART_TreeView = Me.Template.FindName("PART_TreeView", Me)
            Me.PART_DropDownButton = Me.Template.FindName("PART_DropDownButton", Me)
            Me.PART_ClickToEdit = Template.FindName("PART_ClickToEdit", Me)
            Me.PART_SelectedFolderControl = Template.FindName("PART_SelectedFolderControl", Me)

            AddHandler Me.SizeChanged,
                Sub(s As Object, e As SizeChangedEventArgs)
                    Me.PART_SelectedFolderControl.MaxWidth = If(Me.ActualWidth - Me.PART_DropDownButton.ActualWidth <= 0, Me.ActualWidth, Me.ActualWidth - Me.PART_DropDownButton.ActualWidth)
                End Sub
            AddHandler Me.KeyDown,
                Sub(sender As Object, e As KeyEventArgs)
                    If e.Key = Key.Enter OrElse e.Key = Key.Space OrElse (e.Key = Key.Down AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) Then
                        Me.PART_DropDownButton.IsChecked = True
                        e.Handled = True
                    ElseIf e.Key = Key.Delete Or e.Key = Key.Back Then
                        Me.Folder = Nothing
                        e.Handled = True
                    End If
                End Sub
            AddHandler Me.PreviewMouseDown,
                 Sub(sender As Object, e As MouseButtonEventArgs)
                     Me.Focus()
                 End Sub
            AddHandler Me.PART_ClickToEdit.MouseUp,
                Sub(s As Object, e As MouseButtonEventArgs)
                    Me.PART_DropDownButton.IsChecked = Not Me.PART_DropDownButton.IsChecked
                End Sub
            AddHandler Me.PART_Popup.Opened,
                Sub(sender As Object, e As EventArgs)
                    If Not _overrideCursor Is Nothing Then
                        _overrideCursor.Dispose()
                    End If
                    Me.PART_TreeView.PART_ListBox.Focus()
                    Dim __ = Me.PART_TreeView.SetSelectedFolder(Me.Folder)
                End Sub
            AddHandler Me.PART_Popup.Closed,
                Sub(sender As Object, e As EventArgs)
                    Me.PART_DropDownButton.IsChecked = False
                    Me.IsDropDownOpen = False
                    Me.Focus()
                End Sub
            AddHandler Me.PART_TreeView.AfterFolderOpened,
                Sub(sender As Object, e As EventArgs)
                    Me.PART_DropDownButton.IsChecked = False
                    Me.IsDropDownOpen = False
                End Sub
            AddHandler Me.PART_DropDownButton.Checked,
                Sub(sender As Object, e As EventArgs)
                    _overrideCursor = Shell.OverrideCursor(Cursors.Wait)
                    Me.IsDropDownOpen = True
                End Sub
            AddHandler Me.PART_DropDownButton.Unchecked,
                Sub(sender As Object, e As EventArgs)
                    Me.IsDropDownOpen = False
                End Sub
        End Sub

        Public Property ShowDropDownButton As Boolean
            Get
                Return GetValue(ShowDropDownButtonProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(ShowDropDownButtonProperty, value)
            End Set
        End Property

        Public Property Watermark As String
            Get
                Return GetValue(WatermarkProperty)
            End Get
            Set(value As String)
                SetCurrentValue(WatermarkProperty, value)
            End Set
        End Property

        Public Property IsDropDownOpen As Boolean
            Get
                Return GetValue(IsDropDownOpenProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsDropDownOpenProperty, value)
            End Set
        End Property

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
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
            Dim tv As FolderComboBox = d
            tv.setDoShowEncryptedOrCompressedFilesInColor()
        End Sub

        Public Property IsCompactMode As Boolean
            Get
                Return GetValue(IsCompactModeProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(IsCompactModeProperty, value)
            End Set
        End Property

        Private Sub setIsCompactMode()
            If Me.IsCompactModeOverride.HasValue Then
                Me.IsCompactMode = Me.IsCompactModeOverride.Value
            Else
                Me.IsCompactMode = Shell.Settings.IsCompactMode
            End If
        End Sub

        Public Property IsCompactModeOverride As Boolean?
            Get
                Return GetValue(IsCompactModeOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(IsCompactModeOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnIsCompactModeOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As FolderComboBox = d
            bfv.setIsCompactMode()
        End Sub

        Public Property DoShowAllFoldersInTreeView As Boolean
            Get
                Return GetValue(DoShowAllFoldersInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowAllFoldersInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowAllFoldersInTreeView()
            If Me.DoShowAllFoldersInTreeViewOverride.HasValue Then
                Me.DoShowAllFoldersInTreeView = Me.DoShowAllFoldersInTreeViewOverride.Value
            Else
                Me.DoShowAllFoldersInTreeView = Shell.Settings.DoShowAllFoldersInTreeView
            End If
        End Sub

        Public Property DoShowAllFoldersInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowAllFoldersInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowAllFoldersInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowAllFoldersInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As FolderComboBox = d
            bfv.setDoShowAllFoldersInTreeView()
        End Sub

        Public Property DoShowAvailabilityStatusInTreeView As Boolean
            Get
                Return GetValue(DoShowAvailabilityStatusInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowAvailabilityStatusInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowAvailabilityStatusInTreeView()
            If Me.DoShowAvailabilityStatusInTreeViewOverride.HasValue Then
                Me.DoShowAvailabilityStatusInTreeView = Me.DoShowAvailabilityStatusInTreeViewOverride.Value
            Else
                Me.DoShowAvailabilityStatusInTreeView = Shell.Settings.DoShowAvailabilityStatusInTreeView
            End If
        End Sub

        Public Property DoShowAvailabilityStatusInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowAvailabilityStatusInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowAvailabilityStatusInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowAvailabilityStatusInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As FolderComboBox = d
            bfv.setDoShowAvailabilityStatusInTreeView()
        End Sub

        Public Property DoExpandTreeViewToCurrentFolder As Boolean
            Get
                Return GetValue(DoExpandTreeViewToCurrentFolderProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoExpandTreeViewToCurrentFolderProperty, value)
            End Set
        End Property

        Private Sub setDoExpandTreeViewToCurrentFolder()
            If Me.DoExpandTreeViewToCurrentFolderOverride.HasValue Then
                Me.DoExpandTreeViewToCurrentFolder = Me.DoExpandTreeViewToCurrentFolderOverride.Value
            Else
                Me.DoExpandTreeViewToCurrentFolder = Shell.Settings.DoExpandTreeViewToCurrentFolder
            End If
        End Sub

        Public Property DoExpandTreeViewToCurrentFolderOverride As Boolean?
            Get
                Return GetValue(DoExpandTreeViewToCurrentFolderOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoExpandTreeViewToCurrentFolderOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoExpandTreeViewToCurrentFolderOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As FolderComboBox = d
            bfv.setDoExpandTreeViewToCurrentFolder()
        End Sub

        Public Property DoShowLibrariesInTreeView As Boolean
            Get
                Return GetValue(DoShowLibrariesInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowLibrariesInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowLibrariesInTreeView()
            If Me.DoShowLibrariesInTreeViewOverride.HasValue Then
                Me.DoShowLibrariesInTreeView = Me.DoShowLibrariesInTreeViewOverride.Value
            Else
                Me.DoShowLibrariesInTreeView = Shell.Settings.DoShowLibrariesInTreeView
            End If
        End Sub

        Public Property DoShowLibrariesInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowLibrariesInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowLibrariesInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowLibrariesInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As FolderComboBox = d
            bfv.setDoShowLibrariesInTreeView()
        End Sub
    End Class
End Namespace