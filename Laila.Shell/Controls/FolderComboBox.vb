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
        Inherits BaseControl

        Public Shared ReadOnly ShowDropDownButtonProperty As DependencyProperty = DependencyProperty.Register("ShowDropDownButton", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly WatermarkProperty As DependencyProperty = DependencyProperty.Register("Watermark", GetType(String), GetType(FolderComboBox), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsDropDownOpenProperty As DependencyProperty = DependencyProperty.Register("IsDropDownOpen", GetType(Boolean), GetType(FolderComboBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

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
    End Class
End Namespace