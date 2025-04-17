Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class ItemsContextMenu
        Inherits System.Windows.Controls.ContextMenu

        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColor", GetType(Boolean), GetType(ItemsContextMenu), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColorOverride", GetType(Boolean?), GetType(ItemsContextMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged))

        Public Event ItemClicked As EventHandler

        Private PART_ListBox As ListBox

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ItemsContextMenu), New FrameworkPropertyMetadata(GetType(ItemsContextMenu)))
        End Sub

        Protected Overrides Sub OnOpened(e As RoutedEventArgs)
            MyBase.OnOpened(e)

            UIHelper.OnUIThreadAsync(
                Sub()
                    Me.Focus()
                    Me.PART_ListBox.Focus()
                End Sub, Threading.DispatcherPriority.ContextIdle)
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_ListBox = Me.Template.FindName("PART_ListBox", Me)

            AddHandler Me.PART_ListBox.PreviewMouseUp, AddressOf onListBoxPreviewMouseUp
            AddHandler Me.PART_ListBox.PreviewKeyDown, AddressOf onListBoxPreviewKeyDown

            AddHandler Shell.Settings.PropertyChanged,
                Sub(s As Object, e As PropertyChangedEventArgs)
                    Select Case e.PropertyName
                        Case "DoShowEncryptedOrCompressedFilesInColor"
                            setDoShowEncryptedOrCompressedFilesInColor()
                    End Select
                End Sub
            setDoShowEncryptedOrCompressedFilesInColor()
        End Sub

        Private Sub onListBoxPreviewKeyDown(sender As Object, e As KeyEventArgs)
            Select Case e.Key
                Case Key.Up
                    If Me.PART_ListBox.Equals(Keyboard.FocusedElement) Then
                        Dim listBoxItem As ListBoxItem = Me.PART_ListBox.ItemContainerGenerator.ContainerFromIndex(Me.PART_ListBox.Items.Count - 1)
                        If Not listBoxItem Is Nothing Then
                            listBoxItem.Focus()
                            e.Handled = True
                        End If
                    End If
                Case Key.Down
                    If Me.PART_ListBox.Equals(Keyboard.FocusedElement) Then
                        Dim listBoxItem As ListBoxItem = Me.PART_ListBox.ItemContainerGenerator.ContainerFromIndex(0)
                        If Not listBoxItem Is Nothing Then
                            listBoxItem.Focus()
                            e.Handled = True
                        End If
                    End If
                Case Key.Space, Key.Enter
                    If Not Keyboard.FocusedElement Is Nothing AndAlso TypeOf Keyboard.FocusedElement Is ListBoxItem Then
                        Using Shell.OverrideCursor(Cursors.Wait)
                            Dim listBoxItem As ListBoxItem = Keyboard.FocusedElement
                            Dim clickedItem As Item = TryCast(listBoxItem?.DataContext, Item)
                            If Not clickedItem Is Nothing Then
                                RaiseEvent ItemClicked(clickedItem, New EventArgs())
                                Me.IsOpen = False
                                e.Handled = True
                            End If
                        End Using
                    End If
            End Select
        End Sub

        Private Sub onListBoxPreviewMouseUp(sender As Object, e As MouseButtonEventArgs)
            If Not e.OriginalSource Is Nothing Then
                Using Shell.OverrideCursor(Cursors.Wait)
                    Dim listBoxItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
                    Dim clickedItem As Item = TryCast(listBoxItem?.DataContext, Item)
                    If Not clickedItem Is Nothing Then
                        RaiseEvent ItemClicked(clickedItem, New EventArgs())
                        Me.IsOpen = False
                        e.Handled = True
                    End If
                End Using
            End If
        End Sub

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
            Dim icm As ItemsContextMenu = d
            icm.setDoShowEncryptedOrCompressedFilesInColor()
        End Sub
    End Class
End Namespace