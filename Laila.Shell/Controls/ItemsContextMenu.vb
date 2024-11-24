Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class ItemsContextMenu
        Inherits System.Windows.Controls.ContextMenu

        Public Event ItemClicked As EventHandler

        Private PART_ListBox As ListBox

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ItemsContextMenu), New FrameworkPropertyMetadata(GetType(ItemsContextMenu)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_ListBox = Me.Template.FindName("PART_ListBox", Me)

            AddHandler Me.PART_ListBox.PreviewMouseUp, AddressOf onListBoxPreviewMouseUp
        End Sub

        Private Sub onListBoxPreviewMouseUp(sender As Object, e As MouseButtonEventArgs)
            If Not e.OriginalSource Is Nothing Then
                Dim listBoxItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
                Dim clickedItem As Item = TryCast(listBoxItem?.DataContext, Item)
                If Not clickedItem Is Nothing Then
                    RaiseEvent ItemClicked(clickedItem, New EventArgs())
                    Me.IsOpen = False
                End If
            End If
        End Sub
    End Class
End Namespace