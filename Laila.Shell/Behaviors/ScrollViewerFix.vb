Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Behaviors
    Public Class ScrollViewerFix
        Public Shared Function GetIsEnabled(obj As DependencyObject) As Boolean
            Return obj.GetValue(IsEnabledProperty)
        End Function

        Public Shared Sub SetIsEnabled(obj As DependencyObject, value As Boolean)
            obj.SetValue(IsEnabledProperty, value)
        End Sub

        Public Shared ReadOnly IsEnabledProperty As DependencyProperty =
          DependencyProperty.RegisterAttached(
              "IsEnabled",
              GetType(Boolean),
              GetType(ScrollViewerFix),
              New UIPropertyMetadata(
                  False,
                  Sub(o As DependencyObject, e As DependencyPropertyChangedEventArgs)
                      Dim scrollViewer As ScrollViewer = o
                      If e.NewValue Then
                          AddHandler scrollViewer.PreviewMouseWheel, AddressOf HandlePreviewMouseWheel
                      Else
                          RemoveHandler scrollViewer.PreviewMouseWheel, AddressOf HandlePreviewMouseWheel
                      End If
                  End Sub))

        Private Shared Sub HandlePreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
            Dim sv As ScrollViewer = sender
            Dim parent As UIElement = UIHelper.GetParent(sv)
            Dim hit As Boolean = hitTopOrBottom(e.Delta, sv)
            If parent Is Nothing OrElse Not hit Then Return

            Dim argsCopy = New MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta) With {
                .RoutedEvent = UIElement.MouseWheelEvent,
                .Source = e.Source
            }
            parent.RaiseEvent(argsCopy)
        End Sub

        Private Shared Function hitTopOrBottom(delta As Double, sv As ScrollViewer) As Boolean
            Dim contentVerticalOffset = sv.ContentVerticalOffset
            Dim atTop As Boolean = contentVerticalOffset = 0
            Dim movedUp As Boolean = delta > 0
            Dim hitTop As Boolean = atTop AndAlso movedUp

            Dim atBottom As Boolean = contentVerticalOffset = sv.ScrollableHeight
            Dim movedDown As Boolean = delta < 0
            Dim hitBottom As Boolean = atBottom AndAlso movedDown

            Return hitTop OrElse hitBottom
        End Function
    End Class
End Namespace