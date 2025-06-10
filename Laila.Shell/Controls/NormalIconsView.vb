Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class NormalIconsView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(NormalIconsView), New FrameworkPropertyMetadata(GetType(NormalIconsView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.DragViewStrategy = New HorizontalDragViewStrategy(Me.PART_DragInsertIndicator, Me)
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double,
                                                       ByRef displayNameElemant As FrameworkElement)
            point = Me.PointFromScreen(listBoxItem.PointToScreen(New Point(0, 0)))
            point.X += -1
            point.Y += 48 + 2
            listBoxItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = listBoxItem.DesiredSize.Width - 9
            size.Height = listBoxItem.DesiredSize.Height - 48 - 4
            textAlignment = TextAlignment.Center
            fontSize = Me.FontSize
            displayNameElemant = UIHelper.FindVisualChildren(Of TextBlock)(listBoxItem) _
                .FirstOrDefault(Function(b) b.Name = "PART_DisplayName")
        End Sub
    End Class
End Namespace