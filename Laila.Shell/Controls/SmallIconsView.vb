Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class SmallIconsView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(SmallIconsView), New FrameworkPropertyMetadata(GetType(SmallIconsView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.DragViewStrategy = New HorizontalDragViewStrategy(Me.PART_DragInsertIndicator, Me)
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double,
                                                       ByRef displayNameElemant As FrameworkElement)
            Dim textBlock As TextBlock = UIHelper.FindVisualChildren(Of TextBlock)(listBoxItem) _
                .FirstOrDefault(Function(b) b.Name = "PART_DisplayName")
            point = Me.PointFromScreen(textBlock.PointToScreen(New Point(0, 0)))
            point.X += -2
            point.Y += -1
            size.Width = Double.NaN
            size.Height = textBlock.ActualHeight + 2
            textAlignment = TextAlignment.Left
            fontSize = textBlock.FontSize
            displayNameElemant = textBlock
        End Sub
    End Class
End Namespace