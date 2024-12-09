Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class SmallIconsView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(SmallIconsView), New FrameworkPropertyMetadata(GetType(SmallIconsView)))
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            Dim textBlock As TextBlock = UIHelper.FindVisualChildren(Of TextBlock)(listBoxItem) _
                .FirstOrDefault(Function(b) b.Name = "PART_DisplayName")
            point = Me.PointFromScreen(textBlock.PointToScreen(New Point(0, 0)))
            point.X += -2
            point.Y += -1
            size.Width = textBlock.ActualWidth + 4
            size.Height = textBlock.ActualHeight + 2
            textAlignment = TextAlignment.Left
            fontSize = textBlock.FontSize
        End Sub
    End Class
End Namespace