Imports System.Windows
Imports System.Windows.Controls

Namespace Controls
    Public Class ExtraLargeIconsView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ExtraLargeIconsView), New FrameworkPropertyMetadata(GetType(ExtraLargeIconsView)))
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            point = Me.PointFromScreen(listBoxItem.PointToScreen(New Point(0, 0)))
            point.X += 0
            point.Y += 216 + 2
            listBoxItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = listBoxItem.DesiredSize.Width - 10
            size.Height = listBoxItem.DesiredSize.Height - 216 - 4
            textAlignment = TextAlignment.Center
            fontSize = Me.FontSize
        End Sub
    End Class
End Namespace