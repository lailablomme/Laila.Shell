Imports System.Windows
Imports System.Windows.Controls

Namespace Controls
    Public Class NormalIconsView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(NormalIconsView), New FrameworkPropertyMetadata(GetType(NormalIconsView)))
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            point = Me.PointFromScreen(listBoxItem.PointToScreen(New Point(0, 0)))
            point.X += -1
            point.Y += 48 + 2
            listBoxItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = listBoxItem.DesiredSize.Width - 9
            size.Height = listBoxItem.DesiredSize.Height - 48 - 4
            textAlignment = TextAlignment.Center
            fontSize = Me.FontSize
        End Sub
    End Class
End Namespace