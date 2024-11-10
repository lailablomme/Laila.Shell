Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Controls
    Public Class NormalIconsView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(NormalIconsView), New FrameworkPropertyMetadata(GetType(NormalIconsView)))
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listViewItem As ListViewItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            point = Me.PointFromScreen(listViewItem.PointToScreen(New Point(0, 0)))
            point.X += 2
            point.Y += 48 + 2
            listViewItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = listViewItem.DesiredSize.Width - 4
            size.Height = listViewItem.DesiredSize.Height - 48 - 4
            textAlignment = TextAlignment.Center
            fontSize = Me.FontSize
        End Sub
    End Class
End Namespace