Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Controls
    Public Class ContentView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ContentView), New FrameworkPropertyMetadata(GetType(ContentView)))
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listViewItem As ListViewItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            point = Me.PointFromScreen(listViewItem.PointToScreen(New Point(0, 0)))
            point.X += 16 + 4 + 48 + 2
            point.Y += 1
            listViewItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = 300 + 4
            size.Height = 20
            textAlignment = TextAlignment.Left
            fontSize = 14
        End Sub
    End Class
End Namespace