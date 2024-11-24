Imports System.ComponentModel
Imports System.Data
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports Laila.Shell.Helpers
Imports WpfToolkit.Controls

Namespace Controls
    Public Class ListView
        Inherits BaseFolderView

        Private _isLoaded As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ListView), New FrameworkPropertyMetadata(GetType(ListView)))
        End Sub

        Protected Overrides Sub OnBeforeRestoreScrollOffset()
            'Dim view As CollectionView = CollectionViewSource.GetDefaultView(ItemsControl.GetItemsOwner(Me).ItemsSource)
            'If Not (view Is Nothing OrElse Not view.GroupDescriptions.Count > 0) Then
            '    Dim totalWidth As Double = 0
            '    For Each g As CollectionViewGroup In view.Groups
            '        totalWidth += g.ItemCount / Math.Floor((Me.ActualHeight - 25) / 20) * 240
            '    Next
            '    _scrollViewer..ScrollableWidth = totalWidth
            'End If
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            point = Me.PointFromScreen(listBoxItem.PointToScreen(New Point(0, 0)))
            point.X += 16 + 4 + 2 + 16 + 4
            point.Y += 1
            listBoxItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = listBoxItem.DesiredSize.Width - 16 - 4 - 2
            size.Height = listBoxItem.DesiredSize.Height
            textAlignment = TextAlignment.Left
            fontSize = Me.FontSize
        End Sub
    End Class
End Namespace