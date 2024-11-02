Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Controls
    Public Class ExtraLargeIconsView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ExtraLargeIconsView), New FrameworkPropertyMetadata(GetType(ExtraLargeIconsView)))
        End Sub

        Protected Overrides Sub MakeBinding()
            MyBase.MakeBinding()

            Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Folder.Items)
            view.SortDescriptions.Add(New SortDescription() With {
                .PropertyName = "ItemNameDisplaySortValue",
                .Direction = ListSortDirection.Ascending
            })
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listViewItem As ListViewItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            point = Me.PointFromScreen(listViewItem.PointToScreen(New Point(0, 0)))
            point.X += 2
            point.Y += 216 + 2
            listViewItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = listViewItem.DesiredSize.Width - 4
            size.Height = listViewItem.DesiredSize.Height - 216 - 4
            textAlignment = TextAlignment.Center
            fontSize = Me.FontSize
        End Sub
    End Class
End Namespace