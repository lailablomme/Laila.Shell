Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Controls
    Public Class ListView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ListView), New FrameworkPropertyMetadata(GetType(ListView)))
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
                                                       ByRef point As Point, ByRef size As Size)
            point = Me.PointFromScreen(listViewItem.PointToScreen(New Point(0, 0)))
            point.X += 16 + 4 + 2
            point.Y += 1
            listViewItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = listViewItem.DesiredSize.Width - 16 - 4 - 2
            size.Height = listViewItem.DesiredSize.Height
            textAlignment = TextAlignment.Left
        End Sub
    End Class
End Namespace