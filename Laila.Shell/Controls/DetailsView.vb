Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Helpers
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Namespace Controls
    Public Class DetailsView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(DetailsView), New FrameworkPropertyMetadata(GetType(DetailsView)))
        End Sub

        Protected Overrides Sub ClearBinding()
            MyBase.ClearBinding()
            If Not Me.PART_ListView Is Nothing Then
                CType(Me.PART_ListView.View, GridView).Columns.Clear()
            End If
        End Sub

        Protected Overrides Sub MakeBinding()
            MyBase.MakeBinding()
            If Not Me.PART_ListView Is Nothing Then
                Me.ColumnsIn = buildColumnsIn()
            End If
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listViewItem As ListViewItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size)
            Dim column As Column = Me.Folder.Columns("System.ItemNameDisplay")
            If Not column Is Nothing Then
                Dim headers As IEnumerable(Of GridViewColumnHeader) =
                                            UIHelper.FindVisualChildren(Of GridViewColumnHeader)(Me.PART_ListView)
                Dim header As GridViewColumnHeader =
                                            headers.FirstOrDefault(Function(h) Not h.Column Is Nothing _
                                                AndAlso h.Column.GetValue(Behaviors.GridViewExtBehavior.PropertyNameProperty) _
                                                    = String.Format("PropertiesByKeyAsText[{0}].Value", column.PROPERTYKEY.ToString()))
                If Not header Is Nothing Then
                    Dim width As Double = header.ActualWidth
                    Dim ptLeft As Point = Me.PointFromScreen(header.PointToScreen(New Point(0, 0)))
                    If header.Column.GetValue(Behaviors.GridViewExtBehavior.ColumnIndexProperty) = 0 Then
                        ptLeft.X += 20
                        width -= 20
                    End If
                    Dim ptTop As Point = Me.PointFromScreen(listViewItem.PointToScreen(New Point(0, 0)))

                    point.X = ptLeft.X + 5
                    point.Y = ptTop.Y
                    size.Width = width - 5
                    size.Height = listViewItem.ActualHeight
                End If
            End If
            textAlignment = TextAlignment.Left
        End Sub
    End Class
End Namespace