Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop.Windows
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Windows.Media

Namespace Controls.Parts
    Public Class HorizontalDragViewStrategy
        Implements IDragViewStrategy

        Private PART_DragInsertIndicator As Grid
        Private _parent As BaseFolderView

        Public Sub New(dragInsertIndicator As Grid, parent As BaseFolderView)
            Me.PART_DragInsertIndicator = dragInsertIndicator
            _parent = parent
        End Sub

        Public Sub GetInsertIndex(ptWIN32 As WIN32POINT, overListBoxItem As ListBoxItem, overItem As Item, ByRef dragInsertParent As ISupportDragInsert, ByRef insertIndex As Integer) Implements IDragViewStrategy.GetInsertIndex
            Dim ptItem As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, overListBoxItem)
            If ptItem.X <= 10 Then
                If TypeOf overItem.LogicalParent Is ISupportDragInsert Then
                    dragInsertParent = overItem.LogicalParent
                    Dim view As ListCollectionView = CollectionViewSource.GetDefaultView(overItem.LogicalParent.Items)
                    insertIndex = view.IndexOf(overItem)
                    Debug.WriteLine("* Drag: insert before item")
                End If
            ElseIf ptItem.X >= overListBoxItem.ActualWidth - 10 Then
                If TypeOf overItem.LogicalParent Is ISupportDragInsert Then
                    dragInsertParent = overItem.LogicalParent
                    Debug.WriteLine("* Drag: insert after not expanded item")
                    Dim view As ListCollectionView = CollectionViewSource.GetDefaultView(overItem.LogicalParent.Items)
                    insertIndex = view.IndexOf(overItem) + 1
                End If
            Else
                Debug.WriteLine("* Full over item")
                dragInsertParent = Nothing
            End If
        End Sub

        Public Sub SetDragInsertIndicator(overListBoxItem As ListBoxItem, overItem As Item, visibility As Visibility, beforeIndex As Integer) Implements IDragViewStrategy.SetDragInsertIndicator
            If visibility <> visibility.Collapsed Then
                Dim size As Size = _parent.GetListBoxClientSize()

                Dim groupByPropertyKeyText As String = Nothing
                If Not String.IsNullOrWhiteSpace(If(overItem.LogicalParent, overItem.Parent).ItemsGroupByPropertyName) Then
                    groupByPropertyKeyText = If(overItem.LogicalParent, overItem.Parent).ItemsGroupByPropertyName _
                        .Substring(If(overItem.LogicalParent, overItem.Parent).ItemsGroupByPropertyName.IndexOf("[") + 1)
                    groupByPropertyKeyText = groupByPropertyKeyText.Substring(0, groupByPropertyKeyText.IndexOf("]"))
                End If

                Dim view As ListCollectionView = CollectionViewSource.GetDefaultView(overItem.LogicalParent.Items)
                If view.IndexOf(overItem) = beforeIndex Then
                    Dim pt As Point = _parent.PointFromScreen(overListBoxItem.PointToScreen(New Point(0, 0)))

                    Dim prevItem As ListViewItem = _parent.PART_ListBox.ItemContainerGenerator.ContainerFromIndex(beforeIndex - 1)
                    If Not prevItem Is Nothing AndAlso Not overListBoxItem.Equals(prevItem) _
                        AndAlso Me.GetAreAdjecentItems(overListBoxItem, prevItem) _
                        AndAlso (String.IsNullOrWhiteSpace(groupByPropertyKeyText) _
                                 OrElse overItem.PropertiesByKeyAsText(groupByPropertyKeyText).Value _
                                    .Equals(CType(prevItem.DataContext, Item).PropertiesByKeyAsText(groupByPropertyKeyText).Value)) Then
                        Dim pt2 As Point = _parent.PointFromScreen(prevItem.PointToScreen(New Point(prevItem.ActualWidth, 0)))
                        pt.X = pt.X - (pt.X - pt2.X) / 2
                    End If

                    Me.PART_DragInsertIndicator.Margin = New Thickness(pt.X - 3, pt.Y, 0, 0)
                    Me.PART_DragInsertIndicator.Height = Math.Max(overListBoxItem.ActualHeight, If(prevItem?.ActualHeight, 0))
                Else
                    Dim pt As Point = _parent.PointFromScreen(overListBoxItem.PointToScreen(New Point(overListBoxItem.ActualWidth, 0)))

                    Dim nextItem As ListViewItem = _parent.PART_ListBox.ItemContainerGenerator.ContainerFromIndex(beforeIndex)
                    If Not nextItem Is Nothing AndAlso Not overListBoxItem.Equals(nextItem) _
                        AndAlso Me.GetAreAdjecentItems(overListBoxItem, nextItem) _
                        AndAlso (String.IsNullOrWhiteSpace(groupByPropertyKeyText) _
                                 OrElse overItem.PropertiesByKeyAsText(groupByPropertyKeyText).Value _
                                    .Equals(CType(nextItem.DataContext, Item).PropertiesByKeyAsText(groupByPropertyKeyText).Value)) Then
                        Dim pt2 As Point = _parent.PointFromScreen(nextItem.PointToScreen(New Point(0, 0)))
                        pt.X = pt.X + (pt2.X - pt.X) / 2
                    End If

                    Me.PART_DragInsertIndicator.Margin = New Thickness(pt.X - 3, pt.Y, 0, 0)
                    Me.PART_DragInsertIndicator.Height = Math.Max(overListBoxItem.ActualHeight, If(nextItem?.ActualHeight, 0))
                End If

                Dim rectGeometry As New RectangleGeometry(
                    New Rect(-Me.PART_DragInsertIndicator.Margin.Left, -Me.PART_DragInsertIndicator.Margin.Top,
                             size.Width, size.Height))
                Me.PART_DragInsertIndicator.Clip = rectGeometry
            End If
            Me.PART_DragInsertIndicator.Visibility = visibility
        End Sub

        Public Function GetOverListBoxItem(ptWIN32 As WIN32POINT) As ListBoxItem Implements IDragViewStrategy.GetOverListBoxItem
            ' translate point to listview
            Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _parent.PART_ListBox)

            ' find which item we're over
            Dim overTreeViewItem As ListBoxItem
            Dim overObject As IInputElement = _parent.PART_ListBox.InputHitTest(pt)
            If TypeOf overObject Is ListBoxItem Then
                overTreeViewItem = overObject
            Else
                overTreeViewItem = UIHelper.GetParentOfType(Of ListBoxItem)(overObject)
            End If
            ' apprently there's some space between the items, so try again
            If overTreeViewItem Is Nothing Then
                overObject = _parent.PART_ListBox.InputHitTest(New Point(pt.X - 10, pt.Y))
                If TypeOf overObject Is ListBoxItem Then
                    overTreeViewItem = overObject
                Else
                    overTreeViewItem = UIHelper.GetParentOfType(Of ListBoxItem)(overObject)
                End If
            End If
            ' apprently there's some space between the items, so try again
            If overTreeViewItem Is Nothing Then
                overObject = _parent.PART_ListBox.InputHitTest(New Point(pt.X + 10, pt.Y))
                If TypeOf overObject Is ListBoxItem Then
                    overTreeViewItem = overObject
                Else
                    overTreeViewItem = UIHelper.GetParentOfType(Of ListBoxItem)(overObject)
                End If
            End If

            Return If(Not overTreeViewItem Is Nothing,
                overTreeViewItem, Nothing)
        End Function

        Public Function GetAreAdjecentItems(item1 As ListBoxItem, item2 As ListBoxItem) As Boolean Implements IDragViewStrategy.GetAreAdjecentItems
            Dim pt1 As Point = _parent.PointFromScreen(item1.PointToScreen(New Point(0, 0)))
            Dim pt2 As Point = _parent.PointFromScreen(item2.PointToScreen(New Point(0, 0)))
            Return pt1.Y = pt2.Y
        End Function
    End Class
End Namespace