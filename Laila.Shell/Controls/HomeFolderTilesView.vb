Imports Laila.Shell.Behaviors
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Helpers
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media

Namespace Controls
    Public Class HomeFolderTilesView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(HomeFolderTilesView), New FrameworkPropertyMetadata(GetType(HomeFolderTilesView)))
        End Sub

        Private _oldFocusedElement As FrameworkElement = Nothing
        Private _isWorkingFocus As Boolean

        Protected Overrides Sub PART_ListBox_Loaded()
            MyBase.PART_ListBox_Loaded()

            _selectionHelper.Unhook()
            _selectionHelper = Nothing

            AddHandler Me.PART_ListBox.SelectionChanged,
                Sub(s As Object, e As SelectionChangedEventArgs)
                    If e.OriginalSource.Equals(Me.PART_ListBox) Then
                        Me.SelectedItems = Me.PART_ListBox.SelectedItems.Cast(Of Item).ToList()
                    End If
                End Sub

            AddHandler Me.PreviewGotKeyboardFocus,
                Sub(s As Object, e As KeyboardFocusChangedEventArgs)
                    If Not _isWorkingFocus Then
                        If (e.OldFocus Is Nothing OrElse Not e.OldFocus.Equals(_oldFocusedElement)) _
                            AndAlso Not _oldFocusedElement Is Nothing AndAlso Not Object.ReferenceEquals(_oldFocusedElement.DataContext, BindingOperations.DisconnectedSource) Then
                            _isWorkingFocus = True
                            _oldFocusedElement.Focus()
                            e.Handled = True
                            _isWorkingFocus = False
                        ElseIf TypeOf e.NewFocus Is ListBoxItem OrElse TypeOf e.NewFocus Is button Then
                            _oldFocusedElement = e.NewFocus
                        End If
                    End If
                End Sub
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.DragViewStrategy = New HorizontalDragViewStrategy(Me.PART_DragInsertIndicator, Me)
        End Sub

        Public Overrides Sub SetSelectedItemsSoft(items As IEnumerable(Of Item))
            SharedSelectedItemsBehavior.SelectItems(items)
        End Sub

        Protected Overrides Sub ToggleCheckBox(checkBox As CheckBox)
            SharedSelectedItemsBehavior.IsSetInternally = True
            MyBase.ToggleCheckBox(checkBox)
            SharedSelectedItemsBehavior.IsSetInternally = False
        End Sub

        Public Overridable ReadOnly Property TilesGroupName As String
            Get
                Return "Quick launch"
            End Get
        End Property

        Protected Overrides Sub OnRequestBringIntoView(s As Object, e As RequestBringIntoViewEventArgs)
            If TypeOf e.OriginalSource Is ListBoxItem AndAlso UIHelper.IsAncestor(Me.PART_ListBox, e.OriginalSource) Then
                Dim item As ListBoxItem = e.OriginalSource
                If Not item Is Nothing AndAlso UIHelper.IsAncestor(PART_ScrollViewer, item) Then
                    Dim transform As GeneralTransform = item.TransformToAncestor(PART_ScrollViewer)
                    Dim itemRect As Rect = transform.TransformBounds(New Rect(0, 0, item.ActualWidth, item.ActualHeight))

                    ' Check if item is outside the viewport and adjust scrolling, but only vertically
                    If itemRect.Top < 0 Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + itemRect.Top - 2)
                    ElseIf itemRect.Bottom > PART_ScrollViewer.ViewportHeight Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + (itemRect.Bottom - PART_ScrollViewer.ViewportHeight) + 2)
                    End If
                    If itemRect.Left < 0 Then
                        PART_ScrollViewer.ScrollToHorizontalOffset(PART_ScrollViewer.HorizontalOffset + itemRect.Left)
                    ElseIf itemRect.Right > PART_ScrollViewer.ViewportWidth Then
                        PART_ScrollViewer.ScrollToHorizontalOffset(PART_ScrollViewer.HorizontalOffset + (itemRect.Right - PART_ScrollViewer.ViewportWidth))
                    End If
                    e.Handled = True
                End If
            ElseIf Not If(TypeOf e.OriginalSource Is Expander, e.OriginalSource, UIHelper.GetParentOfType(Of Expander)(e.OriginalSource)) Is Nothing _
                AndAlso UIHelper.GetParentOfType(Of ListBox)(e.OriginalSource)?.Equals(Me.PART_ListBox) Then
                e.Handled = True
            End If
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
            textAlignment = textAlignment.Left
            fontSize = textBlock.FontSize
        End Sub
    End Class
End Namespace