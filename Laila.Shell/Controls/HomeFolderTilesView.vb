Imports Laila.Shell.Behaviors
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Helpers
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input

Namespace Controls
    Public Class HomeFolderTilesView
        Inherits BaseFolderView

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(HomeFolderTilesView), New FrameworkPropertyMetadata(GetType(HomeFolderTilesView)))
        End Sub

        Private _oldFocusedElement As UIElement = Nothing
        Private _isWorkingFocus As Boolean

        Protected Overrides Sub PART_ListBox_Loaded()
            MyBase.PART_ListBox_Loaded()

            _selectionHelper.Unhook()
            _selectionHelper = Nothing

            AddHandler Me.PreviewGotKeyboardFocus,
                Sub(s As Object, e As KeyboardFocusChangedEventArgs)
                    If Not _isWorkingFocus Then
                        If (e.OldFocus Is Nothing OrElse Not e.OldFocus.Equals(_oldFocusedElement)) _
                            AndAlso Not _oldFocusedElement Is Nothing Then
                            _isWorkingFocus = True
                            _oldFocusedElement.Focus()
                            e.Handled = True
                            _isWorkingFocus = False
                        Else
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