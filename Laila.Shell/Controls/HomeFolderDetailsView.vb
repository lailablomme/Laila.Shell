Imports System.Windows.Controls
Imports Laila.Shell.Controls.Parts

Namespace Controls
    Public Class HomeFolderDetailsView
        Inherits HomeFolderTilesView

        Protected PART_DragInsertIndicatorVertical As Grid

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_DragInsertIndicatorVertical = Template.FindName("PART_DragInsertIndicatorVertical", Me)

            Me.DragViewStrategy = New VerticalDragViewStrategy(Me.PART_DragInsertIndicatorVertical, Me)
        End Sub

        Public Overrides ReadOnly Property TilesGroupName As String
            Get
                Return String.Empty
            End Get
        End Property
    End Class
End Namespace