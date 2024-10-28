Imports System.Windows
Imports System.Windows.Controls

Namespace Helpers
    Public Class TreeViewItemStyleSelector
        Inherits StyleSelector

        Public Property TreeViewItemStyle As Style
        Public Property TreeViewSeparatorItemStyle As Style
        Public Property TreeViewPlaceholderItemStyle As Style

        Public Overrides Function SelectStyle(item As Object, container As DependencyObject) As Style
            If TypeOf item Is SeparatorFolder Then
                Return Me.TreeViewSeparatorItemStyle
            ElseIf TypeOf item Is PinnedAndFrequentPlaceholderFolder Then
                Return Me.TreeViewPlaceholderItemStyle
            Else
                Return Me.TreeViewItemStyle
            End If
        End Function
    End Class
End Namespace