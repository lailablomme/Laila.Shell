Imports System.Windows
Imports System.Windows.Controls

Namespace Controls.Parts
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
            ElseIf TypeOf item Is Folder Or TypeOf item Is Item Then
                Return Me.TreeViewItemStyle
            Else
                Return Nothing
            End If
        End Function
    End Class
End Namespace