Imports System.Windows
Imports System.Windows.Media

Namespace Helpers
    Public Class UIHelper
        Public Shared Function GetParent(obj As DependencyObject) As DependencyObject
            If obj Is Nothing Then
                Return Nothing
            End If

            If TypeOf obj Is ContentElement Then
                Dim parent As DependencyObject = ContentOperations.GetParent(obj)
                If Not parent Is Nothing Then
                    Return parent
                End If
                If TypeOf obj Is FrameworkContentElement Then
                    Return CType(obj, FrameworkContentElement).Parent
                Else
                    Return Nothing
                End If
            End If
            Dim lp As DependencyObject = LogicalTreeHelper.GetParent(obj)
            Dim vp As DependencyObject = VisualTreeHelper.GetParent(obj)
            Return If(lp Is Nothing, vp, lp)
        End Function

        Public Shared Function GetParentOfType(Of T As DependencyObject)(child As DependencyObject) As T
            Dim parent = GetParent(child)
            If Not parent Is Nothing AndAlso Not TypeOf parent Is T Then
                Return GetParentOfType(Of T)(parent)
            End If
            Return parent
        End Function

        Public Shared Iterator Function FindVisualChildren(Of T As DependencyObject)(depObj As DependencyObject, Optional isDeep As Boolean = True) As IEnumerable(Of T)
            For i = 0 To VisualTreeHelper.GetChildrenCount(depObj) - 1
                Dim isMatch As Boolean = False
                Dim child As DependencyObject = VisualTreeHelper.GetChild(depObj, i)
                If Not child Is Nothing AndAlso TypeOf child Is T Then
                    isMatch = True
                    Yield child
                End If

                If Not isMatch Or isDeep Then
                    For Each childOfChild In FindVisualChildren(Of T)(child, isDeep)
                        Yield childOfChild
                    Next
                End If
            Next
        End Function
    End Class
End Namespace