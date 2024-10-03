Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Laila.Shell.ViewModels
Imports Microsoft

Namespace Helpers
    Public Class UIHelper
        Public Shared Sub OnUIThread(action As Action, Optional priority As DispatcherPriority = DispatcherPriority.Normal)
            Dim appl As System.Windows.Application = System.Windows.Application.Current
            If Not appl Is Nothing Then
                If appl.Dispatcher.CheckAccess AndAlso priority = DispatcherPriority.Normal Then
                    action()
                Else
                    appl.Dispatcher.Invoke(
                        Sub()
                            action()
                        End Sub, priority)
                End If
            End If
        End Sub

        Public Shared Sub OnUIThreadAsync(action As Action, Optional priority As DispatcherPriority = DispatcherPriority.Normal)
            Dim appl As System.Windows.Application = System.Windows.Application.Current
            If Not appl Is Nothing Then
                appl.Dispatcher.BeginInvoke(
                    Sub()
                        action()
                    End Sub, priority)
            End If
        End Sub

        Public Shared Function WIN32POINTToControl(ptWIN32 As WIN32POINT, control As Control) As Point
            Dim pt As Point = New Point(ptWIN32.x, ptWIN32.y)
            Return control.PointFromScreen(pt)
        End Function

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