Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Laila.Shell.ViewModels
Imports Microsoft

Namespace Helpers
    Public Class UIHelper
        Public Shared Function GetScreenSize(visual As Visual, isWorkingArea As Boolean) As Size
            Dim handle As IntPtr = CType(PresentationSource.FromVisual(visual), HwndSource).Handle
            Dim currentScreen As System.Windows.Forms.Screen = System.Windows.Forms.Screen.FromHandle(handle)
            Dim dpi As DpiScale = VisualTreeHelper.GetDpi(visual)
            If isWorkingArea Then
                Return New Size(
                    1.0 * currentScreen.WorkingArea.Width / (dpi.PixelsPerInchX / 96),
                    1.0 * currentScreen.WorkingArea.Height / (dpi.PixelsPerInchY / 96)
                )
            Else
                Return New Size(
                    1.0 * currentScreen.Bounds.Width / (dpi.PixelsPerInchX / 96),
                    1.0 * currentScreen.Bounds.Height / (dpi.PixelsPerInchY / 96)
                )
            End If
        End Function

        Public Shared Sub OnUIThread(action As Action, Optional priority As DispatcherPriority = DispatcherPriority.Normal, Optional doForceInvoke As Boolean = False)
            Dim appl As System.Windows.Application = System.Windows.Application.Current
            If Not appl Is Nothing Then
                If Not doForceInvoke AndAlso appl.Dispatcher.CheckAccess AndAlso priority = DispatcherPriority.Normal Then
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

        Public Shared Function WIN32POINTToUIElement(ptWIN32 As WIN32POINT, element As UIElement) As Point
            Dim pt As Point = New Point(ptWIN32.x, ptWIN32.y)
            Return element.PointFromScreen(pt)
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

        Public Shared Function GetBindingsDeep(obj As DependencyObject) As IEnumerable(Of Binding)
            Dim result As List(Of Binding) = New List(Of Binding)()

            Dim objectBindings As Dictionary(Of DependencyProperty, Binding) = GetBindings(obj)
            result.AddRange(objectBindings.Values.ToList())

            For i = 0 To VisualTreeHelper.GetChildrenCount(obj) - 1
                Dim child As DependencyObject = VisualTreeHelper.GetChild(obj, i)
                Dim childBindings As IEnumerable(Of Binding) = GetBindingsDeep(child)
                result.AddRange(childBindings)
            Next

            Return result
        End Function

        Public Shared Function GetBindings(obj As DependencyObject) As Dictionary(Of DependencyProperty, Binding)
            Dim result As Dictionary(Of DependencyProperty, Binding) = New Dictionary(Of DependencyProperty, Binding)()
            Dim properties As List(Of DependencyProperty) = GetDependencyProperties(obj)
            For Each p In properties
                Dim b As BindingBase = BindingOperations.GetBindingBase(obj, p)
                If Not b Is Nothing AndAlso TypeOf b Is Binding Then
                    result.Add(p, b)
                End If
            Next
            Return result
        End Function

        Public Shared Function GetDependencyProperties(obj As DependencyObject) As List(Of DependencyProperty)
            Dim result As List(Of DependencyProperty) = New List(Of DependencyProperty)()
            For Each fi As FieldInfo In obj.GetType().GetFields(BindingFlags.FlattenHierarchy Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static Or BindingFlags.Public)
                If fi.FieldType = GetType(DependencyProperty) Then
                    result.Add(fi.GetValue(Nothing))
                End If
            Next
            Return result
        End Function
    End Class
End Namespace