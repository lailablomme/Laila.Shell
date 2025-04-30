Imports System.Windows.Controls

Namespace Controls
    Public Class NewItemMenu
        Inherits BaseContextMenu

        Protected Overrides Function AddItems() As Task
            Dim menuItems As List(Of Control) = GetMenuItems()

            Dim newMenuItem As MenuItem = menuItems _
                .FirstOrDefault(Function(c) TypeOf c Is MenuItem _
                    AndAlso Not c.Tag Is Nothing _
                    AndAlso (CType(c.Tag, Tuple(Of Integer, String, Object)).Item2 = "New" _
                             OrElse CType(c, MenuItem).Items.Cast(Of Control).ToList() _
                                .Exists(Function(sc) TypeOf sc Is MenuItem _
                                    AndAlso CType(sc.Tag, Tuple(Of Integer, String, Object)).Item2 = "NewFolder")))

            If Not newMenuItem Is Nothing Then
                For Each item In newMenuItem.Items.Cast(Of Control).ToList()
                    newMenuItem.Items.Remove(item)
                    Me.Items.Add(item)
                Next
            End If

            Return Task.CompletedTask
        End Function

        Protected Overrides Function DoRenameAfter(Tag As Tuple(Of Integer, String, Object)) As Boolean
            Return True
        End Function
    End Class
End Namespace