Imports System.Windows.Controls

Namespace Controls
    Public Class NewItemMenu
        Inherits BaseMenu

        Protected Overrides Sub AddItems()
            Dim menuItems As List(Of Control) = getMenuItems()

            Dim newMenuItem As MenuItem = menuItems _
                .FirstOrDefault(Function(c) TypeOf c Is MenuItem _
                    AndAlso Not c.Tag Is Nothing _
                    AndAlso CType(c.Tag, Tuple(Of Integer, String)).Item2 = "New")

            If Not newMenuItem Is Nothing Then
                For Each item In newMenuItem.Items.Cast(Of Control).ToList()
                    newMenuItem.Items.Remove(item)
                    Me.Items.Add(item)
                Next
            End If
        End Sub

        Protected Overrides Function DoRenameAfter(Tag As Tuple(Of Integer, String)) As Boolean
            Return True
        End Function
    End Class
End Namespace