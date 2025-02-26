Namespace Controls.Parts
    Public Class AllFoldersTreeViewSection
        Inherits BaseTreeViewSection

        Friend Overrides Sub Initialize()
            Me.Items.Add(Shell.Desktop.Clone())
            CType(Me.Items(0), Folder).IsExpanded = True
        End Sub
    End Class
End Namespace