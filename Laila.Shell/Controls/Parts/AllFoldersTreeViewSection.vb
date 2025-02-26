Namespace Controls.Parts
    Public Class AllFoldersTreeViewSection
        Inherits BaseTreeViewSection

        Friend Overrides Sub Initialize()
            Me.Items.Add(Shell.Desktop.Clone())
        End Sub
    End Class
End Namespace