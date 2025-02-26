Imports Microsoft.VisualBasic.Devices

Namespace Controls.Parts
    Public Class EnvironmentTreeViewSection
        Inherits BaseTreeViewSection

        Friend Overrides Sub Initialize()
            Me.Items.Add(New SeparatorFolder())

            If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.ThisPc) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.ThisPc).Clone())
            If Not Settings.IsWindows7OrLower AndAlso Me.TreeView.DoShowLibrariesInTreeView Then
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Libraries) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Libraries).Clone())
            End If
            If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Network) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Network).Clone())
        End Sub
    End Class
End Namespace