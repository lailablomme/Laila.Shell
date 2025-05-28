Namespace Controls.Parts
    Public Class SystemTreeViewSection
        Inherits BaseTreeViewSection

        Friend Overrides Sub Initialize()
            If Not Helpers.OSVersionHelper.IsWindows81OrLower Then
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Home) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Home).Clone())
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Gallery) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Gallery).Clone())
            Else
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Favorites) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Favorites).Clone())
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Libraries) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Libraries).Clone())
            End If

            For Each item In Shell.GetSpecialFolders() _
                .Where(Function(f) Shell.PrivilegedCloudProviders.Contains(f.Key)) _
                .OrderBy(Function(f) f.Value.DisplayName)
                Me.Items.Add(item.Value.Clone())
            Next
        End Sub
    End Class
End Namespace