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

            If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.OneDrive) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.OneDrive).Clone())
            If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.OneDriveBusiness) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.OneDriveBusiness).Clone())
        End Sub
    End Class
End Namespace