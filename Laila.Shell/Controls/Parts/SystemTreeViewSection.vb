Namespace Controls.Parts
    Public Class SystemTreeViewSection
        Inherits BaseTreeViewSection

        Friend Overrides Sub Initialize()
            If Not Settings.IsWindows8_1OrLower Then
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Home) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Home).Clone())
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Gallery) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Gallery).Clone())
            Else
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Favorites) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Favorites).Clone())
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Libraries) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Libraries).Clone())
            End If
        End Sub
    End Class
End Namespace