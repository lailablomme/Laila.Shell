Imports Microsoft.VisualBasic.Devices
Imports Microsoft.Win32

Namespace Controls.Parts
    Public Class EnvironmentTreeViewSection
        Inherits BaseTreeViewSection

        Private Const DESKTOP_SHELL_EXTENSIONS_KEY As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace"
        Private Const CLASS_KEY As String = "CLSID\{0}\ShellFolder"

        Friend Overrides Sub Initialize()
            Me.Items.Add(New SeparatorFolder())

            Dim shellExts As List(Of Item) = New List(Of Item)()
            For Each guid In getShellExtensions(Registry.CurrentUser, DESKTOP_SHELL_EXTENSIONS_KEY)
                Dim item As Item = Item.FromParsingName("shell:::" & guid, Nothing)
                If Not item Is Nothing AndAlso TypeOf item Is Folder _
                    AndAlso Not Shell.GetSpecialFolders().Where(Function(f) Shell.PrivilegedCloudProviders.Contains(f.Key)) _
                        .Any(Function(f) Item.ArePathsEqual(item.FullPath, f.Value?.FullPath)) Then
                    shellExts.Add(item)
                End If
            Next

            For Each item In shellExts.OrderBy(Function(i) i.DisplayName)
                Me.Items.Add(item)
            Next

            If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.ThisPc) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.ThisPc).Clone())
            If Not Helpers.OSVersionHelper.IsWindows7OrLower AndAlso Me.TreeView.DoShowLibrariesInTreeView Then
                If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Libraries) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Libraries).Clone())
            End If
            If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.Network) Then Me.Items.Add(Shell.GetSpecialFolder(SpecialFolders.Network).Clone())
        End Sub

        Private Function getShellExtensions(root As RegistryKey, keyName As String) As List(Of String)
            Dim result As List(Of String) = New List(Of String)()

            Using key As RegistryKey = root.OpenSubKey(keyName)
                If key IsNot Nothing Then
                    For Each subKeyName As String In key.GetSubKeyNames()
                        Using subKey As RegistryKey = Registry.ClassesRoot.OpenSubKey(String.Format(CLASS_KEY, subKeyName))
                            If Not subKey Is Nothing Then
                                result.Add(subKeyName)
                            End If
                        End Using
                    Next
                End If
            End Using

            Return result
        End Function
    End Class
End Namespace