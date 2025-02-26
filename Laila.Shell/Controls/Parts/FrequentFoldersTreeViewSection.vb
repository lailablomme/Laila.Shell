Imports System.Windows.Threading
Imports Laila.Shell.Controls.TreeView
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop.Items

Namespace Controls.Parts
    Public Class FrequentFoldersTreeViewSection
        Inherits BaseTreeViewSection

        Private _timer As DispatcherTimer = Nothing
        Private _signature As String = String.Empty

        Friend Overrides Sub Initialize()
            updateFrequentFolders()

            If _timer Is Nothing Then
                _timer = New DispatcherTimer()
                AddHandler _timer.Tick,
                    Sub(s As Object, e As EventArgs)
                        updateFrequentFolders()
                    End Sub
                _timer.Interval = TimeSpan.FromMinutes(1)
                _timer.IsEnabled = True
            End If

            AddHandler PinnedItems.ItemPinned,
                Sub(s As Object, e As PinnedItemEventArgs)
                    updateFrequentFolders()
                End Sub
            AddHandler PinnedItems.ItemUnpinned,
                Sub(s As Object, e As PinnedItemEventArgs)
                    updateFrequentFolders()
                End Sub

            AddHandler Shell.Notification,
                Sub(s As Object, e As NotificationEventArgs)
                    Select Case e.Event
                        Case SHCNE.RMDIR, SHCNE.DELETE
                            UIHelper.OnUIThread(
                                Sub()
                                    If Not Me.Items.FirstOrDefault(Function(i) _
                                        Not i.disposedValue _
                                        AndAlso Not i.FullPath Is Nothing _
                                        AndAlso i.FullPath.Equals(e.Item1.FullPath)) Is Nothing Then
                                        updateFrequentFolders()
                                    End If
                                End Sub)
                        Case SHCNE.UPDATEDIR
                            UIHelper.OnUIThread(
                                Sub()
                                    If (Not Me.Items.FirstOrDefault(
                                        Function(i)
                                            If Not i.disposedValue Then
                                                Return (Not i.Parent Is Nothing _
                                                    AndAlso Not i.Parent.FullPath Is Nothing _
                                                    AndAlso i.Parent.FullPath.Equals(e.Item1.FullPath))
                                            Else
                                                Return False
                                            End If
                                        End Function) Is Nothing _
                                    OrElse Shell.Desktop.FullPath.Equals(e.Item1.FullPath)) Then
                                        updateFrequentFolders()
                                    End If
                                End Sub)
                    End Select
                End Sub
        End Sub

        Private Sub updateFrequentFolders()
            Dim folders As IEnumerable(Of Folder) = FrequentFolders.GetMostFrequent()
            Dim sig As String = String.Join(";", folders.Select(Function(f) If(f.FullPath, "?")))
            If Not _signature.Equals(sig) Then ' avoid constant flicker
                _signature = sig

                Dim selectedPidl As Pidl = Nothing
                If Me.Items.Contains(Me.TreeView.SelectedItem) Then
                    selectedPidl = Me.TreeView.SelectedItem.Pidl.Clone()
                End If

                For Each item In Me.Items.ToList()
                    Me.Items.Remove(item)
                Next

                For Each folder In folders
                    Me.Items.Add(folder)
                    If folder.Pidl.Equals(selectedPidl) Then Me.TreeView.SetSelectedItem(folder)
                Next

                If Not selectedPidl Is Nothing Then
                    selectedPidl.Dispose()
                End If
            End If
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)

            _timer.IsEnabled = False
        End Sub
    End Class
End Namespace