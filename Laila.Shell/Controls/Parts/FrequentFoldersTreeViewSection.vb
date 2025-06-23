Imports System.Threading
Imports System.Windows.Threading
Imports Laila.Shell.Controls.TreeView
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop.Items

Namespace Controls.Parts
    Public Class FrequentFoldersTreeViewSection
        Inherits BaseTreeViewSection
        Implements IProcessNotifications

        Public Property IsProcessingNotifications As Boolean = True Implements IProcessNotifications.IsProcessingNotifications

        Public Property NotificationThreadId As Integer? Implements IProcessNotifications.NotificationThreadId

        Private _timer As Timer = Nothing
        Private _signature As String = String.Empty

        Friend Overrides Sub Initialize()
            updateFrequentFolders()

            If _timer Is Nothing Then
                _timer = New Timer(New TimerCallback(
                    Sub()
                        updateFrequentFolders()
                    End Sub), Nothing, 1000 * 60, 1000 * 60)
            End If

            AddHandler PinnedItems.ItemPinned,
                Sub(s As Object, e As PinnedItemEventArgs)
                    updateFrequentFolders()
                End Sub
            AddHandler PinnedItems.ItemUnpinned,
                Sub(s As Object, e As PinnedItemEventArgs)
                    updateFrequentFolders()
                End Sub

            Shell.SubscribeToNotifications(Me)
        End Sub

        Private Sub updateFrequentFolders()
            Dim folders As IEnumerable(Of Folder) = FrequentFolders.GetMostFrequent()
            Dim sig As String = String.Join(";", folders.Select(Function(f) If(f.FullPath, "?")))
            If Not _signature.Equals(sig) Then ' avoid constant flicker
                _signature = sig

                UIHelper.OnUIThread(
                    Sub()
                        Dim selectedPidl As Pidl = Nothing
                        If Me.Items.Contains(Me.TreeView.SelectedItem) Then
                            selectedPidl = Me.TreeView.SelectedItem.Pidl?.Clone()
                        End If

                        For i = 0 To folders.Count - 1
                            If i < Me.Items.Count Then
                                If Not If(Me.Items(i).Pidl?.Equals(folders(i).Pidl), False) Then
                                    If i + 1 < Me.Items.Count AndAlso If(Me.Items(i + 1).Pidl?.Equals(folders(i).Pidl), False) Then
                                        Me.Items.Remove(Me.Items(i))
                                    Else
                                        Me.Items.Insert(i, folders(i))
                                    End If
                                    If folders(i).Pidl?.Equals(selectedPidl) Then Me.TreeView.SetSelectedItem(folders(i))
                                End If
                            Else
                                Me.Items.Add(folders(i))
                            End If
                        Next
                        For i = folders.Count To Me.Items.Count - 1
                            Me.Items.Remove(Me.Items.Last())
                        Next

                        If Not selectedPidl Is Nothing Then
                            selectedPidl.Dispose()
                        End If
                    End Sub)
            End If
        End Sub

        Protected Friend Overridable Sub ProcessNotification(e As NotificationEventArgs) Implements IProcessNotifications.ProcessNotification
            Dim mustUpdate As Boolean = False

            Select Case e.Event
                Case SHCNE.RMDIR, SHCNE.DELETE
                    mustUpdate = Not Me.Items.ToList().FirstOrDefault(Function(i) _
                        Not i.disposedValue _
                        AndAlso Not i.FullPath Is Nothing _
                        AndAlso Item.ArePathsEqual(i.FullPath, e.Item1.FullPath)) Is Nothing
                Case SHCNE.UPDATEDIR
                    mustUpdate = (Not Me.Items.ToList().FirstOrDefault(
                            Function(i)
                                If Not i.disposedValue Then
                                    Return (Not i.Parent Is Nothing _
                                        AndAlso Not i.Parent.FullPath Is Nothing _
                                        AndAlso Item.ArePathsEqual(i.Parent.FullPath, e.Item1.FullPath))
                                Else
                                    Return False
                                End If
                            End Function) Is Nothing _
                        OrElse Item.ArePathsEqual(Shell.Desktop.FullPath, e.Item1.FullPath))
            End Select

            If mustUpdate Then
                Shell.GlobalThreadPool.Add(
                    Sub()
                        updateFrequentFolders()
                    End Sub)
            End If
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)

            _timer.Dispose()
        End Sub
    End Class
End Namespace