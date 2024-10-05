'Imports System.Collections.ObjectModel
'Imports System.Runtime.InteropServices
'Imports System.Threading
'Imports System.Windows.Data
'Imports Laila.Shell.Data
'Imports Laila.Shell.Helpers

'Public Class TreeViewFolder
'    Inherits Folder
'    Implements ITreeViewItemData

'    Private _isSelected As Boolean
'    Private _isExpanded As Boolean
'    Private _parent As ITreeViewItemData
'    Friend _folders As ObservableCollection(Of TreeViewFolder)
'    Private _isLoading As Boolean
'    Private _fromThread As Boolean
'    Friend _foldersLock As Object = New Object()

'    Public Overloads Shared Function FromParsingName(parsingName As String, logicalParent As Folder) As TreeViewFolder
'        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
'        If Not shellItem2 Is Nothing Then
'            Dim attr As Integer = SFGAO.FOLDER
'            shellItem2.GetAttributes(attr, attr)
'            If CBool(attr And SFGAO.FOLDER) Then
'                Return New TreeViewFolder(shellItem2, logicalParent)
'            Else
'                Throw New InvalidOperationException("Only folders.")
'            End If
'        Else
'            Return Nothing
'        End If
'    End Function

'    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder)
'        MyBase.New(shellItem2, logicalParent, Nothing)
'    End Sub

'    Public Property IsSelected As Boolean Implements ITreeViewItemData.IsSelected
'        Get
'            Return _isSelected
'        End Get
'        Set(value As Boolean)
'            SetValue(_isSelected, value)
'        End Set
'    End Property

'    Public Property IsExpanded As Boolean Implements ITreeViewItemData.IsExpanded
'        Get
'            Return _isExpanded
'        End Get
'        Set(value As Boolean)
'            SetValue(_isExpanded, value)
'            If value AndAlso (_folders Is Nothing OrElse (_folders.Count = 1 AndAlso TypeOf _folders(0) Is DummyFolder)) Then
'                _folders = Nothing
'                NotifyOfPropertyChange("FoldersThreaded")
'            End If
'        End Set
'    End Property

'    Private Property ITreeViewItemData_Parent As ITreeViewItemData Implements ITreeViewItemData.Parent
'        Get
'            Return _parent
'        End Get
'        Set(value As ITreeViewItemData)
'            SetValue(_parent, value)
'        End Set
'    End Property

'    Public Property IsLoading As Boolean
'        Get
'            Return _isLoading
'        End Get
'        Set(value As Boolean)
'            SetValue(_isLoading, value)
'        End Set
'    End Property

'    Public Overrides ReadOnly Property ItemsThreaded As ObservableCollection(Of Item)
'        Get
'            Return Nothing
'        End Get
'    End Property

'    Public Overrides Property Items As ObservableCollection(Of Item)
'        Get
'            Return Nothing
'        End Get
'        Set(value As ObservableCollection(Of Item))
'            MyBase.Items = Nothing
'        End Set
'    End Property

'    Public Overridable ReadOnly Property FoldersThreaded As ObservableCollection(Of TreeViewFolder)
'        Get
'            If _folders Is Nothing AndAlso (Me.IsExpanded OrElse Me.ITreeViewItemData_Parent Is Nothing OrElse Me.ITreeViewItemData_Parent.IsExpanded) Then
'                Dim result As ObservableCollection(Of TreeViewFolder) = New ObservableCollection(Of TreeViewFolder)()
'                BindingOperations.EnableCollectionSynchronization(result, _foldersLock)

'                Dim t As Thread = New Thread(New ThreadStart(
'                    Sub()
'                        If _folders Is Nothing Then
'                            _fromThread = True
'                            updateFolders(result, False)
'                            Me.Folders = result
'                            _fromThread = False
'                        End If
'                    End Sub))

'                t.Start()

'                Return Nothing
'            Else
'                Return _folders
'            End If
'        End Get
'    End Property

'    Public Overridable Property Folders As ObservableCollection(Of TreeViewFolder)
'        Get
'            If _folders Is Nothing Then
'                Dim result As ObservableCollection(Of TreeViewFolder) = New ObservableCollection(Of TreeViewFolder)()
'                BindingOperations.EnableCollectionSynchronization(result, _foldersLock)
'                updateFolders(result, False)
'                Me.Folders = result
'            End If

'            Return _folders
'        End Get
'        Friend Set(value As ObservableCollection(Of TreeViewFolder))
'            SetValue(_folders, value)
'            Me.NotifyOfPropertyChange("FoldersThreaded")
'        End Set
'    End Property

'    Private Sub updateFolders(items As ObservableCollection(Of TreeViewFolder), isOnUIThread As Boolean)
'        If Me.IsExpanded OrElse Not _fromThread Then
'            Me.IsLoading = True
'            UIHelper.OnUIThread(
'                Sub()
'                End Sub, Windows.Threading.DispatcherPriority.ContextIdle)
'        End If

'        updateItems(SHCONTF.FOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, Me.IsExpanded OrElse Not _fromThread,
'                    Function(item As Item) As Boolean
'                        Return Not items.FirstOrDefault(Function(i) i.FullPath = item.FullPath AndAlso Not i.disposedValue) Is Nothing
'                    End Function,
'                    Sub(item As TreeViewFolder)
'                        items.Add(item)
'                    End Sub,
'                    Sub(item As TreeViewFolder)
'                        items.Remove(item)
'                    End Sub,
'                    Function(paths As List(Of String)) As List(Of Item)
'                        Return items.Where(Function(i) Not paths.Contains(i.FullPath) _
'                                               AndAlso Not TypeOf i Is DummyFolder).Cast(Of Item).ToList()
'                    End Function,
'                    Function(shellItem2 As IShellItem2)
'                        Return New TreeViewFolder(shellItem2, Me)
'                    End Function,
'                    Function(shellItem2 As IShellItem2)
'                        Return New Item(shellItem2, Me)
'                    End Function,
'                    Sub(path As String)
'                        Dim item As Item = items.FirstOrDefault(Function(i) i.FullPath = path AndAlso Not i.disposedValue)
'                        If Not item Is Nothing Then
'                            item._shellItem2.Update(IntPtr.Zero)
'                            For Each prop In item.GetType().GetProperties()
'                                item.NotifyOfPropertyChange(prop.Name)
'                            Next
'                        End If
'                    End Sub,
'                    Sub()
'                        items.Clear()
'                        items.Add(New DummyFolder("Loading...", Nothing))
'                    End Sub,
'                    0, isOnUIThread)

'        For Each item In items
'            item.ITreeViewItemData_Parent = Me
'        Next

'        Me.IsLoading = False
'    End Sub

'    Protected Overrides Async Sub shell_Notification(sender As Object, e As NotificationEventArgs)
'        Dim t As Func(Of Task) =
'            Async Function() As Task
'                If Not _shellItem2 Is Nothing AndAlso Not disposedValue Then
'                    Select Case e.Event
'                        Case SHCNE.MKDIR, SHCNE.CREATE
'                            Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
'                            If Not item1 Is Nothing Then
'                                Dim attr As SFGAO = SFGAO.FOLDER
'                                item1.GetAttributes(attr, attr)
'                                If attr.HasFlag(SFGAO.FOLDER) Then
'                                    Dim parentShellItem2 As IShellItem2
'                                    Try
'                                        item1.GetParent(parentShellItem2)
'                                        Dim parentFullPath As String
'                                        parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
'                                        If Me.FullPath.Equals(parentFullPath) Then
'                                            SyncLock _itemsLock
'                                                UIHelper.OnUIThread(
'                                                    Sub()
'                                                        If Not _folders Is Nothing AndAlso _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue) Is Nothing Then
'                                                            _folders.Add(New TreeViewFolder(item1, Me))
'                                                        End If
'                                                    End Sub)
'                                            End SyncLock
'                                        End If
'                                    Finally
'                                        If Not parentShellItem2 Is Nothing Then
'                                            Marshal.ReleaseComObject(parentShellItem2)
'                                        End If
'                                    End Try
'                                End If
'                            End If
'                        Case SHCNE.RMDIR, SHCNE.DELETE
'                            SyncLock _itemsLock
'                                UIHelper.OnUIThread(
'                                    Sub()
'                                        If Not _folders Is Nothing Then
'                                            Dim item As Item = _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue)
'                                            If Not item Is Nothing Then
'                                                _folders.Remove(item)
'                                            End If
'                                        End If
'                                    End Sub)
'                            End SyncLock
'                        Case SHCNE.DRIVEADD
'                            If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
'                                SyncLock _itemsLock
'                                    UIHelper.OnUIThread(
'                                        Sub()
'                                            If Not _folders Is Nothing AndAlso _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue) Is Nothing Then
'                                                Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
'                                                If Not item1 Is Nothing Then
'                                                    _folders.Add(New TreeViewFolder(item1, Me))
'                                                End If
'                                            End If
'                                        End Sub)
'                                End SyncLock
'                            End If
'                        Case SHCNE.DRIVEREMOVED
'                            If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
'                                SyncLock _itemsLock
'                                    UIHelper.OnUIThread(
'                                        Sub()
'                                            Dim item As Item = _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue)
'                                            If Not _folders Is Nothing Then
'                                                _folders.Remove(item)
'                                            End If
'                                        End Sub)
'                                End SyncLock
'                            End If
'                        Case SHCNE.UPDATEDIR
'                            If (Me.FullPath.Equals(e.Item1Path) OrElse Shell.Desktop.FullPath.Equals(e.Item1Path)) AndAlso Not _folders Is Nothing Then
'                                _fromThread = True
'                                updateFolders(_folders, True)
'                                _fromThread = False
'                            End If
'                    End Select
'                End If
'            End Function

'        Await Task.Run(t)
'    End Sub

'    Protected Overrides Sub Dispose(disposing As Boolean)
'        MyBase.Dispose(disposing)
'    End Sub
'End Class
