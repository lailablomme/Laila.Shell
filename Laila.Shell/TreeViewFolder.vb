Imports System.Collections.ObjectModel
Imports System.Threading
Imports Laila.Shell.Data
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.ComponentModel
Imports System.Windows.Data
Imports Laila.Shell.Helpers

Public Class TreeViewFolder
    Inherits Folder
    Implements ITreeViewItemData

    Private _isSelected As Boolean
    Private _isExpanded As Boolean
    Private _parent As ITreeViewItemData
    Friend _folders As ObservableCollection(Of TreeViewFolder)
    Private _isLoading As Boolean
    Private _fromThread As Boolean
    Friend _foldersLock As Object = New Object()

    Public Overloads Shared Function FromParsingName(parsingName As String, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean), list As IList) As TreeViewFolder
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
        If Not shellItem2 Is Nothing Then
            Dim attr As Integer = SFGAO.FOLDER
            shellItem2.GetAttributes(attr, attr)
            If CBool(attr And SFGAO.FOLDER) Then
                Return New TreeViewFolder(shellItem2, logicalParent, setIsLoadingAction, list)
            Else
                Throw New InvalidOperationException("Only folders.")
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean), list As IList)
        MyBase.New(shellItem2, logicalParent, setIsLoadingAction, list)
    End Sub

    Public Property IsSelected As Boolean Implements ITreeViewItemData.IsSelected
        Get
            Return _isSelected
        End Get
        Set(value As Boolean)
            SetValue(_isSelected, value)
        End Set
    End Property

    Public Property IsExpanded As Boolean Implements ITreeViewItemData.IsExpanded
        Get
            Return _isExpanded
        End Get
        Set(value As Boolean)
            SetValue(_isExpanded, value)
            If value AndAlso (_folders Is Nothing OrElse (_folders.Count = 1 AndAlso TypeOf _folders(0) Is DummyTreeViewFolder)) Then
                _folders = Nothing
                NotifyOfPropertyChange("FoldersThreaded")
            End If
        End Set
    End Property

    Private Property ITreeViewItemData_Parent As ITreeViewItemData Implements ITreeViewItemData.Parent
        Get
            Return _parent
        End Get
        Set(value As ITreeViewItemData)
            SetValue(_parent, value)
        End Set
    End Property

    Public Property IsLoading As Boolean
        Get
            Return _isLoading
        End Get
        Set(value As Boolean)
            SetValue(_isLoading, value)
        End Set
    End Property

    Public Overrides ReadOnly Property ItemsThreaded As ObservableCollection(Of Item)
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides Property Items As ObservableCollection(Of Item)
        Get
            Return Nothing
        End Get
        Set(value As ObservableCollection(Of Item))
            MyBase.Items = Nothing
        End Set
    End Property

    Public Overridable ReadOnly Property FoldersThreaded As ObservableCollection(Of TreeViewFolder)
        Get
            If _folders Is Nothing AndAlso (Me.IsExpanded OrElse Me.ITreeViewItemData_Parent Is Nothing OrElse Me.ITreeViewItemData_Parent.IsExpanded) Then
                If Not _setIsLoadingAction Is Nothing Then
                    _setIsLoadingAction(True)
                End If

                Dim t As Thread = New Thread(New ThreadStart(
                    Sub()
                        Dim fp As String = Me.FullPath
                        If String.IsNullOrWhiteSpace(fp) Then
                            Dim i As Int16 = 9
                        End If
                        Debug.WriteLine(fp)
                        SyncLock _foldersLock
                            _fromThread = True
                            Try
                                Dim result As ObservableCollection(Of TreeViewFolder) = Me.Folders
                            Finally
                                _fromThread = False
                            End Try
                        End SyncLock

                        If Not _setIsLoadingAction Is Nothing Then
                            _setIsLoadingAction(False)
                        End If
                    End Sub))

                t.Start()

                Return New ObservableCollection(Of TreeViewFolder)() From {
                    New DummyTreeViewFolder("Loading...", Nothing)
                }
            Else
                Return _folders
            End If
        End Get
    End Property

    Public Overridable Property Folders As ObservableCollection(Of TreeViewFolder)
        Get
            SyncLock _foldersLock
                If _folders Is Nothing Then
                    Dim result As ObservableCollection(Of TreeViewFolder) = New ObservableCollection(Of TreeViewFolder)()
                    updateFolders(result)
                    Me.Folders = result
                    Me.NotifyOfPropertyChange("FoldersThreaded")
                End If

                Return _folders
            End SyncLock
        End Get
        Friend Set(value As ObservableCollection(Of TreeViewFolder))
            SetValue(_folders, value)
        End Set
    End Property

    Private Sub updateFolders(items As ObservableCollection(Of TreeViewFolder))
        updateItems(SHCONTF.FOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, Me.IsExpanded OrElse Not _fromThread,
                    Function(item As Item) As Boolean
                        Return Not items.FirstOrDefault(Function(i) i.FullPath = item.FullPath) Is Nothing
                    End Function,
                    Sub(item As Item)
                        items.Add(item)
                    End Sub,
                    Sub(item As Item)
                        items.Remove(item)
                    End Sub,
                    Function(paths As List(Of String)) As List(Of Item)
                        Return items.Where(Function(i) Not paths.Contains(i.FullPath)).Cast(Of Item).ToList()
                    End Function,
                    Function(shellItem2 As IShellItem2)
                        Return New TreeViewFolder(shellItem2, Me, _setIsLoadingAction, items)
                    End Function,
                    Function(shellItem2 As IShellItem2)
                        Return New Item(shellItem2, Me, _setIsLoadingAction, items)
                    End Function,
                    Sub(path As String)
                        Dim item As Item = items.FirstOrDefault(Function(i) i.FullPath = path)
                        If Not item Is Nothing Then
                            item._shellItem2.Update(IntPtr.Zero)
                            For Each prop In item.GetType().GetProperties()
                                item.NotifyOfPropertyChange(prop.Name)
                            Next
                        End If
                    End Sub,
                    8)

        For Each item In items
            item.ITreeViewItemData_Parent = Me
        Next

        If Not Me.IsExpanded AndAlso items.Count > 0 AndAlso _fromThread Then
            UIHelper.OnUIThread(
                Sub()
                    items.Clear()
                    items.Add(New DummyTreeViewFolder("Loading...", Nothing))
                End Sub)
        End If
    End Sub

    Protected Overrides Async Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        Dim t As Func(Of Task) =
            Async Function() As Task
                If Not _shellItem2 Is Nothing Then
                    Select Case e.Event
                        Case SHCNE.MKDIR, SHCNE.CREATE
                            Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                            If Not item1 Is Nothing Then
                                Dim attr As SFGAO = SFGAO.FOLDER
                                item1.GetAttributes(attr, attr)
                                If attr.HasFlag(SFGAO.FOLDER) Then
                                    Dim parentShellItem2 As IShellItem2
                                    Try
                                        item1.GetParent(parentShellItem2)
                                        Dim parentFullPath As String
                                        parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                                        If Me.FullPath.Equals(parentFullPath) Then
                                            SyncLock _foldersLock
                                                If Not _folders Is Nothing AndAlso _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                                                    UIHelper.OnUIThread(
                                                    Sub()
                                                        _folders.Add(New TreeViewFolder(item1, Me, _setIsLoadingAction, _folders))
                                                    End Sub)
                                                End If
                                            End SyncLock
                                        End If
                                    Finally
                                        If Not parentShellItem2 Is Nothing Then
                                            Marshal.ReleaseComObject(parentShellItem2)
                                        End If
                                    End Try
                                End If
                            End If
                        Case SHCNE.RMDIR, SHCNE.DELETE
                            SyncLock _foldersLock
                                If Not _folders Is Nothing Then
                                    Dim item As Item = _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path)
                                    If Not item Is Nothing Then
                                        UIHelper.OnUIThread(
                                            Sub()
                                                _folders.Remove(item)
                                            End Sub)
                                    End If
                                End If
                            End SyncLock
                        Case SHCNE.DRIVEADD
                            If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
                                SyncLock _foldersLock
                                    If Not _folders Is Nothing AndAlso _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                                        Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                                        If Not item1 Is Nothing Then
                                            UIHelper.OnUIThread(
                                                Sub()
                                                    _folders.Add(New TreeViewFolder(item1, Me, _setIsLoadingAction, _folders))
                                                End Sub)
                                        End If
                                    End If
                                End SyncLock
                            End If
                        Case SHCNE.DRIVEREMOVED
                            If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
                                SyncLock _foldersLock
                                    Dim item As Item = _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path)
                                    If Not _folders Is Nothing Then
                                        UIHelper.OnUIThread(
                                            Sub()
                                                _folders.Remove(item)
                                            End Sub)
                                    End If
                                End SyncLock
                            End If
                        Case SHCNE.UPDATEDIR
                            If (Me.FullPath.Equals(e.Item1Path) OrElse Shell.Desktop.FullPath.Equals(e.Item1Path)) AndAlso Not _folders Is Nothing Then
                                Me.IsLoading = True

                                SyncLock _foldersLock
                                    _fromThread = True
                                    Try
                                        updateFolders(_folders)
                                    Finally
                                        _fromThread = False
                                    End Try
                                End SyncLock

                                Me.IsLoading = False
                            End If
                    End Select
                End If
            End Function

        Await Task.Run(t)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        MyBase.Dispose(disposing)
    End Sub
End Class
