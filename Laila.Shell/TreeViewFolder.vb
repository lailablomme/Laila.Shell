Imports System.Collections.ObjectModel
Imports System.Threading
Imports Laila.Shell.Data
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.ComponentModel
Imports System.Windows.Data

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

    Public Overloads Shared Function FromParsingName(parsingName As String, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean)) As TreeViewFolder
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
        Dim attr As Integer = SFGAO.FOLDER
        shellItem2.GetAttributes(attr, attr)
        If CBool(attr And SFGAO.FOLDER) Then
            Return New TreeViewFolder(Folder.GetIShellFolderFromIShellItem2(shellItem2), shellItem2, logicalParent, setIsLoadingAction)
        Else
            Throw New InvalidOperationException("Only folders.")
        End If
    End Function

    Public Sub New(shellFolder As IShellFolder, shellItem2 As IShellItem2, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean))
        MyBase.New(shellFolder, shellItem2, logicalParent, setIsLoadingAction)
    End Sub

    Public Sub New(bindingParent As Folder, pidl As IntPtr, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean))
        MyBase.New(bindingParent, pidl, logicalParent, setIsLoadingAction)
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
                        SyncLock _foldersLock
                            _fromThread = True
                            Dim result As ObservableCollection(Of TreeViewFolder) = Me.Folders
                            _fromThread = False
                        End SyncLock

                        If Not _setIsLoadingAction Is Nothing Then
                            _setIsLoadingAction(False)
                        End If
                    End Sub))

                t.Start()

                Return New ObservableCollection(Of TreeViewFolder)() From {
                    New DummyTreeViewFolder("Loading...")
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
        Dim paths As List(Of String) = New List(Of String)

        Dim flags As UInt32 = CType(SHCONTF.FOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, UInt32)

        If Not isWindows7OrLower() Then
            Dim bindCtx As ComTypes.IBindCtx, bindCtxPtr As IntPtr
            Functions.CreateBindCtx(0, bindCtxPtr)
            bindCtx = Marshal.GetTypedObjectForIUnknown(bindCtxPtr, GetType(ComTypes.IBindCtx))

            Dim propertyBag As IPropertyBag, propertyBagPtr As IntPtr
            Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBagPtr)
            propertyBag = Marshal.GetTypedObjectForIUnknown(propertyBagPtr, GetType(IPropertyBag))

            Dim var As New PROPVARIANT()
            var.vt = VarEnum.VT_UI4
            var.union.uintVal = flags
            propertyBag.Write("SHCONTF", var)

            bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag)
            Dim ptr2 As IntPtr
            bindCtxPtr = Marshal.GetIUnknownForObject(bindCtx)

            ShellItem2.BindToHandler(bindCtxPtr, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, ptr2)
            Dim enumShellItems As IEnumShellItems = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IEnumShellItems))

            Try
                Dim shellItemArray(0) As IShellItem, feteched As UInt32 = 1
                Application.Current.Dispatcher.Invoke(
                                Sub()
                                    enumShellItems.Next(1, shellItemArray, feteched)
                                End Sub)
                Dim onlyOnce As Boolean = True
                While feteched = 1 AndAlso (Me.IsExpanded OrElse onlyOnce OrElse Not _fromThread)
                    Dim tvf As TreeViewFolder = New TreeViewFolder(Folder.GetIShellFolderFromIShellItem2(shellItemArray(0)), shellItemArray(0), Me, _setIsLoadingAction)
                    tvf.ITreeViewItemData_Parent = Me
                    Dim cache As String = tvf.DisplayName
                    paths.Add(tvf.FullPath)
                    If items.FirstOrDefault(Function(i) i.FullPath = tvf.FullPath) Is Nothing Then
                        Application.Current.Dispatcher.Invoke(
                            Sub()
                                items.Add(tvf)
                            End Sub)
                    Else
                        tvf.Dispose()
                    End If
                    onlyOnce = False
                    Thread.Sleep(10)
                    Application.Current.Dispatcher.Invoke(
                        Sub()
                            enumShellItems.Next(1, shellItemArray, feteched)
                        End Sub)
                End While
            Catch ex As Exception
                ' TODO handle access denied
            End Try
        Else
            Dim list As IEnumIDList
            Me.ShellFolder.EnumObjects(Nothing, flags, list)
            If Not list Is Nothing Then
                Dim pidl(0) As IntPtr, fetched As Integer
                list.Next(1, pidl, fetched)
                Dim onlyOnce As Boolean = True
                While fetched = 1 AndAlso (Me.IsExpanded OrElse onlyOnce OrElse Not _fromThread)
                    Dim tvf As TreeViewFolder
                    Application.Current.Dispatcher.Invoke(
                        Sub()
                            tvf = New TreeViewFolder(Me, pidl(0), Me, _setIsLoadingAction)
                            Dim cache As String = tvf.DisplayName
                            paths.Add(tvf.FullPath)
                            If items.FirstOrDefault(Function(i) i.FullPath = tvf.FullPath) Is Nothing Then
                                items.Add(tvf)
                            Else
                                tvf.Dispose()
                            End If
                            list.Next(1, pidl, fetched)
                        End Sub)
                    Application.Current.Dispatcher.Invoke(
                        Sub()
                            list.Next(1, pidl, fetched)
                        End Sub)
                    onlyOnce = False
                    Thread.Sleep(10)
                End While
            End If
        End If

        If Not Me.IsExpanded AndAlso items.Count > 0 AndAlso _fromThread Then
            Application.Current.Dispatcher.Invoke(
                Sub()
                    items.Clear()
                    items.Add(New DummyTreeViewFolder("Loading..."))
                End Sub)
        Else
            Dim toBeRemoved As List(Of TreeViewFolder) = items.Where(Function(i) Not paths.Contains(i.FullPath)).ToList()
            For Each item In toBeRemoved
                Application.Current.Dispatcher.Invoke(
                    Sub()
                        items.Remove(item)
                    End Sub)
            Next
        End If
    End Sub

    Protected Overrides Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        Select Case e.Event
            Case SHCNE.MKDIR
                If Not e.Item1 Is Nothing Then
                    Dim parentShellItem2 As IShellItem2
                    e.Item1.GetParent(parentShellItem2)
                    Dim parentFullPath As String
                    parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                    If Me.FullPath.Equals(parentFullPath) Then
                        If Not _folders Is Nothing AndAlso _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                            _folders.Add(New TreeViewFolder(Folder.GetIShellFolderFromIShellItem2(e.Item1), e.Item1, Me, _setIsLoadingAction))
                        End If
                    End If
                End If
            Case SHCNE.RMDIR
                If Not _folders Is Nothing Then
                    Dim item As Item = _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path)
                    If Not item Is Nothing Then
                        _folders.Remove(item)
                    End If
                End If
            Case SHCNE.DRIVEADD
                If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
                    If Not _folders Is Nothing AndAlso _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                        _folders.Add(New TreeViewFolder(Folder.GetIShellFolderFromIShellItem2(e.Item1), e.Item1, Me, _setIsLoadingAction))
                    End If
                End If
            Case SHCNE.DRIVEREMOVED
                If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
                    Dim item As Item = _folders.FirstOrDefault(Function(i) i.FullPath = e.Item1Path)
                    If Not _folders Is Nothing Then
                        _folders.Remove(item)
                    End If
                End If
            Case SHCNE.UPDATEDIR
                SyncLock _foldersLock
                    If (Me.FullPath.Equals(e.Item1Path) OrElse Shell.Desktop.FullPath.Equals(e.Item1Path)) AndAlso Not _folders Is Nothing Then
                        If Not _setIsLoadingAction Is Nothing Then
                            _setIsLoadingAction(True)
                        End If

                        Dim thread As Thread = New Thread(New ThreadStart(
                            Sub()
                                SyncLock _foldersLock
                                    updateFolders(_folders)
                                End SyncLock

                                Application.Current.Dispatcher.Invoke(
                                    Sub()
                                        Dim view As ICollectionView = CollectionViewSource.GetDefaultView(_folders)
                                        view.Refresh()
                                    End Sub)

                                If Not _setIsLoadingAction Is Nothing Then
                                    _setIsLoadingAction(False)
                                End If
                            End Sub))

                        thread.Start()
                    End If
                End SyncLock
        End Select
    End Sub
End Class
