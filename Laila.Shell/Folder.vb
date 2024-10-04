﻿Imports System.Collections.ObjectModel
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Data
Imports Laila.Shell.Helpers

Public Class Folder
    Inherits Item

    Private _columns As List(Of Column)
    Private _items As ObservableCollection(Of Item)
    Protected _itemsLock As Object = New Object()
    Friend _shellFolder As IShellFolder
    Private _columnManager As IColumnManager
    Protected _setIsLoadingAction As Action(Of Boolean)

    Public Shared Function FromKnownFolderGuid(knownFolderGuid As Guid, setIsLoadingAction As Action(Of Boolean), cachedIconSize As Integer) As Folder
        Return FromParsingName("shell:::" & knownFolderGuid.ToString("B"), Nothing, setIsLoadingAction, cachedIconSize)
    End Function

    Friend Shared Function GetIShellFolderFromIShellItem2(shellItem2 As IShellItem2) As IShellFolder
        Dim ptr2 As IntPtr
        shellItem2.BindToHandler(Nothing, Guids.BHID_SFObject, GetType(IShellFolder).GUID, ptr2)
        Return Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IShellFolder))
    End Function

    Friend Shared Function GetIShellFolderFromPidl(pidl As IntPtr, bindingParent As Folder) As IShellFolder
        Dim ptr As IntPtr
        Try
            bindingParent._shellFolder.BindToObject(pidl, IntPtr.Zero, Guids.IID_IShellFolder, ptr)
            Return Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellFolder))
        Finally
            Marshal.Release(ptr)
        End Try
    End Function

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean), cachedIconSize As Integer)
        MyBase.New(shellItem2, logicalParent, cachedIconSize)

        If Not shellItem2 Is Nothing Then
            _shellFolder = Folder.GetIShellFolderFromIShellItem2(shellItem2)
        End If

        _setIsLoadingAction = setIsLoadingAction
    End Sub

    Public ReadOnly Property Columns(canonicalName As String) As Column
        Get
            Return Me.Columns.SingleOrDefault(Function(c) canonicalName.Equals(c.CanonicalName))
        End Get
    End Property

    Public ReadOnly Property Columns As IEnumerable(Of Column)
        Get
            If _columns Is Nothing Then
                _columns = New List(Of Column)()

                Dim ptr As IntPtr
                _shellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, ptr)
                Dim shellView As IShellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))
                _columnManager = shellView
                Dim count As Integer
                _columnManager.GetColumnCount(CM_ENUM_FLAGS.CM_ENUM_ALL, count)
                Dim propertyKeys(count - 1) As PROPERTYKEY
                _columnManager.GetColumns(CM_ENUM_FLAGS.CM_ENUM_ALL, propertyKeys, count)
                Dim index As Integer = 0
                For Each propertyKey In propertyKeys
                    Dim col As Column = New Column(propertyKey, _columnManager, index)
                    If Not col._propertyDescription Is Nothing Then
                        _columns.Add(col)
                        index += 1
                    End If
                Next
            End If

            Return _columns
        End Get
    End Property

    Public Overridable ReadOnly Property ItemsThreaded As ObservableCollection(Of Item)
        Get
            If _items Is Nothing Then
                Dim result As ObservableCollection(Of Item) = New ObservableCollection(Of Item)()
                BindingOperations.EnableCollectionSynchronization(result, _itemsLock)

                If Not _setIsLoadingAction Is Nothing Then
                    _setIsLoadingAction(True)
                End If

                Dim t As Thread = New Thread(New ThreadStart(
                    Sub()
                        SyncLock _itemsLock
                            If _items Is Nothing Then
                                updateItems(result, False)
                                Me.Items = result
                            End If
                        End SyncLock

                        If Not _setIsLoadingAction Is Nothing Then
                            _setIsLoadingAction(False)
                        End If
                    End Sub))

                t.Start()

                Return Nothing
            Else
                Return _items
            End If
        End Get
    End Property

    Public Overridable Property Items As ObservableCollection(Of Item)
        Get
            SyncLock _itemsLock
                If _items Is Nothing Then
                    Dim result As ObservableCollection(Of Item) = New ObservableCollection(Of Item)()
                    BindingOperations.EnableCollectionSynchronization(result, _itemsLock)
                    updateItems(result, False)
                    Me.Items = result
                End If

                Return _items
            End SyncLock
        End Get
        Set(value As ObservableCollection(Of Item))
            SetValue(_items, value)
            Me.NotifyOfPropertyChange("ItemsThreaded")
        End Set
    End Property

    Protected Sub updateItems(items As ObservableCollection(Of Item), isOnUIThread As Boolean)
        updateItems(SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, True,
                    Function(item As Item) As Boolean
                        Return Not items.FirstOrDefault(Function(i) i.FullPath = item.FullPath AndAlso Not i.disposedValue) Is Nothing
                    End Function,
                    Sub(item As Item)
                        items.Add(item)
                    End Sub,
                    Sub(item As Item)
                        items.Remove(item)
                    End Sub,
                    Function(paths As List(Of String)) As List(Of Item)
                        Return items.Where(Function(i) Not paths.Contains(i.FullPath)).ToList()
                    End Function,
                    Function(shellItem2 As IShellItem2)
                        Return New Folder(shellItem2, Me, _setIsLoadingAction, _cachedIconSize)
                    End Function,
                    Function(shellItem2 As IShellItem2)
                        Return New Item(shellItem2, Me, _cachedIconSize)
                    End Function,
                    Sub(path As String)
                        Dim item As Item = items.FirstOrDefault(Function(i) i.FullPath = path AndAlso Not i.disposedValue)
                        If Not item Is Nothing Then
                            item._shellItem2.Update(IntPtr.Zero)
                            For Each prop In item.GetType().GetProperties()
                                item.NotifyOfPropertyChange(prop.Name)
                            Next
                        End If
                    End Sub,
                    Sub()
                    End Sub, 0, isOnUIThread)
    End Sub

    Protected Sub updateItems(flags As UInt32, condition As Boolean,
                              exists As Func(Of Item, Boolean), add As Action(Of Item), remove As Action(Of Item),
                              getToBeRemoved As Func(Of List(Of String), List(Of Item)),
                              makeNewFolder As Func(Of IShellItem2, Item), makeNewItem As Func(Of IShellItem2, Item),
                              updateProperties As Action(Of String), addLoadingItem As Action, uiHelp As Integer, isOnUIThread As Boolean)
        If _shellItem2 Is Nothing Then Return

        Dim paths As List(Of String) = New List(Of String)

        If Not condition AndAlso (Me.Attributes.HasFlag(SFGAO.ISSLOW) OrElse Me.FullPath.StartsWith("\\")) AndAlso Me.Attributes.HasFlag(SFGAO.HASSUBFOLDER) Then
            addLoadingItem()
        Else
            Dim toAdd As List(Of Item) = New List(Of Item)()
            Dim toUpdate As List(Of String) = New List(Of String)()

            SyncLock _itemsLock
                If Not isWindows7OrLower() Then
                    Dim bindCtx As ComTypes.IBindCtx, bindCtxPtr As IntPtr
                    Functions.CreateBindCtx(0, bindCtxPtr)
                    bindCtx = Marshal.GetTypedObjectForIUnknown(bindCtxPtr, GetType(ComTypes.IBindCtx))

                    If Not bindCtx Is Nothing AndAlso Not IntPtr.Zero.Equals(bindCtxPtr) Then
                        Dim propertyBag As IPropertyBag, propertyBagPtr As IntPtr
                        Try
                            Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBagPtr)
                            propertyBag = Marshal.GetTypedObjectForIUnknown(propertyBagPtr, GetType(IPropertyBag))
                        Finally
                            If Not IntPtr.Zero.Equals(propertyBagPtr) Then
                                Marshal.Release(bindCtxPtr)
                            End If
                        End Try

                        If Not propertyBag Is Nothing Then
                            Dim var As PROPVARIANT
                            Dim enumShellItems As IEnumShellItems
                            Try
                                var.vt = VarEnum.VT_UI4
                                var.union.uintVal = flags
                                propertyBag.Write("SHCONTF", var) '  STR_ENUM_ITEMS_FLAGS 

                                bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag) ' STR_PROPERTYBAG_PARAM 

                                Dim ptr2 As IntPtr
                                Try
                                    _shellItem2.BindToHandler(bindCtxPtr, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, ptr2)
                                    If Not IntPtr.Zero.Equals(ptr2) Then
                                        enumShellItems = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IEnumShellItems))
                                    End If
                                Finally
                                    If Not IntPtr.Zero.Equals(ptr2) Then Marshal.Release(ptr2)
                                End Try

                                If Not enumShellItems Is Nothing Then
                                    Dim shellItemArray(0) As IShellItem, fetched As UInt32 = 1
                                    enumShellItems.Next(1, shellItemArray, fetched)
                                    If fetched = 1 AndAlso Not condition Then
                                        addLoadingItem()
                                    Else
                                        While fetched = 1
                                            Dim attr2 As Integer = SFGAO.FOLDER
                                            shellItemArray(0).GetAttributes(attr2, attr2)
                                            Dim fullPath As String = Item.GetFullPathFromShellItem2(shellItemArray(0))
                                            Dim newItem As Item
                                            If CBool(attr2 And SFGAO.FOLDER) Then
                                                newItem = makeNewFolder(shellItemArray(0))
                                            Else
                                                newItem = makeNewItem(shellItemArray(0))
                                            End If
                                            If Not newItem Is Nothing Then
                                                paths.Add(newItem.FullPath)
                                            End If
                                            If Not newItem Is Nothing Then
                                                If Not exists(newItem) Then
                                                    toAdd.Add(newItem)
                                                Else
                                                    toUpdate.Add(newItem.FullPath)
                                                    newItem.Dispose()
                                                End If
                                            End If
                                            enumShellItems.Next(1, shellItemArray, fetched)
                                        End While
                                    End If
                                End If
                            Catch ex As Exception
                                toAdd.Add(New DummyTreeViewFolder(ex.Message, Nothing))
                            Finally
                                If Not enumShellItems Is Nothing Then
                                    Marshal.ReleaseComObject(enumShellItems)
                                End If
                                Marshal.ReleaseComObject(bindCtx)
                                Marshal.ReleaseComObject(propertyBag)
                                var.Dispose()
                            End Try
                        End If
                    End If
                Else
                    Dim list As IEnumIDList
                    _shellFolder.EnumObjects(Nothing, flags, list)
                    If Not list Is Nothing Then
                        Dim pidl(0) As IntPtr, fetched As Integer
                        While list.Next(1, pidl, fetched) = 0
                            Dim attr2 As Integer = SFGAO.FOLDER
                            _shellFolder.GetAttributesOf(1, pidl, attr2)
                            Dim shellItem2 As IShellItem2 = Item.GetIShellItem2FromPidl(pidl(0), _shellFolder)
                            Dim path As String = Item.GetFullPathFromShellItem2(shellItem2)
                            Dim newItem As Item
                            If CBool(attr2 And SFGAO.FOLDER) Then
                                newItem = makeNewFolder(shellItem2)
                            Else
                                newItem = makeNewItem(shellItem2)
                            End If
                            paths.Add(newItem.FullPath)
                            If Not exists(newItem) Then
                                toAdd.Add(newItem)
                            Else
                                toUpdate.Add(newItem.FullPath)
                                newItem.Dispose()
                            End If
                        End While
                    End If
                End If

                If Not isOnUIThread Then
                    For Each item In toAdd
                        add(item)
                        If uiHelp <> 0 Then Thread.Sleep(uiHelp)
                    Next
                    For Each item In toUpdate
                        updateProperties(item)
                        If uiHelp <> 0 Then Thread.Sleep(uiHelp)
                    Next
                    For Each item In getToBeRemoved(paths)
                        remove(item)
                        item.Dispose()
                        If uiHelp <> 0 Then Thread.Sleep(uiHelp)
                    Next
                Else
                    UIHelper.OnUIThread(
                        Sub()
                            For Each item In toAdd
                                add(item)
                                If uiHelp <> 0 Then Thread.Sleep(uiHelp)
                            Next
                            For Each item In toUpdate
                                updateProperties(item)
                                If uiHelp <> 0 Then Thread.Sleep(uiHelp)
                            Next
                            For Each item In getToBeRemoved(paths)
                                remove(item)
                                item.Dispose()
                                If uiHelp <> 0 Then Thread.Sleep(uiHelp)
                            Next
                        End Sub)
                End If
            End SyncLock
        End If
    End Sub


    Protected Overrides Async Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        MyBase.shell_Notification(sender, e)

        Dim t As Func(Of Task) =
            Async Function() As Task
                If Not _shellItem2 Is Nothing AndAlso Not disposedValue Then
                    Select Case e.Event
                        Case SHCNE.CREATE
                            If Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                                Dim parentShellItem2 As IShellItem2
                                Try
                                    Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                                    If Not item1 Is Nothing Then
                                        item1.GetParent(parentShellItem2)
                                        Dim parentFullPath As String
                                        parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                                        If Me.FullPath.Equals(parentFullPath) Then
                                            SyncLock _itemsLock
                                                UIHelper.OnUIThread(
                                                    Sub()
                                                        If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue) Is Nothing Then
                                                            _items.Add(New Item(item1, Me, _cachedIconSize))
                                                        End If
                                                    End Sub)
                                            End SyncLock
                                        End If
                                    End If
                                Finally
                                    If Not parentShellItem2 Is Nothing Then
                                        Marshal.ReleaseComObject(parentShellItem2)
                                    End If
                                End Try
                            End If
                        Case SHCNE.MKDIR
                            If Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                                Dim parentShellItem2 As IShellItem2
                                Try
                                    Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                                    If Not item1 Is Nothing Then
                                        item1.GetParent(parentShellItem2)
                                        Dim parentFullPath As String
                                        parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                                        If Me.FullPath.Equals(parentFullPath) Then
                                            SyncLock _itemsLock
                                                UIHelper.OnUIThread(
                                                    Sub()
                                                        If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue) Is Nothing Then
                                                            _items.Add(New Folder(item1, Me, _setIsLoadingAction, _cachedIconSize))
                                                        End If
                                                    End Sub)
                                            End SyncLock
                                        End If
                                    End If
                                Finally
                                    If Not parentShellItem2 Is Nothing Then
                                        Marshal.ReleaseComObject(parentShellItem2)
                                    End If
                                End Try
                            End If
                        Case SHCNE.RMDIR, SHCNE.DELETE
                            SyncLock _itemsLock
                                If Not _items Is Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                                    Dim item As Item = _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue)
                                    If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                                        UIHelper.OnUIThread(
                                            Sub()
                                                Shell.RaiseFolderNotificationEvent(Me, New Events.FolderNotificationEventArgs() With {
                                                    .Folder = item,
                                                    .[Event] = e.Event
                                                })
                                                _items.Remove(item)
                                            End Sub)
                                    End If
                                End If
                            End SyncLock
                        Case SHCNE.DRIVEADD
                            If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                                SyncLock _itemsLock
                                    UIHelper.OnUIThread(
                                        Sub()
                                            If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue) Is Nothing Then
                                                Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                                                If Not item1 Is Nothing Then
                                                    _items.Add(New Folder(item1, Me, _setIsLoadingAction, _cachedIconSize))
                                                End If
                                            End If
                                        End Sub)
                                End SyncLock
                            End If
                        Case SHCNE.DRIVEREMOVED
                            If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                                SyncLock _itemsLock
                                    Dim item As Item = _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue)
                                    If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                                        UIHelper.OnUIThread(
                                            Sub()
                                                Shell.RaiseFolderNotificationEvent(Me, New Events.FolderNotificationEventArgs() With {
                                                    .Folder = item,
                                                    .[Event] = e.Event
                                                })
                                                _items.Remove(item)
                                            End Sub)
                                    End If
                                End SyncLock
                            End If
                        Case SHCNE.UPDATEDIR, SHCNE.UPDATEITEM
                            If (Me.FullPath.Equals(e.Item1Path) OrElse Shell.Desktop.FullPath.Equals(e.Item1Path)) AndAlso Not _items Is Nothing Then
                                SyncLock _itemsLock
                                    updateItems(_items, True)
                                End SyncLock
                            End If
                    End Select
                End If
            End Function

        Await Task.Run(t)
    End Sub

    Protected Function isWindows7OrLower() As Boolean
        Dim osVersion As Version = Environment.OSVersion.Version
        ' Windows 7 has version number 6.1
        Return osVersion.Major < 6 OrElse (osVersion.Major = 6 AndAlso osVersion.Minor <= 1)
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If Not _items Is Nothing Then
                For Each item In _items
                    item.Dispose()
                Next
            End If
        End If

        If Not _shellFolder Is Nothing Then
            Marshal.ReleaseComObject(_shellFolder)
            _shellFolder = Nothing
        End If
        If Not _columnManager Is Nothing Then
            Marshal.ReleaseComObject(_columnManager)
            _columnManager = Nothing
        End If

        MyBase.Dispose(disposing)
    End Sub
End Class
