Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Media
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Folders
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.System

Public Class SearchFolder
    Inherits Folder

    Public Property Terms As String

    Private _threadCompletionSource As TaskCompletionSource = New TaskCompletionSource()

    Public Shared Function FromTerms(terms As String, parent As Folder) As SearchFolder
        Dim threadId As Integer = Shell.GlobalThreadPool.GetNextFreeThreadId()
        Dim folder As SearchFolder = Shell.GlobalThreadPool.Run(
            Function() As SearchFolder
                Return New SearchFolder(getShellItem(terms, parent), parent, threadId) With {.ActiveView = New Guid("9d96a6be-4061-4c89-9cb7-d97c6b8dfe41"), .Terms = terms}
            End Function,, threadId)

        folder.ItemsSortPropertyName = "PropertiesByKeyAsText[49691C90-7E17-101A-A91C-08002B2ECDA9:3].Value"
        folder.ItemsSortDirection = ComponentModel.ListSortDirection.Descending

        Return folder
    End Function

    Private Shared Function getShellItem(terms As String, parent As Folder) As IShellItem2
        Dim factory As ISearchFolderItemFactory = Nothing, array As IShellItemArray = Nothing
        Dim queryParserManager As IQueryParserManager = Nothing, queryParser As IQueryParser = Nothing
        Dim condition As ICondition = Nothing, resolvedCondition As ICondition = Nothing
        Dim querySolution As IQuerySolution = Nothing
        Try
            ' make factory
            factory = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_SearchFolderItemFactory))

            ' set scope
            Dim pidls As List(Of Pidl)
            If parent.Pidl.Equals(Shell.GetSpecialFolder(SpecialFolders.ThisPc).Pidl) Then
                pidls = parent.GetItems().Where(Function(i) TypeOf i Is Folder).Select(Function(i) i.Pidl).ToList()
            Else
                pidls = New List(Of Pidl)() From {parent.Pidl}
            End If
            Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), array)
            factory.SetScope(array)

            ' make query
            queryParserManager = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_QueryParserManager))
            queryParserManager.CreateLoadedParser("SystemIndex", &H800, GetType(IQueryParser).GUID, queryParser)
            queryParser.Parse(terms, Nothing, querySolution)
            querySolution.GetQuery(condition, Nothing)

            ' resolve query and set condition
            Dim systemTime As SYSTEMTIME
            Functions.GetLocalTime(systemTime)
            querySolution.Resolve(condition, &H40, systemTime, resolvedCondition)
            factory.SetCondition(resolvedCondition)

            ' set folder properties
            factory.SetDisplayName(String.Format(My.Resources.SearchFolder_SearchResultsFor, parent.DisplayName))
            factory.SetFolderTypeID(Guids.FOLDERTYPEID_GenericSearchResults)

            ' get and return shellitem
            Dim shellItem As IShellItem2 = Nothing
            factory.GetShellItem(GetType(IShellItem2).GUID, shellItem)
            Return shellItem
        Finally
            If Not factory Is Nothing Then
                Marshal.ReleaseComObject(factory)
                factory = Nothing
            End If
            If Not array Is Nothing Then
                Marshal.ReleaseComObject(array)
                array = Nothing
            End If
            If Not queryParserManager Is Nothing Then
                Marshal.ReleaseComObject(queryParserManager)
                queryParserManager = Nothing
            End If
            If Not queryParser Is Nothing Then
                Marshal.ReleaseComObject(queryParser)
                queryParser = Nothing
            End If
            If Not querySolution Is Nothing Then
                Marshal.ReleaseComObject(querySolution)
                querySolution = Nothing
            End If
            If Not condition Is Nothing Then
                Marshal.ReleaseComObject(condition)
                condition = Nothing
            End If
            If Not resolvedCondition Is Nothing Then
                Marshal.ReleaseComObject(resolvedCondition)
                resolvedCondition = Nothing
            End If
        End Try
        Return Nothing
    End Function

    Public Sub New(shellItem2 As IShellItem2, parent As Folder, threadId As Integer)
        MyBase.New(shellItem2, parent, False, True, threadId)
    End Sub

    Public Overrides Async Function GetItemsAsync(Optional doRefreshAllExistingItems As Boolean = True,
                                                  Optional doRecursive As Boolean = False,
                                                  Optional isForTree As Boolean = False) As Task(Of List(Of Item))
        Dim tcs As New TaskCompletionSource(Of List(Of Item))

        Dim staThread As Thread = New Thread(New ThreadStart(
            Sub()
                Dim threadCompletionSource As TaskCompletionSource = _threadCompletionSource

                _enumerationLock.Wait()
                Try
                    If Not _isEnumerated OrElse Not _isEnumeratedForTree Then
                        Me.IsLoading = True

                        Dim prevEnumerationCancellationTokenSource As CancellationTokenSource _
                            = _enumerationCancellationTokenSource

                        _enumerationCancellationTokenSource = New CancellationTokenSource()
                        enumerateItems(True, _enumerationCancellationTokenSource.Token, Nothing, doRefreshAllExistingItems)

                        Me.IsLoading = False

                        ' terminate previous enumeration thread
                        If Not prevEnumerationCancellationTokenSource Is Nothing Then
                            prevEnumerationCancellationTokenSource.Cancel()
                        End If
                    End If
                    tcs.SetResult(_items.ToList())
                Catch ex As Exception
                    tcs.SetException(ex)
                Finally
                    If _enumerationLock.CurrentCount = 0 Then
                        _enumerationLock.Release()
                    End If
                End Try

                If Not threadCompletionSource Is Nothing Then
                    threadCompletionSource.Task.Wait()
                End If
            End Sub))
        staThread.IsBackground = True
        staThread.SetApartmentState(ApartmentState.STA)
        staThread.Start()

        Await tcs.Task.WaitAsync(Shell.ShuttingDownToken)
        If Not Shell.ShuttingDownToken.IsCancellationRequested Then
            Return tcs.Task.Result
        End If
        Return Nothing
    End Function

    Public Sub Update(terms As String)
        If Me.Terms.Equals(terms) Then Return

        If Me.IsLoading Then
            ' cancel enumeration
            Me.CancelEnumeration()
        End If

        _items.Clear()

        ' set new terms
        Me.Terms = terms

        ' get new shellitem
        Shell.GlobalThreadPool.Add(
            Sub()
                Dim oldShellItem2 As IShellItem2 = _shellItem2
                _shellItem2 = getShellItem(terms, Me.LogicalParent)
                If Not oldShellItem2 Is Nothing Then
                    Marshal.ReleaseComObject(oldShellItem2)
                    oldShellItem2 = Nothing
                End If

                ' re-enumerate
                _isEnumerated = False
                _isEnumeratedForTree = False
                Dim __ = Me.GetItemsAsync()
            End Sub)
    End Sub

    Public Overrides Sub CancelEnumeration()
        If Not _enumerationCancellationTokenSource Is Nothing Then
            _enumerationCancellationTokenSource.Cancel()
            If _enumerationLock.CurrentCount = 0 Then
                _enumerationLock.Release()
            End If

            ' clear collection
            For Each item In _items.ToList()
                item._parent = Nothing
                SyncLock item._shellItemLockSearchFolder
                    Dim si2 As IShellItem2 = item.ShellItem2
                    If Not si2 Is Nothing Then
                        item._shellItem2 = Nothing
                        Marshal.ReleaseComObject(si2)
                        si2 = Nothing
                    End If
                End SyncLock
            Next

            ' terminate thread
            If Not _threadCompletionSource.Task.IsCompleted Then
                _threadCompletionSource.SetResult()
                _threadCompletionSource = New TaskCompletionSource()
            End If
        End If
    End Sub

    Protected Overrides Function GetNewShellItem() As IShellItem2
        Return getShellItem(Me.Terms, Me.LogicalParent)
    End Function
End Class
