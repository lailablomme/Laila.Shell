Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Media
Imports Laila.Shell.Helpers

Public Class SearchFolder
    Inherits Folder

    Public Property Terms As String

    Private _threadCompletionSource As TaskCompletionSource = New TaskCompletionSource()

    Public Shared Function FromTerms(terms As String, parent As Folder) As SearchFolder
        Return New SearchFolder(getShellItem(terms, parent), parent) With {.View = "Content", .Terms = terms}
    End Function

    Private Shared Function getShellItem(terms As String, parent As Folder) As IShellItem2
        Dim factory As ISearchFolderItemFactory = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_SearchFolderItemFactory))
        Dim arr As IShellItemArray
        Dim pidls As List(Of Pidl)
        If parent.Pidl.Equals(Shell.GetSpecialFolder("This pc").Pidl) Then
            pidls = parent.GetItems().Where(Function(i) TypeOf i Is Folder).Select(Function(i) i.Pidl).ToList()
        Else
            pidls = New List(Of Pidl)() From {parent.Pidl}
        End If
        Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), arr)
        factory.SetScope(arr)
        Dim qpm As IQueryParserManager = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_QueryParserManager))
        Dim qp As IQueryParser
        qpm.CreateLoadedParser("SystemIndex", &H800, GetType(IQueryParser).GUID, qp)
        Dim qs As IQuerySolution
        qp.Parse(terms, Nothing, qs)
        Dim cond As ICondition
        qs.GetQuery(cond, Nothing)
        Dim st As SYSTEMTIME
        Functions.GetLocalTime(st)
        Dim resolvedCond As ICondition
        qs.Resolve(cond, &H40, st, resolvedCond)
        factory.SetCondition(resolvedCond)
        factory.SetDisplayName("Search results for " & parent.DisplayName)
        factory.SetFolderTypeID(Guids.FOLDERTYPEID_GenericSearchResults)
        Dim shellItem As IShellItem2
        factory.GetShellItem(GetType(IShellItem2).GUID, shellItem)
        Return shellItem
    End Function

    Public Sub New(shellItem2 As IShellItem2, parent As Folder)
        MyBase.New(shellItem2, parent, False, True)

        Me.View = "Content"
        Me.ItemsSortPropertyName = "PropertiesByKeyAsText[49691C90-7E17-101A-A91C-08002B2ECDA9:3].Value"
        Me.ItemsSortDirection = ComponentModel.ListSortDirection.Descending
    End Sub

    Public Overrides Async Function GetItemsAsync() As Task(Of List(Of Item))
        Dim tcs As New TaskCompletionSource(Of List(Of Item))

        Dim staThread As Thread = New Thread(New ThreadStart(
            Sub()
                Dim threadCompletionSource As TaskCompletionSource = _threadCompletionSource

                _lock.Wait()
                Try
                    If Not _isEnumerated Then
                        Dim prevEnumerationCancellationTokenSource As CancellationTokenSource _
                            = _enumerationCancellationTokenSource

                        _enumerationCancellationTokenSource = New CancellationTokenSource()
                        enumerateItems(True, _enumerationCancellationTokenSource.Token)

                        ' terminate previous enumeration thread
                        If Not prevEnumerationCancellationTokenSource Is Nothing Then
                            prevEnumerationCancellationTokenSource.Cancel()
                        End If
                    End If
                    tcs.SetResult(_items.ToList())
                Catch ex As Exception
                    tcs.SetException(ex)
                Finally
                    If _lock.CurrentCount = 0 Then
                        _lock.Release()
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
        If Me.IsLoading Then
            ' cancel enumeration
            Me.CancelEnumeration()
        End If

        UIHelper.OnUIThread(
            Sub()
                _items.Clear()
            End Sub)

        ' set new terms
        Me.Terms = terms

        ' get new shellitem
        Dim oldShellItem2 As IShellItem2 = _shellItem2
        _shellItem2 = getShellItem(terms, Me.Parent)
        If Not oldShellItem2 Is Nothing Then
            Marshal.ReleaseComObject(oldShellItem2)
        End If

        ' re-enumerate
        _isEnumerated = False
        Me.GetItemsAsync()
    End Sub

    Friend Overrides Sub CancelEnumeration()
        If Not _enumerationCancellationTokenSource Is Nothing Then
            _enumerationCancellationTokenSource.Cancel()
            If _lock.CurrentCount = 0 Then
                _lock.Release()
            End If

            ' clear collection
            For Each item In _items.ToList()
                item._parent = Nothing
                SyncLock item._shellItemLock
                    Dim si2 As IShellItem2 = item.ShellItem2
                    If Not si2 Is Nothing Then
                        item._shellItem2 = Nothing
                        Marshal.ReleaseComObject(si2)
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
        Return getShellItem(Me.Terms, Me.Parent)
    End Function
End Class
