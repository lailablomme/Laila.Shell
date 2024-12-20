Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Media

Public Class SearchFolder
    Inherits Folder

    Public Property Terms As String
    Public Property Parent As Folder

    Public Shared Function FromTerms(terms As String, parent As Folder) As SearchFolder
        Return New SearchFolder(getShellItem(terms, parent)) With {.View = "Content", .Terms = terms, .Parent = parent}
    End Function

    Private Shared Function getShellItem(terms As String, parent As Folder) As IShellItem2
        Dim factory As ISearchFolderItemFactory = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_SearchFolderItemFactory))
        Dim arr As IShellItemArray
        Dim pidls As List(Of Pidl)
        If parent.Pidl.Equals(Shell.GetSpecialFolder("This computer").Pidl) Then
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
        factory.SetDisplayName("Search in " & parent.DisplayName)
        Dim shellItem As IShellItem2
        factory.GetShellItem(GetType(IShellItem2).GUID, shellItem)
        Return shellItem
    End Function

    Public Sub New(shellItem2 As IShellItem2)
        MyBase.New(shellItem2, Nothing, False, True)

        Me.View = "Content"
        Me.ItemsSortPropertyName = "PropertiesByKeyAsText[49691C90-7E17-101A-A91C-08002B2ECDA9:3].Value"
        Me.ItemsSortDirection = ComponentModel.ListSortDirection.Descending
    End Sub

    Public Sub Update(terms As String)
        If Not _enumerationCancellationTokenSource Is Nothing Then
            _enumerationCancellationTokenSource.Cancel()
        End If
        Me.Terms = terms
        Dim oldShellItem2 As IShellItem2 = _shellItem2
        _shellItem2 = getShellItem(terms, Me.Parent)
        If Not oldShellItem2 Is Nothing Then
            '_shellItemHistory.Add(New Tuple(Of IShellItem2, Date)(_shellItem2, DateTime.Now))
            Marshal.ReleaseComObject(oldShellItem2)
        End If
        For Each item In _items.ToList()
            item._parent = Nothing
        Next
        _items.Clear()
        _isEnumerated = False
        Me.GetItemsAsync()
    End Sub

    Public Sub CancelUpdate()
        If Not _enumerationCancellationTokenSource Is Nothing Then
            _enumerationCancellationTokenSource.Cancel()
            Marshal.ReleaseComObject(_shellItem2)
            _shellItem2 = Nothing
        End If
    End Sub

    Protected Overrides Sub MakeNewShellItem()
        _shellItem2 = getShellItem(Me.Terms, Me.Parent)
    End Sub
End Class
