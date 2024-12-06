Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input

Namespace Controls
    Public Class SearchBox
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(SearchBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private PART_TextBox As TextBox

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(SearchBox), New FrameworkPropertyMetadata(GetType(SearchBox)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_TextBox = Me.Template.FindName("PART_TextBox", Me)

            AddHandler Me.PART_TextBox.PreviewKeyDown,
                Sub(s As Object, e As KeyEventArgs)
                    If e.Key = Key.Enter AndAlso Not String.IsNullOrWhiteSpace(Me.PART_TextBox.Text) Then
                        Dim factory As ISearchFolderItemFactory = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_SearchFolderItemFactory))
                        Dim arr As IShellItemArray
                        Dim pidls As List(Of Pidl) = Me.Folder.Items.Where(Function(i) TypeOf i Is Folder).Select(Function(i) i.Pidl).ToList()
                        If Not Me.Folder.Pidl.Equals(Shell.GetSpecialFolder("This computer").Pidl) Then pidls.Insert(0, Me.Folder.Pidl)
                        Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), arr)
                        'Functions.SHCreateShellItemArrayFromShellItem(Me.Folder.ShellItem2, GetType(IShellItemArray).GUID, arr)
                        factory.SetScope(arr)
                        Dim qpm As IQueryParserManager = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_QueryParserManager))
                        Dim qp As IQueryParser
                        qpm.CreateLoadedParser("SystemIndex", &H800, GetType(IQueryParser).GUID, qp)
                        Dim qs As IQuerySolution
                        qp.Parse(Me.PART_TextBox.Text, Nothing, qs)
                        Dim cond As ICondition
                        qs.GetQuery(cond, Nothing)
                        Dim st As SYSTEMTIME
                        Functions.GetLocalTime(st)
                        Dim resolvedCond As ICondition
                        qs.Resolve(cond, &H40, st, resolvedCond)
                        factory.SetCondition(resolvedCond)
                        factory.SetDisplayName("Search in " & Me.Folder.DisplayName)
                        Dim shellItem As IShellItem2
                        factory.GetShellItem(GetType(IShellItem2).GUID, shellItem)
                        Dim f As Folder = New Folder(shellItem, Nothing)
                        f.View = "Content"
                        Me.Folder = f
                    End If
                End Sub
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(value As Folder)
                SetValue(FolderProperty, value)
            End Set
        End Property
    End Class
End Namespace