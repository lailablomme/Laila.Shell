Imports System.Collections.ObjectModel
Imports System.Windows.Media

Public Class DummyTreeViewFolder
    Inherits TreeViewFolder

    Private _displayName As String

    Public Sub New(displayName As String)
        MyBase.New(Nothing, IntPtr.Zero, Nothing)

        _displayName = displayName
        Me.IsLoading = True
    End Sub

    Public Overrides ReadOnly Property DisplayName As String
        Get
            Return _displayName
        End Get
    End Property

    Public Overrides ReadOnly Property Icon16 As ImageSource
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides Property Folders As ObservableCollection(Of TreeViewFolder)
        Get
            Return Nothing
        End Get
        Friend Set(value As ObservableCollection(Of TreeViewFolder))
        End Set
    End Property
End Class
