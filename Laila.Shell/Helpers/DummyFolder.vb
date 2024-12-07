Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Public Class DummyFolder
    Inherits Folder

    Private _displayName As String

    Public Sub New(displayName As String, logicalParent As Folder, Optional id As String = Nothing)
        MyBase.New(Nothing, logicalParent, True)

        _displayName = displayName
        _fullPath = "dummy" & If(id Is Nothing, Guid.NewGuid().ToString(), id)
        Me.IsLoading = True
    End Sub

    Public Overrides Sub Refresh()
    End Sub

    Public Overrides ReadOnly Property DisplayName As String
        Get
            Return _displayName
        End Get
    End Property

    Public Overrides ReadOnly Property Icon(size As Integer) As ImageSource
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides ReadOnly Property PropertiesByCanonicalName(canonicalName As String) As [Property]
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides ReadOnly Property PropertiesByKey(propertyKey As PROPERTYKEY) As [Property]
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides ReadOnly Property PropertiesByKeyAsText(propertyKey As String) As [Property]
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides Async Function GetItemsAsync() As Task(Of List(Of Item))
        Return New List(Of Item)
    End Function

    Public Overrides ReadOnly Property Items As ObservableCollection(Of Item)
        Get
            Return Nothing
        End Get
    End Property
End Class
