Imports System.Collections.ObjectModel
Imports System.Windows.Media

Public Class DummyFolder
    Inherits Folder

    Private _displayName As String

    Public Sub New(displayName As String, list As IList)
        MyBase.New(Nothing, Nothing)

        _displayName = displayName
        Me.IsLoading = True
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

    Protected Overrides Function getOverlay(isLarge As Boolean) As ImageSource
        Return Nothing
    End Function

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
End Class
