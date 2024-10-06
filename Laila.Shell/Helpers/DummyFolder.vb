﻿Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Public Class DummyFolder
    Inherits Folder

    Private _displayName As String

    Public Sub New(displayName As String, Optional id As String = Nothing)
        MyBase.New(Nothing, Nothing)

        _displayName = displayName
        _fullPath = "dummy" & If(id Is Nothing, Guid.NewGuid().ToString(), id)
        Me.IsLoading = True
    End Sub

    Public Overrides Sub ClearCache()
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
            If Not _icon.ContainsKey(size) Then
                Return Nothing
            End If
            Return _icon(size)
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

    Public Overrides Async Function GetItems() As Task(Of List(Of Item))
        Return New List(Of Item)
    End Function

    Public Overrides ReadOnly Property ItemsThreaded As ObservableCollection(Of Item)
        Get
            Return Nothing
        End Get
    End Property
End Class
