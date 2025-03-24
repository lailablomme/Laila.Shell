﻿Imports System.Collections.ObjectModel
Imports System.Windows.Media
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties

Namespace Helpers
    Public Class DummyFolder
        Inherits Folder

        Public Sub New(displayName As String, logicalParent As Folder, Optional id As String = Nothing)
            MyBase.New(Nothing, logicalParent, True, False, -1)

            _displayName = displayName
            _fullPath = "dummy" & If(id Is Nothing, Guid.NewGuid().ToString(), id)
            Me.IsLoading = True
        End Sub

        Public Overrides Sub Refresh(Optional newShellItem As IShellItem2 = Nothing,
                                     Optional newPidl As Pidl = Nothing,
                                     Optional newFullPath As String = Nothing,
                                     Optional doRefreshImage As Boolean = True)
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

        Public Overrides Function GetItemsAsync(Optional doRefreshAllExistingItems As Boolean = True,
                                                      Optional doRecursive As Boolean = False) As Task(Of List(Of Item))
            Return Task.FromResult(New List(Of Item))
        End Function

        Public Overrides ReadOnly Property Items As ObservableCollection(Of Item)
            Get
                Return Nothing
            End Get
        End Property
    End Class
End Namespace