Imports Laila.Shell.Adorners
Imports Laila.Shell.Controls
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls

Namespace Behaviors
    Public Class GridViewShellBehavior
        Inherits GridViewExtBehavior

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(GridViewShellBehavior), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Protected Overrides Sub SetSort(propertyName As String, direction As ListSortDirection)
            If Not Me.Folder Is Nothing Then
                Dim isSwitchingDirection As Boolean
                If Not propertyName.Equals(Me.Folder.ItemsSortPropertyName) Then
                    Me.Folder.ItemsSortPropertyName = propertyName
                Else
                    isSwitchingDirection = True
                End If

                If Not String.IsNullOrWhiteSpace(Me.Folder.ItemsSortPropertyName) Then
                    If direction <> -1 AndAlso Me.Folder.ItemsSortDirection <> direction Then
                        Me.Folder.ItemsSortDirection = direction
                    Else
                        If isSwitchingDirection Then
                            If Me.Folder.ItemsSortDirection = ListSortDirection.Ascending Then
                                Me.Folder.ItemsSortDirection = ListSortDirection.Descending
                            Else
                                Me.Folder.ItemsSortDirection = ListSortDirection.Ascending
                            End If
                        End If
                    End If
                End If
            End If
        End Sub

        Protected Overrides Sub ResetSortDescriptions(view As ICollectionView)

        End Sub

        Protected Overrides Sub SetGrouping(propertyName As String)
            If Not Me.Folder Is Nothing Then
                Me.Folder.ItemsGroupByPropertyName = propertyName
            End If
        End Sub

        Protected Overrides Sub WriteState(viewName As String, state As GridViewStateData)
            ' we're already writing our state in the FolderView
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property
    End Class
End Namespace