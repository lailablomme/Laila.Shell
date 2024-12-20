Imports Laila.Shell.Adorners
Imports Laila.Shell.Controls
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input

Namespace Behaviors
    Public Class GridViewShellBehavior
        Inherits GridViewExtBehavior

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(GridViewShellBehavior), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private Const CHECKBOX_WIDTH As Double = 14

        Public Sub New()
            Me.LeftMargin = If(Shell.Settings.DoShowCheckBoxesToSelect, CHECKBOX_WIDTH, 0)
            AddHandler Shell.Settings.PropertyChanged,
                Sub(s As Object, e As PropertyChangedEventArgs)
                    Select Case e.PropertyName
                        Case "DoShowCheckBoxesToSelect"
                            Me.LeftMargin = If(Shell.Settings.DoShowCheckBoxesToSelect, CHECKBOX_WIDTH, 0)
                    End Select
                End Sub
        End Sub

        Protected Overrides Sub SetSort(propertyName As String, direction As ListSortDirection)
            Using Shell.OverrideCursor(Cursors.Wait)
                If Not Me.Folder Is Nothing Then
                    Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Folder.Items)
                    Using view.DeferRefresh()
                        Dim isSwitchingDirection As Boolean
                        If Not propertyName.Equals(Me.Folder.ItemsSortPropertyName) Then
                            Me.Folder.ItemsSortPropertyName = propertyName
                        Else
                            isSwitchingDirection = True
                        End If

                        If Not String.IsNullOrWhiteSpace(Me.Folder.ItemsSortPropertyName) Then
                            If direction <> -1 AndAlso Me.Folder.ItemsSortDirection <> direction Then
                                Me.Folder.ItemsSortDirection = direction
                            ElseIf direction = -1 AndAlso isSwitchingDirection Then
                                If Me.Folder.ItemsSortDirection = ListSortDirection.Ascending Then
                                    Me.Folder.ItemsSortDirection = ListSortDirection.Descending
                                Else
                                    Me.Folder.ItemsSortDirection = ListSortDirection.Ascending
                                End If
                            End If
                        End If
                    End Using
                End If
            End Using
        End Sub

        Protected Overrides Sub ResetSortDescriptions(view As ICollectionView)

        End Sub

        Protected Overrides Function GetCurrentSortDirection() As ListSortDirection
            If Not Me.Folder Is Nothing Then
                Return Me.Folder.ItemsSortDirection
            Else
                Return ListSortDirection.Ascending
            End If
        End Function

        Protected Overrides Function GetCurrentSortPropertyName() As String
            If Not Me.Folder Is Nothing Then
                Return Me.Folder.ItemsSortPropertyName
            Else
                Return Nothing
            End If
        End Function

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