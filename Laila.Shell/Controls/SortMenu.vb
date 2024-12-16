Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class SortMenu
        Inherits ContextMenu

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(SortMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private _primaryProperties As List(Of String) = New List(Of String)()
        Private _additionalProperties As List(Of String) = New List(Of String)()
        Private _descriptions As Dictionary(Of String, String) = New Dictionary(Of String, String)()
        Private _didGetProperties As Boolean
        Private _isCheckingInternally As Boolean

        Protected Overrides Sub OnOpened(e As RoutedEventArgs)
            Me.Items.Clear()

            Me.AddSortItems(Me.Items)

            Me.Items.Add(New Separator())

            Dim groupByMenuItem As MenuItem = New MenuItem() With {
                .Header = "Group by",
                .Tag = "GroupBy"
            }
            Me.Items.Add(groupByMenuItem)

            Me.AddGroupByItems(groupByMenuItem.Items)

            MyBase.OnOpened(e)
        End Sub

        Public Sub AddSortItems(menu As ItemCollection)
            getProperties()

            Dim sortPropertyName As String = Me.Folder.ItemsSortPropertyName
            If Not String.IsNullOrWhiteSpace(sortPropertyName) Then
                If sortPropertyName.IndexOf("[") >= 0 Then
                    sortPropertyName = sortPropertyName.Substring(sortPropertyName.IndexOf("[") + 1)
                    sortPropertyName = sortPropertyName.Substring(0, sortPropertyName.IndexOf("]"))
                End If
            End If
            Dim moreMenuItem As MenuItem
            Dim PKEY_System_ItemNameDisplay As String = Me.Folder.PropertiesByCanonicalName("System.ItemNameDisplay").Key.ToString()
            Dim sortMenuItemCheckedAction As RoutedEventHandler = New RoutedEventHandler(
                Sub(s2 As Object, e2 As RoutedEventArgs)
                    If Not _isCheckingInternally Then
                        Using Shell.OverrideCursor(Cursors.Wait)
                            If CType(s2, MenuItem).Tag.ToString().Substring(5).IndexOf(":") >= 0 Then
                                Me.Folder.ItemsSortPropertyName =
                                    String.Format("PropertiesByKeyAsText[{0}].Value",
                                                    CType(s2, MenuItem).Tag.ToString().Substring(5))
                            Else
                                Me.Folder.ItemsSortPropertyName = CType(s2, MenuItem).Tag.ToString().Substring(5)
                            End If
                        End Using
                    End If
                End Sub)
            Dim menuItemUncheckedAction As RoutedEventHandler = New RoutedEventHandler(
                Sub(s2 As Object, e2 As RoutedEventArgs)
                    If Not _isCheckingInternally Then
                        _isCheckingInternally = True
                        CType(s2, MenuItem).IsChecked = True
                        _isCheckingInternally = False
                    End If
                End Sub)
            For Each propName In _primaryProperties
                Dim menuItem As MenuItem = New MenuItem() With {
                    .Header = _descriptions(propName),
                    .Tag = "Sort:" & If(propName = PKEY_System_ItemNameDisplay, "ItemNameDisplaySortValue", propName),
                    .IsCheckable = True
                }
                menuItem.IsChecked = menuItem.Tag.ToString().ToUpper() = "SORT:" & sortPropertyName?.ToUpper()
                AddHandler menuItem.Checked, sortMenuItemCheckedAction
                AddHandler menuItem.Unchecked, menuItemUncheckedAction
                menu.Add(menuItem)
            Next
            If _additionalProperties.Count > 0 Then
                moreMenuItem = New MenuItem() With {.Header = "More", .Tag = "SortMore"}
                menu.Add(moreMenuItem)
                For Each propName In _additionalProperties
                    Dim menuItem As MenuItem = New MenuItem() With {
                        .Header = _descriptions(propName),
                        .Tag = "Sort:" & If(propName = PKEY_System_ItemNameDisplay, "ItemNameDisplaySortValue", propName),
                        .IsCheckable = True
                    }
                    menuItem.IsChecked = menuItem.Tag.ToString().ToUpper() = "SORT:" & sortPropertyName?.ToUpper()
                    AddHandler menuItem.Checked, sortMenuItemCheckedAction
                    AddHandler menuItem.Unchecked, menuItemUncheckedAction
                    moreMenuItem.Items.Add(menuItem)
                Next
            End If
            menu.Add(New Separator())
            Dim sortAscendingMenuItem As MenuItem = New MenuItem() With {
                .Header = "Ascending",
                .IsCheckable = True,
                .Tag = "SortAscending",
                .IsChecked = Me.Folder.ItemsSortDirection = ListSortDirection.Ascending
            }
            menu.Add(sortAscendingMenuItem)
            AddHandler sortAscendingMenuItem.Checked,
                Sub(s2 As Object, e2 As EventArgs)
                    If Not _isCheckingInternally Then
                        Using Shell.OverrideCursor(Cursors.Wait)
                            Me.Folder.ItemsSortDirection = ListSortDirection.Ascending
                        End Using
                    End If
                End Sub
            AddHandler sortAscendingMenuItem.Unchecked, menuItemUncheckedAction
            Dim sortDescendingMenuItem As MenuItem = New MenuItem() With {
                .Header = "Descending",
                .IsCheckable = True,
                .Tag = "SortDescending",
                .IsChecked = Me.Folder.ItemsSortDirection = ListSortDirection.Descending
            }
            menu.Add(sortDescendingMenuItem)
            AddHandler sortDescendingMenuItem.Checked,
                Sub(s2 As Object, e2 As EventArgs)
                    If Not _isCheckingInternally Then
                        Using Shell.OverrideCursor(Cursors.Wait)
                            Me.Folder.ItemsSortDirection = ListSortDirection.Descending
                        End Using
                    End If
                End Sub
            AddHandler sortDescendingMenuItem.Unchecked, menuItemUncheckedAction
        End Sub

        Public Sub AddGroupByItems(menu As ItemCollection)
            Dim groupByPropertyName As String = Me.Folder.ItemsGroupByPropertyName
            If Not String.IsNullOrWhiteSpace(groupByPropertyName) Then
                If groupByPropertyName.IndexOf("[") >= 0 Then
                    groupByPropertyName = groupByPropertyName.Substring(groupByPropertyName.IndexOf("[") + 1)
                    groupByPropertyName = groupByPropertyName.Substring(0, groupByPropertyName.IndexOf("]"))
                End If
            End If
            Dim groupByMenuItemCheckedAction As RoutedEventHandler = New RoutedEventHandler(
                Sub(s2 As Object, e2 As RoutedEventArgs)
                    If Not _isCheckingInternally Then
                        Using Shell.OverrideCursor(Cursors.Wait)
                            Me.Folder.ItemsGroupByPropertyName =
                                String.Format("PropertiesByKeyAsText[{0}].Text",
                                                CType(s2, MenuItem).Tag.ToString().Substring(6))
                        End Using
                    End If
                End Sub)
            Dim menuItemUncheckedAction As RoutedEventHandler = New RoutedEventHandler(
                Sub(s2 As Object, e2 As RoutedEventArgs)
                    If Not _isCheckingInternally Then
                        _isCheckingInternally = True
                        CType(s2, MenuItem).IsChecked = True
                        _isCheckingInternally = False
                    End If
                End Sub)
            For Each propName In _primaryProperties
                Dim menuItem As MenuItem = New MenuItem() With {
                    .Header = _descriptions(propName),
                    .Tag = "Group:" & propName,
                    .IsCheckable = True
                }
                menuItem.IsChecked = menuItem.Tag.ToString().ToUpper() = "GROUP:" & groupByPropertyName?.ToUpper()
                AddHandler menuItem.Checked, groupByMenuItemCheckedAction
                AddHandler menuItem.Unchecked, menuItemUncheckedAction
                menu.Add(menuItem)
            Next
            For Each propName In _additionalProperties
                Dim menuItem As MenuItem = New MenuItem() With {
                    .Header = _descriptions(propName),
                    .Tag = "Group:" & propName,
                    .IsCheckable = True
                }
                menuItem.IsChecked = menuItem.Tag.ToString().ToUpper() = "GROUP:" & groupByPropertyName?.ToUpper()
                AddHandler menuItem.Checked, groupByMenuItemCheckedAction
                AddHandler menuItem.Unchecked, menuItemUncheckedAction
                menu.Add(menuItem)
            Next
            If Not String.IsNullOrWhiteSpace(groupByPropertyName) Then
                Dim groupByNoneMenuItem As MenuItem = New MenuItem() With {
                    .Header = "(None)",
                    .Tag = "Group:",
                    .IsCheckable = True
                }
                AddHandler groupByNoneMenuItem.Checked,
                    Sub(s2 As Object, e2 As EventArgs)
                        Me.Folder.ItemsGroupByPropertyName = Nothing
                    End Sub
                AddHandler groupByNoneMenuItem.Unchecked, menuItemUncheckedAction
                menu.Add(groupByNoneMenuItem)
            End If
        End Sub

        Private Sub getProperties()
            If _didGetProperties Then Return

            For Each column In Folder.Columns.Take(5)
                _primaryProperties.Add(column.PROPERTYKEY.ToString())
                _descriptions.Add(column.PROPERTYKEY.ToString(), column.DisplayName)
            Next

            For Each column In Folder.Columns.Skip(5).Take(5)
                _additionalProperties.Add(column.PROPERTYKEY.ToString())
                _descriptions.Add(column.PROPERTYKEY.ToString(), column.DisplayName)
            Next

            Dim getProperties As Action(Of Item, PROPERTYKEY) =
                Sub(item As Item, pKey As PROPERTYKEY)
                    Dim text As String = item.PropertiesByKey(pKey).Text
                    If String.IsNullOrWhiteSpace(text) Then text = "prop:System.ItemTypeText;System.Size"
                    For Each propName In text.Substring(5).Split(";")
                        propName = propName.TrimStart("*")
                        Dim prop As [Property] = item.PropertiesByCanonicalName(propName)
                        If Not prop Is Nothing _
                                AndAlso Not _primaryProperties.Contains(prop.Key.ToString()) _
                                AndAlso Not _additionalProperties.Contains(prop.Key.ToString()) Then
                            _additionalProperties.Add(prop.Key.ToString())
                            _descriptions.Add(prop.Key.ToString(), prop.DisplayName)
                        End If
                    Next
                End Sub

            Dim PKEY_System_InfoTipText As New PROPERTYKEY() With {
                .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                .pid = 4
            }
            Dim PKEY_System_PropList_TileInfo As New PROPERTYKEY() With {
                .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                .pid = 3
            }

            'If TypeOf Me.Folder Is SearchFolder Then
            '    _additionalProperties.Add("49691C90-7E17-101A-A91C-08002B2ECDA9:3")
            '    _descriptions.Add("49691C90-7E17-101A-A91C-08002B2ECDA9:3", "Relevance")
            'End If

            Dim uniqueTypes As List(Of String) = New List(Of String)()
            Dim items As List(Of Item) = Me.Folder.Items.ToList()

            For Each item In items
                Dim ext As String = IO.Path.GetExtension(item.FullPath)
                If Not String.IsNullOrWhiteSpace(ext) AndAlso Not uniqueTypes.Contains(ext) Then
                    uniqueTypes.Add(ext)
                    getProperties(item, PKEY_System_InfoTipText)
                    getProperties(item, PKEY_System_PropList_TileInfo)
                End If
            Next

            _didGetProperties = True
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