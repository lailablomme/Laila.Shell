Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Helpers
Imports System.ComponentModel
Imports System.Windows.Data
Imports System.Windows.Media

Namespace Controls
    Public Class DetailsView
        Inherits BaseFolderView

        Private PART_Ext As Behaviors.GridViewShellBehavior
        Private _isLoaded As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(DetailsView), New FrameworkPropertyMetadata(GetType(DetailsView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            AddHandler PART_ListBox.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        ' notify of sort/group by changes
                        Me.PART_Ext = Microsoft.Xaml.Behaviors.Interaction.GetBehaviors(Me.PART_ListBox).FirstOrDefault(Function(b) TypeOf b Is Behaviors.GridViewExtBehavior)
                    End If
                End Sub
        End Sub

        Public Function buildColumnsIn(folder As Folder) As Behaviors.GridViewExtBehavior.ColumnsInData
            Dim d As Behaviors.GridViewExtBehavior.ColumnsInData = New Behaviors.GridViewExtBehavior.ColumnsInData()
            d.ViewName = folder.FullPath
            d.Items = New List(Of GridViewColumn)()

            For Each column In folder.Columns.Where(Function(c) Not String.IsNullOrWhiteSpace(c.DisplayName))
                Dim [property] As [Property] = [Property].FromCanonicalName(column.CanonicalName)

                Dim gvc As GridViewColumn = New GridViewColumn()
                gvc.Header = column.DisplayName
                gvc.CellTemplate = getCellTemplate(column, [property])
                gvc.SetValue(Behaviors.GridViewExtBehavior.IsVisibleProperty, column.IsVisible)
                gvc.SetValue(Behaviors.GridViewExtBehavior.PropertyNameProperty, String.Format("PropertiesByKeyAsText[{0}].Value", column.PROPERTYKEY.ToString()))
                If column.CanonicalName = "System.ItemNameDisplay" Then
                    gvc.SetValue(Behaviors.GridViewExtBehavior.SortPropertyNameProperty, "ItemNameDisplaySortValue")
                    gvc.SetValue(Behaviors.GridViewExtBehavior.CanHideProperty, False)
                End If
                'If [property].HasIcon Then
                '    gvc.SetValue(Behaviors.GridViewExtBehavior.ExtraAutoSizeMarginProperty, Convert.ToDouble(15))
                'End If
                gvc.SetValue(Behaviors.GridViewExtBehavior.GroupByPropertyNameProperty, String.Format("PropertiesByKeyAsText[{0}].Text", column.PROPERTYKEY.ToString()))

                Dim isVisibleDescriptor As DependencyPropertyDescriptor =
                    DependencyPropertyDescriptor.FromProperty(Behaviors.GridViewExtBehavior.IsVisibleProperty, gvc.GetType())
                isVisibleDescriptor.AddValueChanged(gvc,
                    Sub(s2 As Object, e2 As EventArgs)
                        column.IsVisible = gvc.GetValue(Behaviors.GridViewExtBehavior.IsVisibleProperty)
                    End Sub)

                d.Items.Add(gvc)
            Next

            Return d
        End Function

        Private Function getCellTemplate(column As Column, [property] As [Property]) As DataTemplate
            Dim template As DataTemplate = New DataTemplate()

            Dim gridFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(Grid))
            Dim columnDefinition0 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition0.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition0)
            Dim columnDefinition1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition1.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition1)
            Dim columnDefinition2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition2.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition2)
            Dim columnDefinition3 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition3.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Star))
            gridFactory.AppendChild(columnDefinition3)

            Dim checkBoxFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(CheckBox))
            checkBoxFactory.SetValue(CheckBox.VerticalAlignmentProperty, VerticalAlignment.Center)
            checkBoxFactory.SetValue(CheckBox.MarginProperty, New Thickness(-2, 0, 4, 0))
            checkBoxFactory.SetValue(CheckBox.IsCheckedProperty, New Binding() With {
                    .Path = New PropertyPath("IsSelected"),
                    .RelativeSource = New RelativeSource() With {
                        .Mode = RelativeSourceMode.FindAncestor,
                        .AncestorType = GetType(ListViewItem)
                    }
                })
            Dim checkBoxTrigger1 As MultiDataTrigger = New MultiDataTrigger()
            checkBoxTrigger1.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath("HasCheckBoxesForItems"),
                    .Source = Me
                },
                .Value = True
            })
            checkBoxTrigger1.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath("IsMouseOver"),
                    .RelativeSource = New RelativeSource() With {
                        .Mode = RelativeSourceMode.FindAncestor,
                        .AncestorType = GetType(ListViewItem)
                    }
                },
                .Value = True
            })
            checkBoxTrigger1.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath(String.Format("ColumnIndexFor[PropertiesByKeyAsText[{0}].Value]", column.PROPERTYKEY.ToString())),
                    .ElementName = "ext",
                    .Mode = BindingMode.OneWay
                },
                .Value = 0
            })
            checkBoxTrigger1.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Visible))
            Dim checkBoxTrigger2 As MultiDataTrigger = New MultiDataTrigger()
            checkBoxTrigger2.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath("HasCheckBoxesForItems"),
                    .Source = Me
                },
                .Value = True
            })
            checkBoxTrigger2.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath("IsSelected"),
                    .RelativeSource = New RelativeSource() With {
                        .Mode = RelativeSourceMode.FindAncestor,
                        .AncestorType = GetType(ListViewItem)
                    }
                },
                .Value = True
            })
            checkBoxTrigger2.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath(String.Format("ColumnIndexFor[PropertiesByKeyAsText[{0}].Value]", column.PROPERTYKEY.ToString())),
                    .ElementName = "ext",
                    .Mode = BindingMode.OneWay
                },
                .Value = 0
            })
            checkBoxTrigger2.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Visible))
            Dim checkBoxTrigger3 As MultiDataTrigger = New MultiDataTrigger()
            checkBoxTrigger3.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath("HasCheckBoxesForItems"),
                    .Source = Me
                },
                .Value = True
            })
            checkBoxTrigger3.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath("IsMouseOver"),
                    .RelativeSource = New RelativeSource() With {
                        .Mode = RelativeSourceMode.FindAncestor,
                        .AncestorType = GetType(ListViewItem)
                    }
                },
                .Value = False
            })
            checkBoxTrigger3.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath("IsSelected"),
                    .RelativeSource = New RelativeSource() With {
                        .Mode = RelativeSourceMode.FindAncestor,
                        .AncestorType = GetType(ListViewItem)
                    }
                },
                .Value = False
            })
            checkBoxTrigger3.Conditions.Add(New Condition() With {
                .Binding = New Binding() With {
                    .Path = New PropertyPath(String.Format("ColumnIndexFor[PropertiesByKeyAsText[{0}].Value]", column.PROPERTYKEY.ToString())),
                    .ElementName = "ext",
                    .Mode = BindingMode.OneWay
                },
                .Value = 0
            })
            checkBoxTrigger3.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Hidden))
            Dim checkBoxStyle As Style = New Style(GetType(CheckBox))
            checkBoxStyle.Setters.Add(New Setter(CheckBox.VisibilityProperty, Visibility.Collapsed))
            checkBoxStyle.Triggers.Add(checkBoxTrigger1)
            checkBoxStyle.Triggers.Add(checkBoxTrigger2)
            checkBoxStyle.Triggers.Add(checkBoxTrigger3)
            checkBoxFactory.SetValue(CheckBox.StyleProperty, checkBoxStyle)
            gridFactory.AppendChild(checkBoxFactory)

            Dim getIconFactory As Func(Of String, FrameworkElementFactory) =
                Function(bindTo As String) As FrameworkElementFactory
                    Dim imageFactory1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                    imageFactory1.SetValue(Grid.ColumnProperty, 1)
                    imageFactory1.SetValue(Image.MarginProperty, New Thickness(0, 0, 4, 0))
                    imageFactory1.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                    imageFactory1.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                    imageFactory1.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                    imageFactory1.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                    imageFactory1.SetValue(Image.SourceProperty, New Binding() With {
                        .Path = New PropertyPath(bindTo),
                        .Mode = BindingMode.OneWay,
                        .IsAsync = True
                    })
                    Dim style As Style = New Style(GetType(Image))
                    style.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Collapsed))
                    style.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(1)))
                    Dim dataTrigger1 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding(String.Format("ColumnIndexFor[PropertiesByKeyAsText[{0}].Value]", column.PROPERTYKEY.ToString())) With
                                   {
                                       .ElementName = "ext",
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = 0
                    }
                    dataTrigger1.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Visible))
                    style.Triggers.Add(dataTrigger1)
                    Dim dataTrigger2 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsHidden") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
                    dataTrigger2.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                    style.Triggers.Add(dataTrigger2)
                    Dim dataTrigger3 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
                    dataTrigger3.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                    style.Triggers.Add(dataTrigger3)
                    imageFactory1.SetValue(Image.StyleProperty, style)
                    Return imageFactory1
                End Function
            gridFactory.AppendChild(getIconFactory("IconAsync[16]"))
            gridFactory.AppendChild(getIconFactory("OverlayImageAsync[16]"))

            If [property].HasIcon Then
                Dim imageFactory2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                imageFactory2.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                imageFactory2.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                imageFactory2.SetValue(Image.SourceProperty, New Binding() With {
                    .Path = New PropertyPath("."),
                    .Mode = BindingMode.OneWay,
                    .IsAsync = True
                })

                Dim itemTemplate As DataTemplate = New DataTemplate()
                itemTemplate.VisualTree = imageFactory2

                Dim stackPanelFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(StackPanel))
                stackPanelFactory.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal)

                Dim itemsPanelTemplate As ItemsPanelTemplate = New ItemsPanelTemplate()
                itemsPanelTemplate.VisualTree = stackPanelFactory

                Dim itemsControlFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(ItemsControl))
                itemsControlFactory.SetValue(Grid.ColumnProperty, 2)
                itemsControlFactory.SetValue(ItemsControl.ItemsPanelProperty, itemsPanelTemplate)
                itemsControlFactory.SetValue(ItemsControl.ItemTemplateProperty, itemTemplate)
                itemsControlFactory.SetValue(ItemsControl.WidthProperty, New Binding() With {
                    .Path = New PropertyPath("StorageProviderUIStatusIconWidth16"),
                    .Mode = BindingMode.OneWay
                })
                itemsControlFactory.SetValue(ItemsControl.ItemsSourceProperty, New Binding() With {
                    .Path = New PropertyPath(String.Format("PropertiesByKeyAsText[{0}].Icons16Async", column.PROPERTYKEY.ToString())),
                    .Mode = BindingMode.OneWay,
                    .IsAsync = True
                })
                itemsControlFactory.SetValue(ItemsControl.MarginProperty, New Thickness(0, 0, 3, 0))

                gridFactory.AppendChild(itemsControlFactory)
            End If

            Dim textBlockFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(TextBlock))
            textBlockFactory.SetValue(Grid.ColumnProperty, 3)
            textBlockFactory.SetValue(Grid.ColumnSpanProperty, 2)
            textBlockFactory.SetValue(TextBlock.TextAlignmentProperty, column.Alignment)
            textBlockFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center)
            textBlockFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis)
            textBlockFactory.SetValue(TextBlock.PaddingProperty, New Thickness(0, 0, 2, 0))
            textBlockFactory.SetValue(TextBlock.TagProperty, "PART_DisplayName")
            textBlockFactory.SetValue(TextBlock.TextProperty, New Binding() With {
                .Path = New PropertyPath(If(column.CanonicalName = "System.ItemNameDisplay", "DisplayName", String.Format("PropertiesByKeyAsText[{0}].Text", column.PROPERTYKEY.ToString()))),
                .Mode = BindingMode.OneWay
            })
            Dim textBlockStyle As Style = New Style(GetType(TextBlock))
            textBlockStyle.Setters.Add(New Setter(TextBlock.ForegroundProperty, Brushes.Black))
            textBlockStyle.Setters.Add(New Setter(TextBlock.OpacityProperty, Convert.ToDouble(1)))
            Dim textBlockDataTrigger1 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
            textBlockDataTrigger1.Setters.Add(New Setter(TextBlock.OpacityProperty, Convert.ToDouble(0.5)))
            textBlockStyle.Triggers.Add(textBlockDataTrigger1)
            Dim textBlockDataTrigger2 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCompressed") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
            textBlockDataTrigger2.Setters.Add(New Setter(TextBlock.ForegroundProperty, Brushes.Blue))
            textBlockStyle.Triggers.Add(textBlockDataTrigger2)
            textBlockFactory.SetValue(Image.StyleProperty, textBlockStyle)
            gridFactory.AppendChild(textBlockFactory)

            template.VisualTree = gridFactory
            Return template
        End Function

        Protected Overrides Sub ClearBinding()
            MyBase.ClearBinding()

            If Not Me.PART_Ext Is Nothing Then
                Me.PART_Ext.Folder = Nothing
            End If
            If Not Me.PART_ListBox Is Nothing Then
                CType(CType(Me.PART_ListBox, System.Windows.Controls.ListView).View, GridView).Columns.Clear()
            End If
        End Sub

        Protected Overrides Sub MakeBinding(folder As Folder)
            If Not Me.PART_Ext Is Nothing Then
                Me.PART_Ext.Folder = folder
            End If
            If Not Me.PART_ListBox Is Nothing Then
                Me.ColumnsIn = buildColumnsIn(folder)
            End If

            MyBase.MakeBinding(folder)
        End Sub

        Protected Overrides Sub Folder_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "ItemsSortPropertyName", "ItemsSortDirection"
                    If Not Me.PART_Ext Is Nothing Then
                        Me.PART_Ext.UpdateSortGlyphs()
                    End If
            End Select

            MyBase.Folder_PropertyChanged(s, e)
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            Dim textBlock As TextBlock = UIHelper.FindVisualChildren(Of TextBlock)(listBoxItem) _
                .FirstOrDefault(Function(b) Not b.Tag Is Nothing AndAlso b.Tag = "PART_DisplayName")
            point = Me.PointFromScreen(textBlock.PointToScreen(New Point(0, 0)))
            point.X += -2
            point.Y += -1
            size.Width = textBlock.ActualWidth + 4
            size.Height = textBlock.ActualHeight + 2
            textAlignment = TextAlignment.Left
            fontSize = textBlock.FontSize
        End Sub

        Public Property ColumnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
            Get
                Return GetValue(ColumnsInProperty)
            End Get
            Set(ByVal value As Behaviors.GridViewExtBehavior.ColumnsInData)
                SetCurrentValue(ColumnsInProperty, value)
            End Set
        End Property
    End Class
End Namespace