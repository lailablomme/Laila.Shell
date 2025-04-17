Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Helpers
Imports System.ComponentModel
Imports System.Windows.Data
Imports System.Windows.Media
Imports Laila.Shell.Controls.Parts

Namespace Controls
    Public Class DetailsView
        Inherits BaseFolderView

        Private PART_Ext As Behaviors.GridViewShellBehavior

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(DetailsView), New FrameworkPropertyMetadata(GetType(DetailsView)))
        End Sub

        Protected Overrides Sub OnRequestBringIntoView(s As Object, e As RequestBringIntoViewEventArgs)
            If TypeOf e.OriginalSource Is ListViewItem AndAlso UIHelper.GetParentOfType(Of ListBox)(e.OriginalSource)?.Equals(Me.PART_ListBox) Then
                Dim item As ListViewItem = e.OriginalSource
                If Not item Is Nothing Then
                    Dim transform As GeneralTransform = item.TransformToAncestor(PART_ScrollViewer)
                    Dim itemRect As Rect = transform.TransformBounds(New Rect(0, 0, item.ActualWidth, item.ActualHeight))
                    Dim headerRowPresenter As GridViewHeaderRowPresenter = UIHelper.FindVisualChildren(Of GridViewHeaderRowPresenter)(Me.PART_ListBox)(0)

                    ' Check if item is outside the viewport and adjust scrolling, but only vertically
                    If itemRect.Top < headerRowPresenter.ActualHeight Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + itemRect.Top - headerRowPresenter.ActualHeight)
                    ElseIf itemRect.Bottom > PART_ScrollViewer.ViewportHeight Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + (itemRect.Bottom - PART_ScrollViewer.ViewportHeight - headerRowPresenter.ActualHeight))
                    End If
                    e.Handled = True
                End If
            ElseIf TypeOf e.OriginalSource Is SlidingExpander AndAlso UIHelper.IsAncestor(Me.PART_ListBox, e.OriginalSource) Then
                Dim parentItem As SlidingExpander = e.OriginalSource
                Dim buttonItem As Button = UIHelper.FindVisualChildren(Of Button)(parentItem)?(0)
                Dim item As Border = UIHelper.GetParentOfType(Of Border)(buttonItem)
                If Not item Is Nothing AndAlso UIHelper.IsAncestor(PART_ScrollViewer, item) Then
                    Dim transform As GeneralTransform = item.TransformToAncestor(PART_ScrollViewer)
                    Dim itemRect As Rect = transform.TransformBounds(New Rect(0, 0, item.ActualWidth, item.ActualHeight))
                    Dim headerRowPresenter As GridViewHeaderRowPresenter = UIHelper.FindVisualChildren(Of GridViewHeaderRowPresenter)(Me.PART_ListBox)(0)

                    ' Check if item is outside the viewport and adjust scrolling, but only vertically
                    If itemRect.Top < 0 Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + itemRect.Top - headerRowPresenter.ActualHeight)
                    ElseIf itemRect.Bottom > PART_ScrollViewer.ViewportHeight Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + (itemRect.Bottom - PART_ScrollViewer.ViewportHeight - headerRowPresenter.ActualHeight))
                    End If
                    e.Handled = True
                End If
            ElseIf Not If(TypeOf e.OriginalSource Is Expander, e.OriginalSource, UIHelper.GetParentOfType(Of Expander)(e.OriginalSource)) Is Nothing _
                AndAlso UIHelper.GetParentOfType(Of ListBox)(e.OriginalSource)?.Equals(Me.PART_ListBox) Then
                e.Handled = True
            End If
        End Sub

        Protected Overrides Sub PART_ListBox_Loaded()
            ' notify of sort/group by changes
            Me.PART_Ext = Microsoft.Xaml.Behaviors.Interaction.GetBehaviors(Me.PART_ListBox).FirstOrDefault(Function(b) TypeOf b Is Behaviors.GridViewExtBehavior)

            MyBase.PART_ListBox_Loaded()
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.DragViewStrategy = New VerticalDragViewStrategy(Me.PART_DragInsertIndicator, Me)

            AddHandler Me.PART_Grid.SizeChanged,
                Sub(s As Object, e As SizeChangedEventArgs)
                    updateClip()
                End Sub
            updateClip()
        End Sub

        Private Sub updateClip()
            PART_Grid.Clip = New RectangleGeometry(New Rect(0, 0, PART_Grid.ActualWidth, PART_Grid.ActualHeight))
        End Sub

        Public Function buildColumnsIn(folder As Folder) As Behaviors.GridViewExtBehavior.ColumnsInData
            Dim d As Behaviors.GridViewExtBehavior.ColumnsInData = New Behaviors.GridViewExtBehavior.ColumnsInData()
            d.ViewName = folder.FullPath
            d.Items = New List(Of GridViewColumn)()
            d.CanSort = folder.CanSort

            If Not folder.Columns Is Nothing Then
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
                    gvc.SetValue(Behaviors.GridViewExtBehavior.GroupByPropertyNameProperty, String.Format("PropertiesByKeyAsText[{0}].GroupByText", column.PROPERTYKEY.ToString()))

                    Dim isVisibleDescriptor As DependencyPropertyDescriptor =
                    DependencyPropertyDescriptor.FromProperty(Behaviors.GridViewExtBehavior.IsVisibleProperty, gvc.GetType())
                    isVisibleDescriptor.AddValueChanged(gvc,
                    Sub(s2 As Object, e2 As EventArgs)
                        column.IsVisible = gvc.GetValue(Behaviors.GridViewExtBehavior.IsVisibleProperty)
                    End Sub)

                    d.Items.Add(gvc)
                Next
            End If

            Return d
        End Function

        Private Function getCellTemplate(column As Column, [property] As [Property]) As DataTemplate
            Dim template As DataTemplate = New DataTemplate()

            Dim gridFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(Grid))
            Dim columnDefinition1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition1.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition1)
            Dim columnDefinition2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition2.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition2)
            Dim columnDefinition3 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition3.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Star))
            gridFactory.AppendChild(columnDefinition3)

            Dim getIconFactory As Func(Of String, FrameworkElementFactory) =
                Function(bindTo As String) As FrameworkElementFactory
                    Dim imageFactory1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                    imageFactory1.SetValue(Grid.ColumnProperty, 0)
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

            Dim getDefaultIconFactory As Func(Of String, Boolean, FrameworkElementFactory) =
                Function(bindTo As String, isFolder As Boolean) As FrameworkElementFactory
                    Dim imageFactory0 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                    imageFactory0.SetValue(Grid.ColumnProperty, 0)
                    imageFactory0.SetValue(Image.MarginProperty, New Thickness(0, 0, 4, 0))
                    imageFactory0.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                    imageFactory0.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                    imageFactory0.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                    imageFactory0.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                    imageFactory0.SetValue(Image.SourceProperty, New Binding() With {
                        .Path = New PropertyPath(bindTo),
                        .Mode = BindingMode.OneWay,
                        .Source = New ImageHelper()
                    })
                    Dim style0 As Style = New Style(GetType(Image))
                    style0.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Collapsed))
                    style0.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(1)))
                    Dim dataTrigger02 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsHidden") With
                                    {
                                        .Mode = BindingMode.OneWay
                                    },
                        .Value = True
                    }
                    dataTrigger02.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                    style0.Triggers.Add(dataTrigger02)
                    Dim dataTrigger03 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                    {
                                        .Mode = BindingMode.OneWay
                                    },
                        .Value = True
                    }
                    dataTrigger03.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                    style0.Triggers.Add(dataTrigger03)
                    Dim dataTrigger04 As MultiDataTrigger = New MultiDataTrigger()
                    Dim dataTrigger04Condition1 As Condition = New Condition() With {
                        .Binding = New Binding() With {
                            .Path = New PropertyPath("IconAsync[16]"),
                            .Mode = BindingMode.OneWay,
                            .IsAsync = True
                        },
                        .Value = Nothing
                    }
                    dataTrigger04.Conditions.Add(dataTrigger04Condition1)
                    Dim dataTrigger04Condition2 As Condition = New Condition() With {
                        .Binding = New Binding() With {
                            .Path = New PropertyPath("IsFolder"),
                            .Mode = BindingMode.OneWay
                        },
                        .Value = isFolder
                    }
                    dataTrigger04.Conditions.Add(dataTrigger04Condition2)
                    Dim dataTrigger04Condition3 As Condition = New Condition() With {
                        .Binding = New Binding(String.Format("ColumnIndexFor[PropertiesByKeyAsText[{0}].Value]", column.PROPERTYKEY.ToString())) With
                                    {
                                        .ElementName = "ext",
                                        .Mode = BindingMode.OneWay
                                    },
                        .Value = 0
                    }
                    dataTrigger04.Conditions.Add(dataTrigger04Condition3)
                    dataTrigger04.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Visible))
                    style0.Triggers.Add(dataTrigger04)
                    style0.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Collapsed))
                    imageFactory0.SetValue(Image.StyleProperty, style0)
                    Return imageFactory0
                End Function

            gridFactory.AppendChild(getDefaultIconFactory("DefaultFileIconSmall", False))
            gridFactory.AppendChild(getDefaultIconFactory("DefaultFolderIconSmall", True))
            gridFactory.AppendChild(getIconFactory("IconAsync[16]"))
            gridFactory.AppendChild(getIconFactory("OverlayImageAsync[16]"))

            If [property]?.HasIcon Then
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
                itemsControlFactory.SetValue(Grid.ColumnProperty, 1)
                itemsControlFactory.SetValue(ItemsControl.ItemsPanelProperty, itemsPanelTemplate)
                itemsControlFactory.SetValue(ItemsControl.ItemTemplateProperty, itemTemplate)
                itemsControlFactory.SetValue(ItemsControl.WidthProperty, New Binding() With {
                    .Path = New PropertyPath("StorageProviderUIStatusIconWidth16"),
                    .Mode = BindingMode.OneWay,
                    .IsAsync = True
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
            textBlockFactory.SetValue(Grid.ColumnProperty, 2)
            textBlockFactory.SetValue(Grid.ColumnSpanProperty, 2)
            textBlockFactory.SetValue(TextBlock.TextAlignmentProperty, column.Alignment)
            textBlockFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center)
            textBlockFactory.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Left)
            textBlockFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis)
            textBlockFactory.SetValue(TextBlock.MarginProperty, New Thickness(0, 0, -3, 0))
            If column.CanonicalName = "System.ItemNameDisplay" Then
                textBlockFactory.SetValue(TextBlock.TagProperty, "PART_DisplayName")
            End If
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
            Dim textBlockDataTrigger2 As MultiDataTrigger = New MultiDataTrigger()
            Dim textBlockDataTrigger2Condition1 As Condition = New Condition() With {
                .Binding = New Binding("DoShowEncryptedOrCompressedFilesInColor") With
                           {
                               .Mode = BindingMode.OneWay,
                               .Source = Me
                           },
                .Value = True
            }
            textBlockDataTrigger2.Conditions.Add(textBlockDataTrigger2Condition1)
            Dim textBlockDataTrigger2Condition2 As Condition = New Condition() With {
                .Binding = New Binding("IsCompressed") With
                           {
                               .Mode = BindingMode.OneWay
                           },
                .Value = True
            }
            textBlockDataTrigger2.Conditions.Add(textBlockDataTrigger2Condition2)
            textBlockDataTrigger2.Setters.Add(New Setter(TextBlock.ForegroundProperty, Brushes.Blue))
            textBlockStyle.Triggers.Add(textBlockDataTrigger2)
            Dim textBlockDataTrigger3 As MultiDataTrigger = New MultiDataTrigger()
            Dim textBlockDataTrigger3Condition1 As Condition = New Condition() With {
                .Binding = New Binding("DoShowEncryptedOrCompressedFilesInColor") With
                           {
                               .Mode = BindingMode.OneWay,
                               .Source = Me
                           },
                .Value = True
            }
            textBlockDataTrigger3.Conditions.Add(textBlockDataTrigger3Condition1)
            Dim textBlockDataTrigger3Condition2 As Condition = New Condition() With {
                .Binding = New Binding("IsEncrypted") With
                           {
                               .Mode = BindingMode.OneWay
                           },
                .Value = True
            }
            textBlockDataTrigger3.Conditions.Add(textBlockDataTrigger3Condition2)
            textBlockDataTrigger3.Setters.Add(New Setter(TextBlock.ForegroundProperty, New SolidColorBrush(ColorConverter.ConvertFromString("#269D27"))))
            textBlockStyle.Triggers.Add(textBlockDataTrigger3)
            Dim textBlockDataTrigger4 As MultiDataTrigger = New MultiDataTrigger()
            Dim textBlockDataTrigger4Condition1 As Condition = New Condition() With {
                .Binding = New Binding("IsDoubleClickToOpenItem") With
                           {
                               .Mode = BindingMode.OneWay,
                               .Source = Me
                           },
                .Value = False
            }
            textBlockDataTrigger4.Conditions.Add(textBlockDataTrigger4Condition1)
            Dim textBlockDataTrigger4Condition2 As Condition = New Condition() With {
                .Binding = New Binding("IsUnderlineItemOnHover") With
                           {
                               .Mode = BindingMode.OneWay,
                               .Source = Me
                           },
                .Value = False
            }
            textBlockDataTrigger4.Conditions.Add(textBlockDataTrigger4Condition2)
            textBlockDataTrigger4.Setters.Add(New Setter(TextBlock.TextDecorationsProperty, TextDecorations.Underline))
            textBlockStyle.Triggers.Add(textBlockDataTrigger4)
            Dim textBlockDataTrigger5 As MultiDataTrigger = New MultiDataTrigger()
            Dim textBlockDataTrigger5Condition1 As Condition = New Condition() With {
                .Binding = New Binding("IsDoubleClickToOpenItem") With
                           {
                               .Mode = BindingMode.OneWay,
                               .Source = Me
                           },
                .Value = False
            }
            textBlockDataTrigger5.Conditions.Add(textBlockDataTrigger5Condition1)
            Dim textBlockDataTrigger5Condition2 As Condition = New Condition() With {
                .Binding = New Binding("IsUnderlineItemOnHover") With
                           {
                               .Mode = BindingMode.OneWay,
                               .Source = Me
                           },
                .Value = True
            }
            textBlockDataTrigger5.Conditions.Add(textBlockDataTrigger5Condition2)
            Dim textBlockDataTrigger5Condition3 As Condition = New Condition() With {
                .Binding = New Binding("Tag") With
                           {
                               .ElementName = "Bd",
                               .Mode = BindingMode.TwoWay
                           },
                .Value = "MouseOver"
            }
            textBlockDataTrigger5.Conditions.Add(textBlockDataTrigger5Condition3)
            textBlockDataTrigger5.Setters.Add(New Setter(TextBlock.TextDecorationsProperty, TextDecorations.Underline))
            textBlockStyle.Triggers.Add(textBlockDataTrigger5)
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
            Dim width As Double = 250
            Dim listView As System.Windows.Controls.ListView = UIHelper.GetParentOfType(Of System.Windows.Controls.ListView)(listBoxItem)
            If Not listView Is Nothing Then
                Dim itemNameColumn As Column = Me.Folder.Columns.FirstOrDefault(Function(c) c.CanonicalName = "System.ItemNameDisplay")
                If Not itemNameColumn Is Nothing Then
                    Dim itemNameGridViewColumn As GridViewColumn = CType(listView.View, GridView).Columns _
                        .FirstOrDefault(Function(c) c.GetValue(Behaviors.GridViewExtBehavior.PropertyNameProperty) _
                            .Equals(String.Format("PropertiesByKeyAsText[{0}].Value", itemNameColumn.PROPERTYKEY.ToString())))
                    If Not itemNameGridViewColumn Is Nothing Then
                        width = itemNameGridViewColumn.ActualWidth
                    End If
                End If
            End If
            point = Me.PointFromScreen(textBlock.PointToScreen(New Point(0, 0)))
            point.X += -2
            point.Y += -1
            size.Width = width - 24
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