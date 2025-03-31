'Imports Microsoft.SqlServer

'<Style TargetType = "Expander" x:Key = "lailaShell_GroupByExpanderStyle" >
'    <Setter Property="Margin" Value="0,0,-5,0"/>
'<Setter Property= "Template" >
'    <Setter.Value>
'        <ControlTemplate TargetType="Expander">
'<Grid x : Name = "MainGrid" >
'    <Grid.Tag>
'        <sys:Double>0</sys:Double>
'    </Grid.Tag>
'<Grid.Resources>
'<Converters:MultiplyConverter x : Key = "multiplyConverter" />
'<converters:MultiplyConverterGridLength x:Key="multiplyConverterGridLength"/>
'                </Grid.Resources>
'                <Grid.RowDefinitions>
'<RowDefinition Height = "Auto" />
'                    <RowDefinition x:Name="ContentRow">
'                        <RowDefinition.Tag>
'                            <sys:Double>0</sys:Double>
'                        </RowDefinition.Tag>
'                        <RowDefinition.Height>
'                            <MultiBinding Converter="{StaticResource multiplyConverterGridLength}">
'                                <Binding Path="ActualHeight" ElementName="contentPresenter"/>
'                                <Binding Path="Tag" RelativeSource="{RelativeSource Self}"/>
'                                <Binding Path="ActualHeight" RelativeSource="{RelativeSource Mode=FindAncestor, AncestorType={x:Type ListBox}}"/>
'                            </MultiBinding>
'                        </RowDefinition.Height>
'                    </RowDefinition>
'                </Grid.RowDefinitions>
'                <Border BorderBrush = "{TemplateBinding BorderBrush}"
'BorderThickness = "0" >
'<DockPanel>
'    <ToggleButton DockPanel.Dock="Left" Margin="12,0,0,0" x:Name="PART_ExpanderToggle"
'        Style="{StaticResource lailaShell_GroupByExpanderToggleButtonStyle}"
'        IsChecked="{Binding IsExpanded, RelativeSource={RelativeSource TemplatedParent}}">
'        <ContentPresenter ContentSource="Header" VerticalAlignment="Center"/>
'    </ToggleButton>
'</DockPanel>
'                </Border>
'                <Border Grid.Row="1" x:Name = "Content" Tag="1.0"
'                    VerticalAlignment = "Bottom" ClipToBounds="True">
'                    <Border.Height>
'<MultiBinding Converter = "{StaticResource multiplyConverter}" >
'                            <Binding Path="ActualHeight" ElementName="contentPresenter"/>
'<Binding Path = "Tag" RelativeSource="{RelativeSource Self}"/>
'                            <Binding Path = "ActualHeight" RelativeSource="{RelativeSource Mode=FindAncestor, AncestorType={x:Type ListBox}}"/>
'                        </MultiBinding>
'                    </Border.Height>
'                    <ContentControl Tag = "1.0" VerticalAlignment="Bottom" x:Name = "contentContainer" >
'                        <ContentControl.Height>
'                            <MultiBinding Converter="{StaticResource multiplyConverter}">
'                                <Binding Path="ActualHeight" ElementName="contentPresenter"/>
'                                <Binding Path="Tag" RelativeSource="{RelativeSource Self}"/>
'                                <Binding Path="ActualHeight" RelativeSource="{RelativeSource Mode=FindAncestor, AncestorType={x:Type ListBox}}"/>
'                            </MultiBinding>
'                        </ContentControl.Height>
'<ContentPresenter VerticalAlignment = "Bottom" x:Name = "contentPresenter" />
'                    </ContentControl>
'                </Border>
'            </Grid>
'            <ControlTemplate.Triggers>
'                <MultiDataTrigger>
'                    <MultiDataTrigger.Conditions>
'                        <Condition Binding="{Binding Path=Tag, ElementName=MainGrid}">
'                            <Condition.Value>
'                                <sys:Double>0</sys:Double>
'                            </Condition.Value>
'                        </Condition>
'                    </MultiDataTrigger.Conditions>
'                    <MultiDataTrigger.Setters>
'                        <Setter TargetName="ContentRow" Property="Tag">
'                            <Setter.Value>
'                                <sys:Double>1</sys:Double>
'                            </Setter.Value>
'                        </Setter>
'                    </MultiDataTrigger.Setters>
'                </MultiDataTrigger>
'                <MultiDataTrigger>
'                    <MultiDataTrigger.Conditions>
'                        <Condition Binding="{Binding IsChecked, ElementName=PART_ExpanderToggle}" Value="True"/>
'                        <!--<Condition Binding="{Binding Path=Tag, ElementName=ContentRow}">
'                                    <Condition.Value>
'                                        <sys:Double>0</sys:Double>
'                                    </Condition.Value>
'                                </Condition>-->
'                        <Condition Binding="{Binding Path=Tag, ElementName=MainGrid}">
'                            <Condition.Value>
'                                <sys:Double>0</sys:Double>
'                            </Condition.Value>
'                        </Condition>
'                    </MultiDataTrigger.Conditions>
'                    <MultiDataTrigger.Setters>
'                        <Setter TargetName="contentContainer" Property="Visibility" Value="Visible"/>
'                    </MultiDataTrigger.Setters>
'                </MultiDataTrigger>
'                <MultiDataTrigger>
'                    <MultiDataTrigger.Conditions>
'                        <Condition Binding="{Binding IsChecked, ElementName=PART_ExpanderToggle}" Value="False"/>
'                        <!--<Condition Binding="{Binding Path=Tag, ElementName=ContentRow}">
'                                    <Condition.Value>
'                                        <sys:Double>1</sys:Double>
'                                    </Condition.Value>
'                                </Condition>-->
'                        <Condition Binding="{Binding Path=Tag, ElementName=MainGrid}">
'                            <Condition.Value>
'                                <sys:Double>0</sys:Double>
'                            </Condition.Value>
'                        </Condition>
'                    </MultiDataTrigger.Conditions>
'                    <MultiDataTrigger.Setters>
'                        <Setter TargetName="contentContainer" Property="Visibility" Value="Collapsed"/>
'                    </MultiDataTrigger.Setters>
'                </MultiDataTrigger>
'                <MultiDataTrigger>
'                    <MultiDataTrigger.Conditions>
'                        <Condition Binding="{Binding IsChecked, ElementName=PART_ExpanderToggle}" Value="True"/>
'                        <Condition Binding="{Binding Path=Tag, ElementName=ContentRow}">
'                            <Condition.Value>
'                                <sys:Double>0</sys:Double>
'                            </Condition.Value>
'                        </Condition>
'                        <Condition Binding="{Binding Path=Tag, ElementName=MainGrid}">
'                            <Condition.Value>
'                                <sys:Double>1</sys:Double>
'                            </Condition.Value>
'                        </Condition>
'                    </MultiDataTrigger.Conditions>
'                    <MultiDataTrigger.EnterActions>
'                        <RemoveStoryboard BeginStoryboardName="collapsedb"/>
'                        <BeginStoryboard x:Name="expandsb">
'                            <Storyboard>
'                                <DoubleAnimation Storyboard.TargetName="ContentRow"
'                                    Storyboard.TargetProperty="Tag"
'                                    From="0" To="1"
'                                    Duration="0:0:0.0"/>
'                            </Storyboard>
'                        </BeginStoryboard>
'                    </MultiDataTrigger.EnterActions>
'                </MultiDataTrigger>
'                <MultiDataTrigger>
'                    <MultiDataTrigger.Conditions>
'                        <Condition Binding="{Binding IsChecked, ElementName=PART_ExpanderToggle}" Value="False"/>
'                        <Condition Binding="{Binding Path=Tag, ElementName=ContentRow}">
'                            <Condition.Value>
'                                <sys:Double>1</sys:Double>
'                            </Condition.Value>
'                        </Condition>
'                        <Condition Binding="{Binding Path=Tag, ElementName=MainGrid}">
'                            <Condition.Value>
'                                <sys:Double>1</sys:Double>
'                            </Condition.Value>
'                        </Condition>
'                    </MultiDataTrigger.Conditions>
'                    <MultiDataTrigger.EnterActions>
'                        <RemoveStoryboard BeginStoryboardName="expandsb"/>
'                        <BeginStoryboard x:Name="collapsedb">
'                            <Storyboard>
'                                <DoubleAnimation Storyboard.TargetName="ContentRow"
'                                    Storyboard.TargetProperty="Tag"
'                                    From="1" To="0"
'                                    Duration="0:0:0.0"/>
'                            </Storyboard>
'                        </BeginStoryboard>
'                    </MultiDataTrigger.EnterActions>
'                </MultiDataTrigger>
'            </ControlTemplate.Triggers>
'        </ControlTemplate>
'    </Setter.Value>
'        </Setter>
'    </Style>
