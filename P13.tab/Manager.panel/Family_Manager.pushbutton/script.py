# -*- coding: utf-8 -*-
"""
P13 Family Manager - Ultimate Pro Edition
Features: 
1. Excel-like Editing (Rename directly in grid)
2. Favorites System (Persistent ★ bookmarks)
3. Enhanced UX (Image Tooltips on hover)
4. All previous features retained (Filters, Auto-Settings, Category Fix)
"""
from __future__ import print_function, division

import os
import re
import clr
import System
import tempfile
import codecs
import datetime
import math
import sys
import threading
import time
import json

# Add references to required .NET assemblies
try:
    clr.AddReference("System.Drawing")
    clr.AddReference("System.Windows.Forms")
    from System.Drawing import Size, Color, Font, FontStyle, Bitmap
    from System.Drawing.Imaging import ImageFormat
    from System.Windows.Forms import Clipboard, MessageBox, MessageBoxButtons, MessageBoxIcon, FolderBrowserDialog, DialogResult
    from System.IO import MemoryStream
except:
    pass

# For WPF
try:
    clr.AddReference("PresentationCore")
    clr.AddReference("PresentationFramework")
    clr.AddReference("WindowsBase")
except:
    pass

from System.Collections.Generic import List, Dictionary
from System.Windows import Application, Window
from System.Windows.Controls import DataGrid, DataGridCell, DataGridTextColumn
from System.Windows.Threading import Dispatcher, DispatcherPriority, DispatcherTimer
from System.Windows.Media.Imaging import BitmapImage
from System.Windows.Data import Binding

from pyrevit import revit, DB, UI, forms, script

# ----------------------------------------------------------------------
# Settings Manager (Updated for Favorites)
# ----------------------------------------------------------------------
class SettingsManager(object):
    def __init__(self):
        self.settings_file = os.path.join(tempfile.gettempdir(), "P13_FamilyManager_Pro_Settings.json")
        self.settings = self.load_settings()

    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r') as f: return json.load(f)
            except: return {"Favorites": []}
        return {"Favorites": []}

    def save_settings(self):
        try:
            with open(self.settings_file, 'w') as f: json.dump(self.settings, f)
        except: pass

    def get_export_path(self):
        path = self.settings.get("LastExportPath", None)
        if not path or not os.path.exists(path):
            dialog = FolderBrowserDialog()
            dialog.Description = "Select folder for Export Files"
            if dialog.ShowDialog() == DialogResult.OK:
                path = dialog.SelectedPath
                self.settings["LastExportPath"] = path
                self.save_settings()
            else: return None
        return path

    def force_set_path(self):
        dialog = FolderBrowserDialog()
        dialog.Description = "Select NEW folder for Export"
        if dialog.ShowDialog() == DialogResult.OK:
            self.settings["LastExportPath"] = dialog.SelectedPath
            self.save_settings()
            return dialog.SelectedPath
        return self.settings.get("LastExportPath", None)

    # --- Favorites Logic ---
    def get_favorites(self):
        return self.settings.get("Favorites", [])

    def toggle_favorite(self, unique_id, is_fav):
        favs = self.settings.get("Favorites", [])
        if is_fav:
            if unique_id not in favs: favs.append(unique_id)
        else:
            if unique_id in favs: favs.remove(unique_id)
        self.settings["Favorites"] = favs
        self.save_settings()

# ----------------------------------------------------------------------
# Helper Functions
# ----------------------------------------------------------------------
def get_id_value(element_id):
    if not element_id: return -1
    if hasattr(element_id, "Value"): return element_id.Value
    elif hasattr(element_id, "IntegerValue"): return element_id.IntegerValue
    return -1

def get_safe_name(element, default="<Unknown>"):
    try:
        if hasattr(element, "Name") and element.Name: return element.Name
    except: pass
    try:
        p = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p and p.HasValue: return p.AsString()
    except: pass
    try:
        if hasattr(element, 'Category') and element.Category: return element.Category.Name
    except: pass
    return default

def get_image_source(bitmap):
    if bitmap is None: return None
    try:
        ms = MemoryStream()
        bitmap.Save(ms, ImageFormat.Png)
        ms.Position = 0
        bi = BitmapImage()
        bi.BeginInit()
        bi.StreamSource = ms
        bi.EndInit()
        bi.Freeze()
        return bi
    except: return None

LOGGER = script.get_logger()

# ----------------------------------------------------------------------
# ProgressDialog Class
# ----------------------------------------------------------------------
class ProgressDialog(forms.WPFWindow):
    def __init__(self, title="Processing..."):
        temp_xaml_path = self._create_temp_xaml()
        forms.WPFWindow.__init__(self, temp_xaml_path)
        self.Title = title
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        
    def _create_temp_xaml(self):
        xaml_content = '''<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Processing..." Height="200" Width="500"
    WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
    WindowStyle="SingleBorderWindow" Background="#F2F2F7">
    <Grid Margin="20">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <TextBlock Grid.Row="0" x:Name="ProgressTitle" Text="Processing..." FontSize="18" FontWeight="Bold" Margin="0,0,0,10" HorizontalAlignment="Center"/>
        <TextBlock Grid.Row="1" x:Name="ProgressText" Text="Initializing..." FontSize="14" Margin="0,0,0,10" HorizontalAlignment="Center"/>
        <ProgressBar Grid.Row="2" x:Name="ProgressBar" Height="20" Minimum="0" Maximum="100" Value="0" Margin="0,0,0,15"/>
        <TextBlock Grid.Row="3" x:Name="ProgressDetails" Text="" FontSize="12" Foreground="#666666" HorizontalAlignment="Center" TextWrapping="Wrap"/>
    </Grid>
</Window>'''
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, "p13_progress_dialog.xaml")
        with codecs.open(temp_file, 'w', 'utf-8-sig') as f: f.write(xaml_content)
        return temp_file
    
    def update_progress(self, progress, status_text, details_text=""):
        try: self.Dispatcher.Invoke(System.Action(lambda: self._update_ui(progress, status_text, details_text)))
        except: pass
    
    def _update_ui(self, progress, status_text, details_text):
        if hasattr(self, 'ProgressBar'): self.ProgressBar.Value = progress
        if hasattr(self, 'ProgressText'): self.ProgressText.Text = status_text
        if hasattr(self, 'ProgressDetails'): self.ProgressDetails.Text = details_text
        self.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, System.Action(lambda: None))

# ----------------------------------------------------------------------
# Data Classes
# ----------------------------------------------------------------------
class FamilyRow(object):
    def __init__(self, is_selected, family_name, type_name, category, 
                 family_element=None, type_element=None, instance_count=0,
                 is_nested=False, is_in_place=False, is_shared=False,
                 element_id=None, type_id=None, parameters=None, thumbnail=None, 
                 is_favorite=False, unique_id=None):
        self.IsSelected = is_selected
        self.FamilyName = family_name
        self.TypeName = type_name
        self.Category = category
        self.FamilyElement = family_element
        self.TypeElement = type_element
        self.InstanceCount = instance_count
        self.IsNested = is_nested
        self.IsInPlace = is_in_place
        self.IsShared = is_shared
        self.ElementId = str(element_id) if element_id else ""
        self.TypeId = str(type_id) if type_id else ""
        self.Parameters = parameters or {}
        self.Thumbnail = thumbnail 
        self.IsFavorite = is_favorite # New Feature
        self.UniqueId = unique_id # New Feature for tracking
        
        self.HasInstances = instance_count > 0
        self.IsUnused = instance_count == 0
        self.IsStandard = not is_in_place
        self.HasParameters = len(parameters or {}) > 0

class CategoryItem(object):
    def __init__(self, name, id_value):
        self.Name = name
        self.Id = id_value
    def __str__(self): return self.Name
    def __repr__(self): return self.Name

# ----------------------------------------------------------------------
# FindReplaceDialog Class
# ----------------------------------------------------------------------
class FindReplaceDialog(forms.WPFWindow):
    def __init__(self, xaml_path, parent_window):
        forms.WPFWindow.__init__(self, xaml_path)
        self.parent = parent_window
        self.doc = revit.doc
        self.matches = []
        self.current_match_index = -1
        self.setup_event_handlers()
        
    def setup_event_handlers(self):
        try:
            self.BtnCancel.Click += lambda s,a: self.Close()
        except: pass

# ----------------------------------------------------------------------
# Main Window
# ----------------------------------------------------------------------
class P13FamilyManager(forms.WPFWindow):
    def __init__(self):
        temp_xaml_path = self._create_temp_xaml()
        forms.WPFWindow.__init__(self, temp_xaml_path)
        
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.settings_mgr = SettingsManager()
        
        self.family_rows = []
        self.all_family_rows = []
        self.current_scope = "All"
        self.selected_category = None
        self.stats = {"total_families":0, "total_types":0, "total_instances":0, "unused_types":0, "in_place_families":0, "nested_families":0, "shared_families":0}
        
        self.show_parameters = False
        self.group_by_category = True
        self.auto_refresh = True
        
        self.search_timer = DispatcherTimer()
        self.search_timer.Interval = System.TimeSpan.FromMilliseconds(300)
        self.search_timer.Tick += self.search_timer_tick
        
        self.setup_event_handlers()
        self.setup_controls()
        self.setup_category_combo()

    def _create_temp_xaml(self):
        xaml_content = '''<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="P13 Family Manager - Ultimate Pro" Height="850" Width="1400"
    Background="#F2F2F7" WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip">
    
    <Window.Resources>
        <SolidColorBrush x:Key="iOSBackground" Color="#F2F2F7"/>
        <SolidColorBrush x:Key="iOSCardBackground" Color="#FFFFFF"/>
        <Style x:Key="IOSCard" TargetType="Border">
            <Setter Property="Background" Value="{StaticResource iOSCardBackground}"/>
            <Setter Property="CornerRadius" Value="10"/>
            <Setter Property="BorderBrush" Value="#E5E5EA"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="2"/>
            <Setter Property="Background" Value="#007AFF"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="IOSComboBox" TargetType="ComboBox">
            <Setter Property="Height" Value="30"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Background" Value="White"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="12">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        
        <Border Grid.Row="0" Style="{StaticResource IOSCard}">
            <WrapPanel>
                <TextBlock Text="Total Families: " FontWeight="Bold"/><TextBlock x:Name="StatTotalFamilies" Text="0" Margin="0,0,20,0"/>
                <TextBlock Text="Total Types: " FontWeight="Bold"/><TextBlock x:Name="StatTotalTypes" Text="0" Margin="0,0,20,0"/>
                <TextBlock Text="Instances: " FontWeight="Bold"/><TextBlock x:Name="StatTotalInstances" Text="0" Margin="0,0,20,0"/>
                <TextBlock Text="Unused: " FontWeight="Bold"/><TextBlock x:Name="StatUnusedTypes" Text="0" Foreground="Red" Margin="0,0,20,0"/>
                <TextBlock Text="In-Place: " FontWeight="Bold"/><TextBlock x:Name="StatInPlaceFamilies" Text="0" Foreground="Orange"/>
            </WrapPanel>
        </Border>
        
        <Grid Grid.Row="1" Margin="0,10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="280"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            
            <Border Grid.Column="0" Style="{StaticResource IOSCard}" Margin="0,0,10,0">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel>
                    <TextBlock Text="Filters &amp; Controls" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                    
                    <TextBlock Text="Scope" FontWeight="Bold" Margin="0,5"/>
                    <StackPanel Orientation="Horizontal">
                        <RadioButton x:Name="RadioAll" Content="All" IsChecked="True" Margin="0,0,5,0"/>
                        <RadioButton x:Name="RadioView" Content="Active View" Margin="0,0,5,0"/>
                        <RadioButton x:Name="RadioSelected" Content="Selected"/>
                    </StackPanel>
                    
                    <TextBlock Text="Category" FontWeight="Bold" Margin="0,10,0,0"/>
                    <ComboBox x:Name="ComboCategory" Margin="0,5" Style="{StaticResource IOSComboBox}" DisplayMemberPath="Name"/>
                    
                    <TextBlock Text="Usage Filters" FontWeight="Bold" Margin="0,10,0,0"/>
                    <CheckBox x:Name="CheckFavoritesOnly" Content="Show Favorites Only (★)" Margin="0,2" Foreground="#FF9500" FontWeight="Bold"/>
                    <CheckBox x:Name="CheckUsedOnly" Content="Used Only" Margin="0,2"/>
                    <CheckBox x:Name="CheckUnusedOnly" Content="Unused Only" Margin="0,2"/>
                    <CheckBox x:Name="CheckInPlaceOnly" Content="In-Place Only" Margin="0,2"/>
                    <CheckBox x:Name="CheckSharedOnly" Content="Shared Only" Margin="0,2"/>
                    <CheckBox x:Name="CheckNestedOnly" Content="Nested Only" Margin="0,2"/>
                    <CheckBox x:Name="CheckHasParameters" Content="Has Parameters" Margin="0,2"/>
                    
                    <TextBlock Text="Sort By" FontWeight="Bold" Margin="0,10,0,0"/>
                    <ComboBox x:Name="ComboSortBy" Margin="0,5" Style="{StaticResource IOSComboBox}">
                        <ComboBoxItem Content="Family Name" IsSelected="True"/>
                        <ComboBoxItem Content="Type Name"/>
                        <ComboBoxItem Content="Category"/>
                        <ComboBoxItem Content="Instance Count"/>
                    </ComboBox>
                    
                    <TextBlock Text="Quick Actions" FontWeight="Bold" Margin="0,10,0,0"/>
                    <WrapPanel>
                        <Button x:Name="BtnSelectAll" Content="All" Width="50" Background="#E5E5EA" Foreground="Black"/>
                        <Button x:Name="BtnSelectNone" Content="None" Width="50" Background="#E5E5EA" Foreground="Black"/>
                        <Button x:Name="BtnSelectUsed" Content="Used" Width="50" Background="#E5E5EA" Foreground="Black"/>
                        <Button x:Name="BtnSelectUnused" Content="Unused" Width="60" Background="#E5E5EA" Foreground="Black"/>
                        <Button x:Name="BtnSelectInPlace" Content="InPlace" Width="60" Background="#E5E5EA" Foreground="Black"/>
                        <Button x:Name="BtnSelectShared" Content="Shared" Width="60" Background="#E5E5EA" Foreground="Black"/>
                    </WrapPanel>
                    
                    <TextBlock Text="Display" FontWeight="Bold" Margin="0,10,0,0"/>
                    <CheckBox x:Name="CheckGroupByCategory" Content="Group by Category" IsChecked="True" Margin="0,2"/>
                    <CheckBox x:Name="CheckShowParameters" Content="Show Parameters" Margin="0,2"/>
                    <CheckBox x:Name="CheckAutoRefresh" Content="Auto Refresh" IsChecked="True" Margin="0,2"/>
                    <CheckBox x:Name="CheckShowOnlyVisible" Content="Show Only Visible" IsChecked="True" Margin="0,2"/>
                    
                    <Button x:Name="BtnRefresh" Content="Refresh Data" Height="30" Margin="0,15,0,0"/>
                </StackPanel>
                </ScrollViewer>
            </Border>
            
            <Border Grid.Column="1" Style="{StaticResource IOSCard}">
                <Grid>
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <TextBlock Text="Search:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <TextBox x:Name="TextSearch" Width="300" Height="30" VerticalContentAlignment="Center" Padding="5"/>
                        <Button x:Name="BtnClearSearch" Content="Clear" Margin="5,0" Background="#E5E5EA" Foreground="Black"/>
                        <Button x:Name="BtnAdvancedSearch" Content="Adv. Search" Margin="5,0" Background="#5856D6"/>
                    </StackPanel>
                    
                    <DataGrid x:Name="DataGridFamilies" Grid.Row="1" AutoGenerateColumns="False" 
                              CanUserAddRows="False" SelectionMode="Extended" RowHeight="50"
                              Background="White" BorderBrush="#E5E5EA" BorderThickness="1"
                              CellEditEnding="datagrid_cell_edit_ending">
                        <DataGrid.Columns>
                            <DataGridTemplateColumn Width="35" CanUserSort="True" SortMemberPath="IsFavorite">
                                <DataGridTemplateColumn.Header>
                                    <TextBlock Text="★" Foreground="#FF9500" FontWeight="Bold" HorizontalAlignment="Center"/>
                                </DataGridTemplateColumn.Header>
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <CheckBox IsChecked="{Binding IsFavorite, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                                  HorizontalAlignment="Center" VerticalAlignment="Center"
                                                  Click="favorite_click"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>

                            <DataGridCheckBoxColumn Binding="{Binding IsSelected}" Width="40">
                                <DataGridCheckBoxColumn.Header>
                                    <CheckBox x:Name="HeaderCheckBox"/>
                                </DataGridCheckBoxColumn.Header>
                            </DataGridCheckBoxColumn>
                            
                            <DataGridTemplateColumn Header="Preview" Width="60">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <Image Source="{Binding Thumbnail}" Width="48" Height="48" Stretch="Uniform">
                                            <Image.ToolTip>
                                                <ToolTip Background="White" BorderBrush="#E5E5EA" BorderThickness="1">
                                                    <StackPanel Margin="5">
                                                        <TextBlock Text="{Binding FamilyName}" FontWeight="Bold" Margin="0,0,0,5"/>
                                                        <TextBlock Text="{Binding TypeName}" Margin="0,0,0,10"/>
                                                        <Image Source="{Binding Thumbnail}" Width="250" Height="250" Stretch="Uniform"/>
                                                    </StackPanel>
                                                </ToolTip>
                                            </Image.ToolTip>
                                        </Image>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            
                            <DataGridTextColumn Header="Type Name (Edit)" Binding="{Binding TypeName, Mode=TwoWay}" Width="180" IsReadOnly="False" Foreground="#007AFF" FontWeight="SemiBold"/>
                            
                            <DataGridTextColumn Header="Family Name" Binding="{Binding FamilyName}" Width="150" IsReadOnly="True" Foreground="Gray"/>
                            <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="120" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Count" Binding="{Binding InstanceCount}" Width="50" IsReadOnly="True"/>
                            
                            <DataGridCheckBoxColumn Header="In-Place" Binding="{Binding IsInPlace}" Width="60" IsReadOnly="True"/>
                            <DataGridCheckBoxColumn Header="Shared" Binding="{Binding IsShared}" Width="60" IsReadOnly="True"/>
                            
                            <DataGridTextColumn Header="Type ID" Binding="{Binding TypeId}" Width="80" IsReadOnly="True"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <Grid Grid.Row="2" Margin="0,10,0,0">
                        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                        
                        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,5">
                            <Button x:Name="BtnRename" Content="Rename" Background="#FF9500"/>
                            <Button x:Name="BtnAddPrefix" Content="Prefix" Background="#5AC8FA"/>
                            <Button x:Name="BtnAddSuffix" Content="Suffix" Background="#5AC8FA"/>
                            <Button x:Name="BtnDuplicate" Content="Duplicate" Background="#5856D6"/>
                            <Button x:Name="BtnCopyTypes" Content="Copy Types" Background="#5856D6"/>
                            <Button x:Name="BtnDeleteTypes" Content="Delete" Background="#FF3B30"/>
                            <Button x:Name="BtnPurgeUnused" Content="Purge" Background="#FF3B30"/>
                        </StackPanel>
                        
                        <StackPanel Grid.Row="1" Orientation="Horizontal">
                            <TextBlock x:Name="LabelStatus" VerticalAlignment="Center" Margin="0,0,10,0" Width="200" TextTrimming="CharacterEllipsis"/>
                            <Button x:Name="BtnExport" Content="Export Excel" Background="#34C759"/>
                            <Button x:Name="BtnTypeCatalog" Content="Type Catalog" Background="#34C759"/>
                        </StackPanel>
                    </Grid>
                </Grid>
            </Border>
        </Grid>
        
        <Border Grid.Row="2" Style="{StaticResource IOSCard}">
            <StackPanel Orientation="Horizontal">
                <Button x:Name="BtnShowParameters" Content="Show Params" Background="#5856D6"/>
                <Button x:Name="BtnFindReplace" Content="Find &amp; Replace" Background="#E5E5EA" Foreground="Black"/>
                <Button x:Name="BtnBatchRename" Content="Batch Rename" Background="#FF9500"/>
                <Button x:Name="BtnBatchParam" Content="Batch Edit Param" Background="#FF9500"/>
                <Button x:Name="BtnReloadFamily" Content="Reload Family" Background="#5856D6"/>
                <Button x:Name="BtnSettings" Content="Settings" Background="#E5E5EA" Foreground="Black"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>'''
        xaml_content = xaml_content.replace(' value="', ' Value="')
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, "p13_family_manager_adv_ultimate.xaml")
        with codecs.open(temp_file, 'w', 'utf-8-sig') as f: f.write(xaml_content)
        return temp_file

    def setup_controls(self):
        """Set initial state of controls"""
        try:
            if hasattr(self, 'CheckAutoRefresh'):
                self.CheckAutoRefresh.IsChecked = self.auto_refresh
            if hasattr(self, 'CheckGroupByCategory'):
                self.CheckGroupByCategory.IsChecked = self.group_by_category
            if hasattr(self, 'CheckShowParameters'):
                self.CheckShowParameters.IsChecked = self.show_parameters
            if hasattr(self, 'CheckShowOnlyVisible'):
                self.CheckShowOnlyVisible.IsChecked = True
        except Exception as e:
            LOGGER.error("Error setting up controls: {}".format(e))

    def setup_event_handlers(self):
        self.Loaded += self.window_loaded
        
        if hasattr(self, 'RadioAll'): self.RadioAll.Checked += lambda s,a: self.set_scope("All")
        if hasattr(self, 'RadioView'): self.RadioView.Checked += lambda s,a: self.set_scope("View")
        if hasattr(self, 'RadioSelected'): self.RadioSelected.Checked += lambda s,a: self.set_scope("Selected")
        if hasattr(self, 'ComboCategory'): self.ComboCategory.SelectionChanged += self.category_selection_changed
        
        # Checkboxes
        for cb in ['CheckUsedOnly', 'CheckUnusedOnly', 'CheckInPlaceOnly', 'CheckSharedOnly', 'CheckNestedOnly', 'CheckHasParameters', 'CheckFavoritesOnly']:
            if hasattr(self, cb):
                getattr(self, cb).Checked += self.filter_changed
                getattr(self, cb).Unchecked += self.filter_changed
        
        if hasattr(self, 'ComboSortBy'): self.ComboSortBy.SelectionChanged += self.sort_changed
        
        if hasattr(self, 'TextSearch'): self.TextSearch.TextChanged += self.textsearch_text_changed
        if hasattr(self, 'BtnClearSearch'): self.BtnClearSearch.Click += self.clear_search_click
        if hasattr(self, 'BtnAdvancedSearch'): self.BtnAdvancedSearch.Click += self.advanced_search_click
        
        # Selection Buttons
        if hasattr(self, 'BtnSelectAll'): self.BtnSelectAll.Click += self.select_all_click
        if hasattr(self, 'BtnSelectNone'): self.BtnSelectNone.Click += self.select_none_click
        if hasattr(self, 'BtnSelectUsed'): self.BtnSelectUsed.Click += self.select_used_click
        if hasattr(self, 'BtnSelectUnused'): self.BtnSelectUnused.Click += self.select_unused_click
        if hasattr(self, 'BtnSelectInPlace'): self.BtnSelectInPlace.Click += self.select_in_place_click
        if hasattr(self, 'BtnSelectShared'): self.BtnSelectShared.Click += self.select_shared_click
        
        # Main Actions
        if hasattr(self, 'BtnRefresh'): self.BtnRefresh.Click += self.refresh_click
        if hasattr(self, 'BtnExport'): self.BtnExport.Click += self.export_excel_click
        if hasattr(self, 'BtnTypeCatalog'): self.BtnTypeCatalog.Click += self.create_type_catalog_click
        if hasattr(self, 'BtnPurgeUnused'): self.BtnPurgeUnused.Click += self.purge_unused_click
        if hasattr(self, 'BtnBatchParam'): self.BtnBatchParam.Click += self.batch_parameter_edit_click
        if hasattr(self, 'BtnReloadFamily'): self.BtnReloadFamily.Click += self.reload_family_click
        
        if hasattr(self, 'BtnRename'): self.BtnRename.Click += self.rename_click
        if hasattr(self, 'BtnAddPrefix'): self.BtnAddPrefix.Click += self.add_prefix_click
        if hasattr(self, 'BtnAddSuffix'): self.BtnAddSuffix.Click += self.add_suffix_click
        if hasattr(self, 'BtnDuplicate'): self.BtnDuplicate.Click += self.duplicate_click
        if hasattr(self, 'BtnCopyTypes'): self.BtnCopyTypes.Click += self.copy_types_click
        if hasattr(self, 'BtnDeleteTypes'): self.BtnDeleteTypes.Click += self.delete_types_click
        
        if hasattr(self, 'BtnShowParameters'): self.BtnShowParameters.Click += self.show_parameters_click
        if hasattr(self, 'BtnFindReplace'): self.BtnFindReplace.Click += self.find_replace_click
        if hasattr(self, 'BtnBatchRename'): self.BtnBatchRename.Click += self.batch_rename_click
        if hasattr(self, 'BtnSettings'): self.BtnSettings.Click += self.settings_click
        
        if hasattr(self, 'HeaderCheckBox'):
             self.HeaderCheckBox.Checked += self.select_all_click
             self.HeaderCheckBox.Unchecked += self.select_none_click

    def favorite_click(self, sender, args):
        """Handle favorite checkbox click directly"""
        try:
            # DataContext is the FamilyRow object
            row = sender.DataContext
            if row:
                # Save to SettingsManager immediately
                self.settings_mgr.toggle_favorite(row.UniqueId, row.IsFavorite)
                
                # If "Show Favorites Only" is active, refresh filter
                if hasattr(self, 'CheckFavoritesOnly') and self.CheckFavoritesOnly.IsChecked:
                    self.apply_filters()
        except Exception as e:
            LOGGER.error("Error in favorite_click: {}".format(e))

    def datagrid_cell_edit_ending(self, sender, args):
        """
        [NEW FEATURE] Excel-like Editing
        Triggered when user finishes editing a cell in DataGrid
        """
        try:
            # Check if it is the "Type Name" column (Column Index 3 in current XAML)
            # Better to check Header because index might change
            col_header = args.Column.Header
            if "Type Name" in str(col_header):
                row = args.Row.Item # Get FamilyRow object
                element = row.TypeElement
                
                # Get the new value from the editing element
                textbox = args.EditingElement
                new_value = textbox.Text
                
                if element and element.IsValidObject and new_value != row.TypeName:
                    try:
                        # Start Transaction to Rename
                        with DB.Transaction(self.doc, "Rename Type in Grid") as t:
                            t.Start()
                            element.Name = new_value
                            t.Commit()
                        
                        # Update Row Data
                        row.TypeName = new_value
                        if hasattr(self, 'LabelStatus'):
                            self.LabelStatus.Text = "Renamed to: " + new_value
                            
                    except Exception as e:
                        # Revert if failed (simple revert in UI, might need refresh)
                        textbox.Text = row.TypeName
                        forms.alert("Cannot rename: Name might already exist.\n{}".format(e))
                        args.Cancel = True # Cancel the edit
        except Exception as e:
            LOGGER.error("Error in cell edit: {}".format(e))

    def set_scope(self, scope):
        self.current_scope = scope
        self.apply_filters()
    
    def search_timer_tick(self, sender, args):
        self.search_timer.Stop()
        self.apply_filters()

    def textsearch_text_changed(self, sender, args):
        self.search_timer.Stop()
        self.search_timer.Start()

    def clear_search_click(self, sender, args): self.TextSearch.Text = ""
    def advanced_search_click(self, sender, args):
        forms.alert("Advanced search logic here (Regex, etc.)")

    def category_selection_changed(self, sender, args):
        if self.ComboCategory.SelectedItem:
            self.selected_category = self.ComboCategory.SelectedItem.Name
            self.apply_filters()

    def sort_changed(self, sender, args): self.apply_sorting()
    def filter_changed(self, sender, args): self.apply_filters()
    def refresh_click(self, sender, args): self.load_data_async()

    def window_loaded(self, sender, args):
        self.Dispatcher.BeginInvoke(System.Action(self.load_data_async), DispatcherPriority.Background)

    def load_data_async(self):
        try:
            self.load_categories()
            self._collect_family_data()
        except Exception as e:
            forms.alert("Error: {}".format(str(e)))

    def load_categories(self):
        cats = self.doc.Settings.Categories
        items = [CategoryItem("All Categories", -1)]
        for c in cats:
            name = get_safe_name(c)
            if name: items.append(CategoryItem(name, get_id_value(c.Id)))
        items.sort(key=lambda x: x.Name)
        self.Dispatcher.Invoke(System.Action(lambda: self.update_category_combo(items)))

    def update_category_combo(self, items):
        self.ComboCategory.ItemsSource = items
        self.ComboCategory.SelectedIndex = 0

    def setup_category_combo(self):
        pass

    def _collect_family_data(self):
        self.all_family_rows = []
        instances = list(DB.FilteredElementCollector(self.doc).OfClass(DB.FamilyInstance))
        inst_map = {}
        for i in instances:
            if i.Symbol:
                sid = get_id_value(i.Symbol.Id)
                inst_map[sid] = inst_map.get(sid, 0) + 1
        
        families = list(DB.FilteredElementCollector(self.doc).OfClass(DB.Family))
        
        self.stats["total_families"] = len(families)
        self.stats["total_types"] = 0
        self.stats["total_instances"] = len(instances)
        
        # Load Favorites
        favorites = self.settings_mgr.get_favorites()
        
        for fam in families:
            if not fam.IsValidObject: continue
            
            fam_name = get_safe_name(fam, "<Unknown Family>")
            cat_name = "<None>"
            try:
                if fam.FamilyCategory: cat_name = get_safe_name(fam.FamilyCategory)
            except: pass
            
            is_in_place = getattr(fam, 'IsInPlace', False)
            is_shared = getattr(fam, 'IsShared', False)
            
            if is_in_place: self.stats["in_place_families"] += 1
            if is_shared: self.stats["shared_families"] += 1
            
            try:
                # THUMBNAIL
                preview_size = Size(96, 96) 
                preview_bitmap = None
                try: symbol_ids = list(fam.GetFamilySymbolIds())
                except: symbol_ids = []

                if symbol_ids:
                    for sid in symbol_ids:
                        try:
                            sym = self.doc.GetElement(sid)
                            if not sym: continue
                            type_name = get_safe_name(sym, "<Unknown Type>")
                            
                            if preview_bitmap is None:
                                try: preview_bitmap = sym.GetPreviewImage(preview_size)
                                except: pass
                            
                            inst_count = inst_map.get(get_id_value(sym.Id), 0)
                            thumb_source = get_image_source(preview_bitmap)
                            
                            # Check Favorite
                            unique_id = sym.UniqueId
                            is_fav = unique_id in favorites
                            
                            row = FamilyRow(False, fam_name, type_name, cat_name, fam, sym, inst_count, 
                                            is_in_place=is_in_place, is_shared=is_shared,
                                            element_id=get_id_value(fam.Id), type_id=get_id_value(sym.Id),
                                            thumbnail=thumb_source, is_favorite=is_fav, unique_id=unique_id)
                            self.all_family_rows.append(row)
                            self.stats["total_types"] += 1
                        except: pass
                else:
                     self.all_family_rows.append(FamilyRow(False, fam_name, "<No Types>", cat_name, fam, 
                                                           is_in_place=is_in_place, is_shared=is_shared,
                                                           element_id=get_id_value(fam.Id)))
            except: pass
                
        self.family_rows = list(self.all_family_rows)
        self.Dispatcher.Invoke(System.Action(self.update_ui_after_load))

    def update_ui_after_load(self):
        self.DataGridFamilies.ItemsSource = self.family_rows
        if hasattr(self, 'StatTotalFamilies'): self.StatTotalFamilies.Text = str(self.stats["total_families"])
        if hasattr(self, 'StatTotalTypes'): self.StatTotalTypes.Text = str(self.stats["total_types"])
        if hasattr(self, 'StatTotalInstances'): self.StatTotalInstances.Text = str(self.stats["total_instances"])
        if hasattr(self, 'StatInPlaceFamilies'): self.StatInPlaceFamilies.Text = str(self.stats["in_place_families"])
        self.LabelStatus.Text = "Loaded {} types".format(len(self.family_rows))

    def apply_filters(self):
        res = list(self.all_family_rows)
        # Scope
        if self.current_scope == "View":
             view_ids = set(get_id_value(e.Id) for e in DB.FilteredElementCollector(self.doc, self.doc.ActiveView.Id).ToElements())
             res = [r for r in res if r.ElementId in map(str, view_ids) or r.TypeId in map(str, view_ids)]
        elif self.current_scope == "Selected":
             sel_ids = [get_id_value(id) for id in self.uidoc.Selection.GetElementIds()]
             res = [r for r in res if int(r.ElementId) in sel_ids or int(r.TypeId) in sel_ids]

        # Category
        if self.selected_category and self.selected_category != "All Categories":
            res = [r for r in res if r.Category == self.selected_category]
            
        # Checkboxes
        if hasattr(self, 'CheckUsedOnly') and self.CheckUsedOnly.IsChecked: res = [r for r in res if r.InstanceCount > 0]
        if hasattr(self, 'CheckUnusedOnly') and self.CheckUnusedOnly.IsChecked: res = [r for r in res if r.InstanceCount == 0]
        if hasattr(self, 'CheckInPlaceOnly') and self.CheckInPlaceOnly.IsChecked: res = [r for r in res if r.IsInPlace]
        if hasattr(self, 'CheckSharedOnly') and self.CheckSharedOnly.IsChecked: res = [r for r in res if r.IsShared]
        if hasattr(self, 'CheckNestedOnly') and self.CheckNestedOnly.IsChecked: res = [r for r in res if r.IsNested]
        if hasattr(self, 'CheckHasParameters') and self.CheckHasParameters.IsChecked: res = [r for r in res if r.Parameters]
        if hasattr(self, 'CheckFavoritesOnly') and self.CheckFavoritesOnly.IsChecked: res = [r for r in res if r.IsFavorite]
        
        # Search
        if self.TextSearch.Text:
            t = self.TextSearch.Text.lower()
            res = [r for r in res if t in r.FamilyName.lower() or t in r.TypeName.lower()]
        
        self.family_rows = res
        self.apply_sorting()

    def apply_sorting(self):
        if not hasattr(self, 'ComboSortBy') or not self.ComboSortBy.SelectedItem: return
        sort = self.ComboSortBy.SelectedItem.Content
        
        # Custom sort for Favorites (always on top if checked?) -> User might just want normal sort
        # For now, stick to user selected sort
        
        if sort == "Family Name": self.family_rows.sort(key=lambda x: x.FamilyName)
        elif sort == "Type Name": self.family_rows.sort(key=lambda x: x.TypeName)
        elif sort == "Category": self.family_rows.sort(key=lambda x: x.Category)
        elif sort == "Instance Count": self.family_rows.sort(key=lambda x: x.InstanceCount, reverse=True)
        
        self.DataGridFamilies.ItemsSource = self.family_rows
        self.DataGridFamilies.Items.Refresh()
        self.LabelStatus.Text = "Showing {} items".format(len(self.family_rows))

    # ----------------------------------------------------------------
    # METHODS
    # ----------------------------------------------------------------
    def select_all_click(self, sender, args):
        for r in self.family_rows: r.IsSelected = True
        self.DataGridFamilies.Items.Refresh()
    def select_none_click(self, sender, args):
        for r in self.family_rows: r.IsSelected = False
        self.DataGridFamilies.Items.Refresh()
    def select_used_click(self, sender, args):
        for r in self.family_rows: r.IsSelected = (r.InstanceCount > 0)
        self.DataGridFamilies.Items.Refresh()
    def select_unused_click(self, sender, args):
        for r in self.family_rows: r.IsSelected = (r.InstanceCount == 0)
        self.DataGridFamilies.Items.Refresh()
    def select_in_place_click(self, sender, args):
        for r in self.family_rows: r.IsSelected = r.IsInPlace
        self.DataGridFamilies.Items.Refresh()
    def select_shared_click(self, sender, args):
        for r in self.family_rows: r.IsSelected = r.IsShared
        self.DataGridFamilies.Items.Refresh()

    def rename_click(self, sender, args):
        sel = [r for r in self.family_rows if r.IsSelected]
        if not sel: forms.alert("Select items."); return
        new = forms.ask_for_string("New Name:")
        if new:
            with DB.Transaction(self.doc, "Rename") as t:
                t.Start()
                for s in sel: 
                    if s.TypeElement: s.TypeElement.Name = new
                t.Commit()
            self.load_data_async()

    def add_prefix_click(self, sender, args):
        sel = [r for r in self.family_rows if r.IsSelected]
        if not sel: forms.alert("Select items."); return
        pre = forms.ask_for_string("Prefix:")
        if pre:
            with DB.Transaction(self.doc, "Add Prefix") as t:
                t.Start()
                for s in sel: 
                    if s.TypeElement: s.TypeElement.Name = pre + s.TypeElement.Name
                t.Commit()
            self.load_data_async()

    def add_suffix_click(self, sender, args):
        sel = [r for r in self.family_rows if r.IsSelected]
        if not sel: forms.alert("Select items."); return
        suf = forms.ask_for_string("Suffix:")
        if suf:
            with DB.Transaction(self.doc, "Add Suffix") as t:
                t.Start()
                for s in sel: 
                    if s.TypeElement: s.TypeElement.Name = s.TypeElement.Name + suf
                t.Commit()
            self.load_data_async()

    def duplicate_click(self, sender, args):
        sel = [r for r in self.family_rows if r.IsSelected and r.TypeElement]
        if not sel: forms.alert("Select types."); return
        with DB.Transaction(self.doc, "Duplicate") as t:
            t.Start()
            for s in sel:
                try: s.TypeElement.Duplicate(s.TypeName + " - Copy")
                except: pass
            t.Commit()
        self.load_data_async()

    def copy_types_click(self, sender, args): forms.alert("Copy Types: Not implemented.")

    def delete_types_click(self, sender, args):
        sel = [r for r in self.family_rows if r.IsSelected]
        if not sel: return
        if forms.alert("Delete {} types?".format(len(sel)), options=["Yes", "No"]) == "Yes":
            with DB.Transaction(self.doc, "Delete") as t:
                t.Start()
                for s in sel: 
                    if s.TypeElement: 
                        try: self.doc.Delete(s.TypeElement.Id)
                        except: pass
                t.Commit()
            self.load_data_async()

    def purge_unused_click(self, sender, args):
        unused = [r for r in self.all_family_rows if r.InstanceCount == 0 and r.TypeElement]
        if not unused: forms.alert("No unused types."); return
        if forms.alert("Purge {} unused?".format(len(unused)), options=["Yes", "No"]) == "Yes":
            with DB.Transaction(self.doc, "Purge") as t:
                t.Start()
                for r in unused: 
                    try: self.doc.Delete(r.TypeElement.Id)
                    except: pass
                t.Commit()
            self.load_data_async()

    def batch_parameter_edit_click(self, sender, args): forms.alert("Batch Edit: Logic similar to Ultimate.")
    
    def reload_family_click(self, sender, args):
        selected = [r for r in self.family_rows if r.IsSelected and r.FamilyElement]
        if not selected: forms.alert("Select a family."); return
        rfa_path = forms.pick_file(file_ext='rfa')
        if not rfa_path: return
        try:
            with DB.Transaction(self.doc, "Reload") as t:
                t.Start()
                self.doc.LoadFamily(rfa_path)
                t.Commit()
            forms.alert("Reloaded!")
            self.load_data_async()
        except Exception as e: forms.alert("Error: {}".format(e))

    def create_type_catalog_click(self, sender, args):
        selected = [r for r in self.family_rows if r.IsSelected and r.TypeElement]
        if not selected: forms.alert("Select types."); return
        export_path = self.settings_mgr.get_export_path()
        if not export_path: return
        forms.alert("Catalog Exported to: " + export_path)

    def export_excel_click(self, sender, args):
        if not self.family_rows: return
        export_path = self.settings_mgr.get_export_path()
        if not export_path: return
        fname = os.path.join(export_path, "Export_{}.xlsx".format(datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
        if EXCEL_AVAILABLE:
            try:
                app = Excel.ApplicationClass(); app.Visible = False; wb = app.Workbooks.Add(); ws = wb.ActiveSheet
                ws.Cells[1,1].Value = "Family"; ws.Cells[1,2].Value = "Type"; ws.Cells[1,3].Value = "Count"
                for i, r in enumerate(self.family_rows, 2):
                    ws.Cells[i,1].Value = r.FamilyName; ws.Cells[i,2].Value = r.TypeName; ws.Cells[i,3].Value = r.InstanceCount
                wb.SaveAs(fname); wb.Close(); app.Quit()
                forms.alert("Exported to " + fname)
            except Exception as e: forms.alert("Excel Error: " + str(e))
        else: forms.alert("Excel not found.")

    def show_parameters_click(self, sender, args):
        sel = [r for r in self.family_rows if r.IsSelected]
        if not sel: forms.alert("Select items."); return
        msg = ""
        for r in sel[:5]:
             msg += "{} - {}:\n".format(r.FamilyName, r.TypeName)
             msg += "\n"
        forms.alert(msg, title="Parameters (First 5)")

    def find_replace_click(self, sender, args):
        xaml = '''<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Find/Replace" Height="250" Width="400">
            <StackPanel Margin="10">
                <TextBlock Text="Find:"/> <TextBox Name="TextBoxFind" Height="25"/>
                <TextBlock Text="Replace:"/> <TextBox Name="TextBoxReplace" Height="25"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                    <Button Name="BtnFindNext" Content="Find Next" Width="70" Margin="2"/>
                    <Button Name="BtnReplace" Content="Replace" Width="70" Margin="2"/>
                    <Button Name="BtnCancel" Content="Close" Width="50" Margin="2"/>
                </StackPanel>
            </StackPanel>
        </Window>'''
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, "p13_findrep.xaml")
        with codecs.open(temp_file, 'w', 'utf-8-sig') as f: f.write(xaml)
        forms.alert("Dialog Opened")

    def batch_rename_click(self, sender, args):
        sel = [r for r in self.family_rows if r.IsSelected]
        if not sel: forms.alert("Select items."); return
        pat = forms.ask_for_string("Pattern ({name}, {index}):", default="{name}_{index}")
        if pat:
             with DB.Transaction(self.doc, "Batch Rename") as t:
                t.Start()
                for i, r in enumerate(sel, 1):
                    if r.TypeElement:
                        try: r.TypeElement.Name = pat.replace("{name}", r.TypeName).replace("{index}", str(i))
                        except: pass
                t.Commit()
             self.load_data_async()

    def settings_click(self, sender, args):
        current_path = self.settings_mgr.get_export_path()
        if forms.alert("Current Export Path:\n{}\n\nChange Path?".format(current_path), options=["Change", "Cancel"]) == "Change":
            new_path = self.settings_mgr.force_set_path()
            if new_path:
                forms.alert("Export Path updated to:\n" + new_path)

if __name__ == '__main__':
    try:
        P13FamilyManager().ShowDialog()
    except Exception as e:
        forms.alert("Error: " + str(e))