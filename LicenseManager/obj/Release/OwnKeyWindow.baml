   M S B A M L   `   `   `  ���� M  ILicenseManager, Version=5.0.0.16384, Culture=neutral, PublicKeyToken=nullK*clr-namespace:HGM.Hotbird64.LicenseManagerHGM.Hotbird64.LicenseManager  /    )HGM.Hotbird64.LicenseManager.OwnKeyWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     <local*clr-namespace:HGM.Hotbird64.LicenseManager       � Title$   Keys found in the Registry��$
��Auto��$
��Auto�� � SizeToContent$ WidthAndHeight=� � 	Resources    	GridStyle �� 
TargetType �� Property �� Value��%,   &	        �� " ��.���  � ��$ 2q� J� CommandBindings )  #System.Windows.Input.CommandBinding  -   /   )HGM.Hotbird64.LicenseManager.OwnKeyWindow   
InstallKey   Command# Z .���� 	 �� Orientation$	 Vertical=�.+��� -   
 ��Name$
 
TopElement��$��Keys in the Registry��$��5,5,5,5q�.��&   System.Windows.Controls.DataGrid  -     Name$ DataGridKeys��$
��Auto��   AutoGenerateColumns .     GridLinesVisibility$	 All=�   HorizontalGridLinesBrush
 �����   VerticalGridLinesBrush
 �����   SelectionMode$ Single=�   SelectionUnit$ CellOrRowHeader=�   RowHeaderWidth$ 20Q� � DataContext ��  � ContextMenu �� .@�m�  m� Command$
 Copy��m� # Z    RowStyle �� )  #System.Windows.Controls.DataGridRow"  .���  �@^�ShowDuration$ 15000����  �@�ToolTip �� ���� *
HiveName+   Columns 0  *System.Windows.Controls.DataGridTextColumn  1  +System.Windows.Controls.DataGridBoundColumn   ElementStyle# [  ,  &System.Windows.Controls.DataGridColumn   Header$ Install Date��   SortDirection$ 	AscendingX�   SortMemberPath$ InstallDate��   
IsReadOnly .    Binding �� *DisplayDate+  # [  $ Product Name�� .  �� *ProductName+  # [  $# Key From Digital Product Id 3�� .  �� *Id3+  # [  $# Key From Digital Product Id 4�� .  �� *Id4+� $��	0,0,10,10q� � Orientation$ 
Horizontal=�$��Right=�.+��� -    ��Name$ InstallButton��$��
Check _Key��$��24��$��60��$��5,0,5,0q�$��5,0,5,0q�{�.  �� -   $ CancelButton��$��_Cancel��  �� IsCancel  . $��24��$��60��$��5,0,5,0q�$��5,0,5,0q�