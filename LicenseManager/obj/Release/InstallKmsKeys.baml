   M S B A M L   `   `   `  ���� M  ILicenseManager, Version=5.0.0.16384, Culture=neutral, PublicKeyToken=nullK*clr-namespace:HGM.Hotbird64.LicenseManagerHGM.Hotbird64.LicenseManager  1    +HGM.Hotbird64.LicenseManager.InstallKmsKeys   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     <local*clr-namespace:HGM.Hotbird64.LicenseManager       � Title$  Install Multiple GVLKs�� � SizeToContent$ WidthAndHeight=� � 
ResizeMode$ NoResize=� � 	Resources ��%    .��� -    �Name$ 
TopElement�� � RowDefinitions �� $��* ��� $
��Auto ��� $��24 �.+��� $��KMS Client Products��$��5,5,5,5q�$��0��$��0��.��&   System.Windows.Controls.DataGrid  -     Name$ DataGrid��$   HorizontalScrollBarVisibility$
 Auto=�   ColumnWidth5  /System.Windows.Controls.DataGridLengthConverter$
 Auto 	   GridLinesVisibility$	 
Horizontal=�
   HorizontalGridLinesBrush

 �����   VerticalGridLinesBrush
 �����   AutoGenerateColumns .  $
��Auto��$
��Auto��   SelectionMode$ Single=�   SelectionUnit$ CellOrRowHeader=�   HeadersVisibility$ Column=�   CanUserDeleteRows .     CanUserAddRows .   � DataContext ��    	CellStyle ��  �� 
TargetType*  $System.Windows.Controls.DataGridCell"   �� Triggers P�  P� Property �@ 
IsSelected P� Value . . ���  �� Property � �� �� Value �	#000000ff��  � �� �	#20000000 � ContextMenu �� .@�m�  m� Command$
 Copy��   Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ Product��   
Foreground
 �   �   
IsReadOnly .     SortMemberPath$  DisplayName��!   SortDirection$! 	AscendingX�1  +System.Windows.Controls.DataGridBoundColumn"   Binding" ��   -   PartialProductKeyColumn��
 �   �$ Installed GVLK�� . $  PartialProductKey��" �� *PartialProductKey+4  .System.Windows.Controls.DataGridTemplateColumn  $ Available GVLK�� . $
  Gvlk��#   CellTemplate# �� $ �� DataType$$ local:KmsLicense��.R��� 
���   �$��5,0,5,0q�% ^� ShowDuration$% 15000������ *Gvlk+���� *	GvlkColor+& � ToolTip& �� *GvlkToolTip+  -   LicenseStatusColumn��$ License Status�� . $  LicenseStatusText��# �� $$ local:KmsLicense��.R��� $��5,0,5,0q����� *LicenseStatusText+���� *LicenseColor+  -   SelectColumn��$ Install��# �� $$ local:KmsLicense��.R���  �� .@�m� $
 Copy��{�.  .+��� -   {��� *IsControlEnabled+' �� Mode$' TwoWay=�( �� UpdateSourceTrigger$( PropertyChanged=�) d� 	IsChecked) �� *IsCheckBoxChecked+$' TwoWay=�$( PropertyChanged=�y��� *CheckBoxVisibility+	� -   {��� *IsControlEnabled+$' TwoWay=�$( PropertyChanged=�) �� *IsRadioButtonChecked+$' TwoWay=�$( PropertyChanged=�y��� *RadioButtonVisibility+  $ Result��*   
Visibility$* 	Collapsed=�-   ResultColumn��$  InstallMessage��# �� $$ local:KmsLicense��.R��� $��5,0,5,0q�$% 30000������ *InstallMessage+���� *InstallMessageColor+& �� *InstallToolTip+�� + �� Orientation$+ 
Horizontal=�$��1��$��Right=�$��0,6,0,10q� ��%z   '��       �� " ��.���  � ��$ 24����  � ��$ 80����  � ��$ 0,0,10,0q���  � ��$ 8,0,8,0q�.+��� -	   , ��Name$, InstallButton��$��_Install GVLKs��- �� 	IsDefault- . �� -
   $, CancelButton��$��_Cancel��. �� IsCancel. . �� $��2�����#f0f0f0]� $��2��.+��� -   / ��Name$/ LabelStatus��$��5,0,0,0q�$��Center=�Ready� -   0 �Name$0 ProgressBar��1 � IsIndeterminate1 .  $y�	Collapsed=�$��5,4,5,4q�