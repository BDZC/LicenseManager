   M S B A M L   `   `   `  ���� M  ILicenseManager, Version=5.0.0.16384, Culture=neutral, PublicKeyToken=nullK*clr-namespace:HGM.Hotbird64.LicenseManagerHGM.Hotbird64.LicenseManager  ,    &HGM.Hotbird64.LicenseManager.ExportIds   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     <local*clr-namespace:HGM.Hotbird64.LicenseManager       � Title$  Export GUIDs��$	��768��$
��1024�� ^� ShowDuration$ 30000��.��� -    �Name$ 
TopElement�� � RowDefinitions �� $
��Auto ��� $��* ��� $
��Auto � � ColumnDefinitions �� -    ��Name$ 
LeftColumn��$	��50* ��� -   $ MiddleColumn��$	��40* ��� -   $ RightColumn��$	��40* �.+��� -    ��Name$ GroupBoxGuidType��$��0��$��2��$��2,0,2,0q�$��	GUID Type��$��Stretch=�.��� $��4,4,2,4q�.+�	� -    	�Name$ RadioButtonSkuIds��$��0,2,0,2q�$��80��$��_SKU IDs�� d� 	IsChecked$
 TrueY�	 � ToolTip$!	 Click to export the SKU IDs��	� -   $ RadioButtonKmsIds��$y�	Collapsed=�$��_KMS IDs��$!	 Click to export the KMS IDs������ 
 �� ElementName$
 RadioButtonSkuIds�� �� Path$ Margin����� $
 RadioButtonSkuIds��$ Width��� -   $ GroupBoxOptions��$��0��$��1��$��2,0,2,0q�$��Options��$��Stretch=�.��� $��4,4,2,4q�.+��� -	    ��Name$ CheckBoxAddComment��$
 TrueY�$y�	Collapsed=�$��_Add string GUID��$`	 ZIf checked all guids are additionally included in their string representation as a comment������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� -
   $ CheckBoxMultiLine��$ FalseY�$y�Visible=�$��_Multi-line attributes��$:	 4If checked each attribute appears on a seperate line������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� -   $ CheckBoxBlankLines��$
 TrueY�$y�Visible=�$��Insert _blank lines��$C	 =If unchecked does not insert blank lines between XML elements������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� -   $ CheckBoxUseTabs��$
 TrueY�$y�Visible=�$��	Use _tabs��$@	 :If unchecked uses 4 spaces for indentation instead of tabs������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� -   $ CheckBoxIncludeSku��$
 TrueY�$y�	Collapsed=�$��Include _SKU IDs��$/	 )If unchecked SKU IDs will not be exported������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� -   $ CheckBoxIncludeBetaSku��$
 TrueY�$y�	Collapsed=�$��Include _Beta SKU IDs��$@	 :If unchecked Beta and Preview SKU IDs will not be exported������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� -   $ CheckBoxIncludeKms��$
 TrueY�$y�	Collapsed=�$��Include all _KMS IDs��$:	 4If unchecked only essential KMS IDs will be exported������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� -   $ CheckBoxIncludeApp��$
 TrueY�$y�	Collapsed=�$��Include A_PP IDs��$/	 )If unchecked APP IDs will not be exported������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� -   $ CheckBoxNoDescription��$ FalseY�$y�	Collapsed=�$��No _description��$8	 2If checked no product description will be included������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� $��0��$��0��$��2,0,2,0q�$��Format��.��� $��4,4,2,4q�.+�	� -   $ RadioButtonVlmcsd��$	��180��$��_vlmcsd��$I	 CClick for a format that can be pasted into the kms.c file of vlmcsd������ $
 RadioButtonSkuIds��$ Margin�	� -   $ RadioButtonPyKms��$��_py-kms classic��$f	 `Click for a format that can be pasted into the kmsBase.py file of the original py-kms by cyrozap������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width�	� -   $ RadioButtonGenericC��$��Generic _C/C++��$&	  Click for a generic C/C++ format������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width�	� -   $ RadioButtonGenericCs��$��_Generic C♯��$#	 Click for a generic C# format������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width�	� -   $ RadioButtonXml��$
 TrueY�$��_License Manager / py-kms��$M	 GClick for a KmsDataBase.xml file as used by License Manager and py-kms ������ $
 RadioButtonSkuIds��$ Margin����� $
 RadioButtonVlmcsd��$ Width��� $��1��$��3��$��0��$��Output��.���� -    ��Name$ TextBoxOutput��$��Consolas)�$��Stretch=�$��Stretch=� �� TextWrapping$ NoWrap=�$ �� HorizontalScrollBarVisibility$
 Auto=�" �� VerticalScrollBarVisibility$
 Auto=� �� 
IsReadOnly . $��4,4,2,4q�� $��8,0,8,4q�$��2��$��3��$��0��$��Center=�$��Stretch=�$
��Auto�� � Orientation$ 
Horizontal=�.+��� -    ��Name$ ButtonSaveAsAscii��$	��120��$��24��$��0,0,5,0q�$��Save To _File��$ 	 Saves the output to a file��� -    �Name$ EncodingPanel��$��Center=�$��Stretch=�$
��Auto��$ 
Horizontal=�.+��� $��	Encoding:��$��Center=�	� -   $ RadioButtonAnsi��$
 TrueY�$
��ANSI��$��Center=�$��0,3,8,0q�$3	 -Encodes the file in the current ANSI codepage��	� -   $ RadioButtonUtf8��$��UTF-8��$��Center=�$��0,3,8,0q�$	 Encodes the file in UTF-8��	� -   $ RadioButtonUtf16Le��$��	UTF-16 LE��$��Center=�$��0,3,8,0q�$.	 (Encodes the file in UTF-16 little-endian��