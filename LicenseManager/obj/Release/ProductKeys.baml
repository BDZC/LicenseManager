   M S B A M L   `   `   `  ���� M  ILicenseManager, Version=5.0.0.16384, Culture=neutral, PublicKeyToken=nullK*clr-namespace:HGM.Hotbird64.LicenseManagerHGM.Hotbird64.LicenseManager  .    (HGM.Hotbird64.LicenseManager.ProductKeys   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     9lm*clr-namespace:HGM.Hotbird64.LicenseManager       � Title$  Install a Generic Key��$	��600��$	��450��.��� -    �Name$ 
TopElement�� � RowDefinitions �� $��* ��� $
��Auto �.+�]� $��5,5,5,5q�.+��� $	��Top=�$��Choose a License Key��R� -    R�Name$ ProductTree��$��0,2,0,0q�����  �� 
TargetType" Q�.�-   �� $��5,5,5,5q�$��1��.+��� -   $��TextBlockGenerated��$
��Wrap=�
���  ��$y�	Collapsed=�][Warning! This key is user-generated and cannot be used to acquire a Windows Genuine Ticket.�� $
��Wrap=�$��0,7,0,2q�Product Key�� -    ��Name$ GvlkKey��   StandardTextBox#��[  ����  �� ElementName$ ProductTree�� �� Path$+ %SelectedItem.(lm:ProductKeys.ToolTip)�� $��0,5,0,5q�$��Right=�.+��� -    ��Name$ InstallButton��$��5,5,5,5q�$��70��	 �� 	IsDefault	 . {�.  .��Check�� -   $��5,5,0,5q�$��70��
 �� IsCancel
 . .��Cancel�� $��0,5,0,0q� �� TextAlignment$ Center=�$��Center=�$
��Wrap=�JHPlease check for updated or additional Product Keys at Microsoft Technet� $��Center=�.+��� $ Center=�$��	15,0,15,0q�$��11(��� -   
����  �$O��Ihttps://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys?�.�	Windows�� $ Center=�$��	15,0,15,0q�$��11(��� -	   
����  �$V��Phttp://technet.microsoft.com/en-us/library/ee624355(v=office.14).aspx#section2_3?�.�
Office2010�� $ Center=�$��	15,0,15,0q�$��11(��� -
   
����  �$>��8http://technet.microsoft.com/en-us/library/dn385360.aspx?�.�
Office2013�� $ Center=�$��	15,0,15,0q�$��11(��� -   
����  �$F��@https://docs.microsoft.com/en-us/DeployOffice/vlactivation/gvlks?�.�Office2016/2019