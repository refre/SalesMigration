   M S B A M L   `   `   `  ����   	Migration    Ista.Migration.Prices   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     -     � Title$  Prices��$	��750��$
��1000�� � 	Resources    xmlfile � XPath��%*   &	        � $ ListeDePrix/Articles��.��� .+��� $��Visible=�.+�&   System.Windows.Controls.DataGrid     AutoGenerateColumns .  $��Visible=�  	 items �� Source �� Source��%!   &	       7   ��  �� 8   ���� # [    Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ ID��   
Visibility$ Hidden=�1  +System.Windows.Controls.DataGridBoundColumn	   Binding	 �� 
 �� XPath$	
 @ID�� �� Mode$ TwoWay=�  $	 nom��	 �� $

 @nom��$ TwoWay=�  $ value��-   valueBox��	 �� $
 @value��$ TwoWay=�  $ relevé��	 �� $
 @releve��$ TwoWay=�  $ vente��	 �� $
 @vente��$ TwoWay=�  $ locationMax15��	 �� $
 @locationMax15��$ TwoWay=� �� 
GroupStyle (  "System.Windows.Controls.GroupStyle  �� -   $��btnSave��$
��Save��$��Right=�$��75��$��30��