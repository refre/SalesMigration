   M S B A M L   `   `   `  ����  	Migration    Ista.Migration.Prices   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   -   5          � Title$  Prices��5   	   $	��750��6   $
��1000��6%    � 	Resources 5         xmlfile � XPath��%8   &	        � 5   
   $ ListeDePrix/Articles��6C   5      .��5      � .+�5	   
   �� $��Visible=�6   .+�5
      &   System.Windows.Controls.DataGrid     AutoGenerateColumns .  $��Visible=�63    5       	 items �� Source �� Source��%8   &	       7   �� 5       69   �� 8   5      5      ��5      �� 5      # [ 6      Columns 5      0  *System.Windows.Controls.DataGridTextColumn  5      ,  &System.Windows.Controls.DataGridColumn   Header$ ID��6T      
Visibility$ Hidden=�6`   1  +System.Windows.Controls.DataGridBoundColumn	   Binding	 6)   �� 
 �� XPath$	
 @ID�� �� Mode$ TwoWay=�6`     5      $	 nom��6U   	 6)   �� $

 @nom��$ TwoWay=�6U     5      $ value��6W   -   5        valueBox��	 5   )   �� $
 @value��$ TwoWay=�6f     5      $ relevé��6X   	 6)   �� $
 @releve��$ TwoWay=�6X     5      $ vente��6W   	 6)   �� $
 @vente��$ TwoWay=�6W     5      $ locationMax15��6_   	 6)   �� $
 @locationMax15��$ TwoWay=�6_    �� 
GroupStyle 5      (  "System.Windows.Controls.GroupStyle  5      5       �� 5"      -   5        $��btnSave��5"      $
��Save��6$   $��Right=�6I   $��75��6e   $��30��6p   5#      5$      5%      