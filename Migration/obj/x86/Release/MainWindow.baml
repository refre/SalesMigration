   M S B A M L   `   `   `  ����   	Migration    Ista.Migration.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml       � Title$  	Migration��$	��750��$
��1000��.��]� ��. .+�o� $	��Top=�.@�m� $��_File��.@�m� -   $��_Quit��m� $��_Prices��.@�m� -   $��_List��m� $��_Help��.@�m� $��_about���� $	��Top=�.+��� -    ��Name$ image1��$��51��$	��112��$��Right=�$
��None=�$	��Top=�$.��(/Migration;component/Images/IstaLogo.png��� -    �Name$ pathGrid��$	��150�� � ColumnDefinitions �� $	��420 ��� $	��420 ��� $��* �.+��� -    ��Name$ GpBoxMigration��$��Migration file��$
��Left=�$	��143��$��15,0,0,0q�$	��400��.���  �� $
��7.5* ��� $
��2.5* �.+��� -    ��Name$ txtMigration��$��8,15,8,0q�$��25��$	��Top=�{�.  �� -    ��Name$ btnMigration��$
��Load��$��25��$
��Left=�$	��Top=�$��75��$��	16,15,0,0q�$��1���� -   $ gpbOffresFiles��$��1��$��Offre files��$
��Left=�$	��143��$��10,0,0,0q�$	��400��.���  �� $��40 ��� $��* ��� $	��100 � � RowDefinitions �� $��* ��� $��* ��� $��* �.+��� -	   $��btnOffre��$
��Load��$��25��$��Right=�$��0,8,6,0q�$	��Top=�$��75��$��2���� -
   $��txtOffre��$��1��$��25��$��8,5,8,0q�{�.  �� -   $��lblBX��$��BX��$��28��$
��Left=�$��6,7,0,0q�$	��Top=��� -   $��lblAN��$��AN��$��1��$��28��$
��Left=�$��6,7,0,0q�$	��Top=��� -   $��lblVV��$��VV��$��2��$��28��$
��Left=�$��6,7,0,0q�$	��Top=��� -   $��
txtOffreAN��$��1��$��1��$��25��$
��Left=�$��8,8,0,0q�$	��Top=�$	��232��{�.  �� -   $��
txtOffreVV��$��1��$��2��$��25��$
��Left=�$��8,9,0,0q�$	��Top=�$	��232��{�.  �� -   $��	btnLoadAN��$
��Load��$��2��$��1��$��25��$
��Left=�$��19,7,0,0q�$	��Top=�$��75���� -   $��	btnLoadVV��$
��Load��$��2��$��2��$��25��$
��Left=�$��19,7,0,0q�$	��Top=�$��75���� -   $ 
btnCompare��$��Compare��$��23��$��2��$��	9,29,8,98q�{�.  �� $	��Top=�$��Comapraison Result��.���� $
��Auto=�$��50��.���� -   $��txtCompareResult��$��	10,5,5,10q�$
��Wrap=��� $��Bottom=�.+�&   System.Windows.Controls.DataGrid  -     Name$ dataMigration��	   AutoGenerateColumns	 .  �� $	��200��$��Center=�$��1,0,0,0q�$	��Top=�$
   HorizontalScrollBarVisibility$

 Auto=�   Columns 4  .System.Windows.Controls.DataGridCheckBoxColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ Physical Person��1  +System.Windows.Controls.DataGridBoundColumn   Binding ��  �� Path$ PhysicalPerson� �� Mode$ TwoWay=�0  *System.Windows.Controls.DataGridTextColumn  $
 Site��   
IsReadOnly .  �� $
 Site�  $ Group Number�� .  �� $ NumeroDeGroupe�  $ 
Group name�� .  �� $ NomDeGroupe�  $ Group address�� .  �� $ AdresseGroupe1�  $ Group Zip code�� .  �� $ CodePostauxGroupe�  $
 City�� .  �� $ LocaliteImmeuble�  $ Holder number�� .  �� $ NumeroDuGerant�  $ Holder name�� .  �� $ NomDeGerant�  $ Holder address�� .  �� $ AdresseGerant1�  $ Holder zip code�� .  �� $ CodePostauxGerant�  $ Language�� .  �� $ Langue�  $ Holder city�� .  �� $ LocaliteGerant�  $ Heater�� .  �� $ ChauffageNombreNRad�  $ 	Hot water�� .  �� $ EauChaudeNombreNRad�  $ 
Cold water�� .  �� $ EauFroideNombreNRad�  $ Integrators�� .  �� $ IntegrateurNombreNRad��� .+��� $��
Item price��$	��175��.���  �� $��25 ��� $��25 ��� $��25 ��� $��25 ��� $��25 ��� $��25 � �� $	��130 ��� $��75 ��� $��75 ��� $��75 ��� $��75 ��� $	��130 ��� $��75 ��� $��75 ��� $��75 ��� $��75 ��� $��* �.+��� -   $��txtQuantityLeft��$��1��$��Quantity���� -   $��txtPriceLeft��$��2��$��
Sell price���� -   $��txtRentLeft��$��3��$��
Rent price���� -   $��txtReadLeft��$��4��$��
Read price���� -   $��txtQuantityCenter��$��6��$��Quantity���� -   $��txtPriceCenter��$��7��$��
Sell price���� -   $��txtRentCenter��$��8��$��
Rent price���� -   $��txtReadCenter��$��9��$��
Read price���� -   $��txtDoprimoIIIRadio��$��Doprimo III Radio:��$��1���� -   $��txtDoprimoIIIRQuant��$��23��$
��Left=�$	��Top=�$��75��$��1��$��1���� -   $��txtDoprimoIIIRPrice��$��23��$
��Left=�$	��Top=�$��75��$��2��$��1���� -    $��txtDoprimoIIIRRent��$��23��$
��Left=�$	��Top=�$��75��$��3��$��1���� -!   $��txtDoprimoIIIRRead��$��23��$
��Left=�$	��Top=�$��75��$��4��$��1���� -"   $��txtDomaquaR��$��3��$��23��$
��Left=�$��Domaqua Radio (cold):��$	��Top=�$	��130���� -#   $��txtDomaquaRQuant��$��23��$
��Left=�$	��Top=�$��75��$��1��$��3���� -$   $��txtDomaquaRPrice��$��23��$
��Left=�$	��Top=�$��75��$��2��$��3���� -%   $��txtDomaquaRRent��$��23��$
��Left=�$	��Top=�$��75��$��3��$��3���� -&   $��txtDomaquaRRead��$��23��$
��Left=�$	��Top=�$��75��$��4��$��3���� -'   $��txtDomaquaRHot��$��5��$��3��$��23��$
��Left=�$��Domaqua Radio (hot):��$	��Top=�$	��120���� -(   $��txtDomaquaRHotQuant��$��23��$
��Left=�$	��Top=�$��75��$��6��$��3���� -)   $��txtDomaquaRHotPrice��$��23��$
��Left=�$	��Top=�$��75��$��7��$��3���� -*   $��txtDomaquaRHotRent��$��23��$
��Left=�$	��Top=�$��75��$��8��$��3���� -+   $��txtDomaquaRHotRead��$��23��$
��Left=�$	��Top=�$��75��$��9��$��3���� -,   $��txtDomaquaTot��$��4��$��23��$
��Left=�$��Domaqua Total:��$	��Top=�$	��130���� --   $��txtDomaquaTotQuant��$��1��$��4�� �� 
IsReadOnly . $��23��$
��Left=�$	��Top=�$��75���� -.   $��txtSenso1_2R��$��23��$
��Left=�$��Sensonic 1/2 Radio:��$	��Top=�$	��130��$��0��$��2���� -/   $��txtSenso1_2RQuant��$��1��$��2��$��23���� -0   $��txtSenso1_2RPrice��$��2��$��2��$��23���� -1   $��txtSenso1_2RRent��$��3��$��2��$��23���� -2   $��txtSenso1_2RRead��$��4��$��2��$��23���� -3   $��txtSenso3_4R��$��23��$
��Left=�$��Sensonic 3/4 Radio:��$	��Top=�$	��130��$��5��$��2���� -4   $��txtSensoR3_4Quant��$��6��$��2��$��23���� -5   $��txtSensoR3_4Price��$��7��$��2��$��23���� -6   $��txtSensoR3_4Rent��$��8��$��2��$��23���� -7   $��txtSensoR3_4Read��$��9��$��2��$��23���� -8   $��btnSavePrices��$
��save��$��5��$��23��$
��Left=�$	��Top=�$	��100����  �� Orientation$ 
Horizontal=�.+��� -9   $ btnWord��$��Building selected��$��27��$
��Left=�$��1,0,0,0q�$	��Top=�$	��120���� -:   $ btnAll��$	��All��$��27��$
��Left=�$��10,0,0,0q�$	��Top=�$��75���� -;   $ btnClear��$y�	Collapsed=�$��Clear��$��27��$
��Left=�$��20,0,0,0q�$	��Top=�$��75��{�.  �� -<   $ btnClole��$��Close��$��27��$��	680,0,0,0q�$	��Top=�$��75�� �� IsCancel . 