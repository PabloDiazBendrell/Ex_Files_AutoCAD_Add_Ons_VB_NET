   M S B A M L   `   `   `  ����L  HMyFirstPlugIn, Version=1.0.0.27031, Culture=neutral, PublicKeyToken=null;"clr-namespace:MyCompany.MySoftwareMyCompany.MySoftware  5   	   '    !MyCompany.MySoftware.winLayerList   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   4local"clr-namespace:MyCompany.MySoftware     5   	     � Title$  
Layer List��5   	   $	��300��6   $	��300��6)    � 
ResizeMode$ NoResize=�65    � ShowInTaskbar .   � WindowStartupLocation$ CenterOwner=�6a   .��5	      �  � RowDefinitions 5
   
   �� 5      $	��31* �6   �� 5      $��32 �6   .+�5   
   �� ��6   �� ��5      �� 5      .R�5      �  � ColumnDefinitions 5      �� 5      $	��10* �6/   �� 5      $
��200* �6/   .+�5      �� $��Center=�6#   $��Center=�6>    d� 	IsChecked 5   #   ��  �� Path$ 
IsSelected� �� UpdateSourceTrigger$ PropertyChanged=��� 5      $��1��6$   $��Center=�64   $��10,2,0,2q�6j   ��6O   �� $
 Name�6j   5      5      5      �� 5   
   $��1��6   $��Right=�6"   	 �� Orientation$	 
Horizontal=�6>   .+�5      �� -   5        $��btn_Ok��5   I   $��76��6   $��4q�6    $��Ok��6+   
 �� 	IsDefault
 . 6I   �� 5      -   5        $��
btn_Cancel��5   L   $��76��6   $��4q�6    $��Cancel��6+    �� IsCancel . 6L   5       5#      5$      