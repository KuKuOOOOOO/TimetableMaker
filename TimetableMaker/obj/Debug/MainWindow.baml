   M S B A M L   `   `   `  ����I  ETimetableMaker, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null/clr-namespace:TimetableMakerTimetableMaker  5   	   E'clr-namespace:TimetableMaker.ViewModelsTimetableMaker.ViewModels  5   	       TimetableMaker.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   .localclr-namespace:TimetableMaker     5   	   X TXceed.Wpf.Toolkit, Version=4.1.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f44xctk)http://schemas.xceed.com/wpf/xaml/toolkit  5   	   0vm'clr-namespace:TimetableMaker.ViewModels  5   	     � Title$  TimetableMaker��5
   	   $	��450��6    $	��800��6-    � DataContext 5      2  ,TimetableMaker.ViewModels.TimetableViewModel  5   
   .��5      � $��0,0,0,30q�6   .+�5   
   � -   5         �Name$ Grid1��5      $
��Left=�6   $	��300��69   $��56,69,0,-246q�6F   $	��Top=�6\   $	��641��6t    � RowDefinitions 5      �� 5      $��40 �6    6.   �� 5      $��40 �6    6.   �� 5      $��60 �6    6.   �� 5      $��60 �6    6.    � ColumnDefinitions 5      �� 5      $	��150 �6#   61   �� 5      6%   .+�5      �� $��老師姓名��6   $��0��6$   $��0��61    �� TextAlignment$ Center=�6A   $��Center=�6X   $��24(�6s   
����� �6�   �� 5      $��課程名稱��6   $��1��6$   $��0��61   $ Center=�6A   $��Center=�6X   $��24(�6s   
����� �6�   �� 5      $��課程開始時間��6   $��2��6&   $��0��63   $ Center=�6C   $��Center=�6Z   $��24(�6u   
����� �6�   �� 5      $��課程結束時間��6   $��3��6&   $��0��63   $ Center=�6C   $��Center=�6Z   $��24(�6u   
����� �6�   �� 5      $��0��64   $��1��6A    �� TextAlignment$ Center=�6Q   $��Center=�6h   $��20(�6�   ��6   �� *TeacherName+6�   �� 5      $��1��61   $��1��6>   $ Center=�6N   $��Center=�6e   $��20(�6�   ��6   �� *	ClassName+6�   &   Xceed.Wpf.Toolkit.DateTimePicker  5       &   Xceed.Wpf.Toolkit.DateTimeUpDown   Format$ Custom=�6"      FormatString$ yyyy/MM/dd HH:mm��62   $��2��6n   $��1��6{   ,  &Xceed.Wpf.Toolkit.Primitives.InputBase	   TextAlignment$	 Center=�6�   $��Center=�6�   $��20(�6�   �  �Xceed.Wpf.Toolkit.Primitives.UpDownBase`1[[System.Nullable`1[[System.DateTime, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089]], mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089]]
   Value
 6R   �� *	StartTime+6�     5!      $ Custom=�6"   $ yyyy/MM/dd HH:mm��62   $��3��6l   $��1��6y   $	 Center=�6�   $��Center=�6�   $��20(�6�   
 6R   �� *	EndTime+6�   5"      �� 5#   
   $��新增課程��6   $
��Left=�6D   $��45��6_   $��24(�6k   $��
56,299,0,0q�6y   $	��Top=�6�   $	��115��6�    J� RenderTransformOrigin$ 1.227,0.233&�6�   ��6    �� *AddClassCommand+6�   �� 5$   
   $��課表匯出��6   $
��Left=�6G   $��45��6b   $��24(�6n   $��582,299,0,0q�6|   $	��Top=�6�   $	��115��6�   $ 1.227,0.233&�6�   ��6    �� *ExportExcelCommand+6�   �� 5%   
   $��預覽課表��6   $
��Left=�6C   $��45��6^   $��24(�6j   $��406,299,0,0q�6x   $	��Top=�6�   $	��115��6�   $ 1.227,0.233&�6�   ��6    �� *PreviewCommand+6�   �� 5&   
   $��讀取課表��6   $
��Left=�6C   $��45��6^   $��24(�6j   $��231,299,0,0q�6x   $	��Top=�6�   $	��115��6�   $ 1.227,0.233&�6�   ��6    �� *LoadingCommand+6�   %  Xceed.Wpf.Toolkit.BusyIndicator  5'   
   $
��Left=�6X   $	��423��6s   $��18(�6�   $��	0,-34,0,0q�6�   $	��Top=�6�   $	��790��6�   $ 1.227,0.233&�6�      BusyContent 6   �� *
BusyText+   IsBusy 6>   �� *isBusy+6�   5(      5)      