FasdUAS 1.101.10   ��   ��    k             l        	  x     �� 
 ��   
 1      ��
�� 
ascr  �� ��
�� 
minv  m         �    2 . 4��       Yosemite (10.10) or later    	 �   4   Y o s e m i t e   ( 1 0 . 1 0 )   o r   l a t e r      x    �� ����    2  	 ��
�� 
osax��        l          x    "��  ��    4  �� 
�� 
scpt  m  
   �    C a l e n d a r L i b   E C  �� ��
�� 
minv  m         �   
 1 . 1 . 5��    * $ put this at the top of your scripts     �     H   p u t   t h i s   a t   t h e   t o p   o f   y o u r   s c r i p t s   ! " ! l     ��������  ��  ��   "  # $ # l     �� % &��   % � � This script converts an Excel timetable to Apple Calendar. Must open the excel timetable in Numbers first, then run this script    & � ' '    T h i s   s c r i p t   c o n v e r t s   a n   E x c e l   t i m e t a b l e   t o   A p p l e   C a l e n d a r .   M u s t   o p e n   t h e   e x c e l   t i m e t a b l e   i n   N u m b e r s   f i r s t ,   t h e n   r u n   t h i s   s c r i p t $  ( ) ( l     ��������  ��  ��   )  * + * l     ,���� , I    �� -��
�� .sysottosnull���     TEXT - m      . . � / / � T h i s   s c r i p t   w i l l   f i r s t   d e l e t e   a n y   c a l e n d a r s   t h a t   c o n t a i n   t h e   w o r d   ' Y e a r '   i n   t h e   c a l e n d a r   n a m e .��  ��  ��   +  0 1 0 l    2���� 2 I   �� 3��
�� .sysodlogaskr        TEXT 3 m     4 4 � 5 5 � T h i s   s c r i p t   w i l l   f i r s t   d e l e t e   a n y   c a l e n d a r s   t h a t   c o n t a i n   t h e   w o r d   ' Y e a r '   i n   t h e   c a l e n d a r   n a m e .��  ��  ��   1  6 7 6 l     ��������  ��  ��   7  8 9 8 l    :���� : I   �� ;��
�� .sysottosnull���     TEXT ; m     < < � = = � W h i c h   c a l e n d a r   a p p   d o   y o u   w a n t   t o   c r e a t e   t h e   e v e n t s   i n ?   A p p l e   C a l e n d a r   o r   M i c r o s o f t   O u t l o o k ?��  ��  ��   9  > ? > l    @���� @ r     A B A I   �� C D
�� .gtqpchltns    @   @ ns   C J     E E  F G F m     H H � I I  A p p l e   C a l e n d a r G  J�� J m     K K � L L " M i c r o s o f t   O u t l o o k��   D �� M��
�� 
prmp M m     N N � O O p W h i c h   c a l e n d a r   a p p   d o   y o u   w a n t   t h e   c a l e n d a r s   c r e a t e d   i n ?��   B o      ���� &0 chosencalendarapp chosenCalendarApp��  ��   ?  P Q P l   % R���� R r    % S T S n    # U V U 4     #�� W
�� 
cobj W m   ! "����  V o     ���� &0 chosencalendarapp chosenCalendarApp T o      ���� &0 chosencalendarapp chosenCalendarApp��  ��   Q  X Y X l     ��������  ��  ��   Y  Z [ Z l  & + \���� \ I  & +�� ]��
�� .sysottosnull���     TEXT ] m   & ' ^ ^ � _ _ � W h e n   i s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r ?   ( S h o u l d   b e   o n   a   M o n d a y ,   y o u   c a n   e n t e r   a   d e l a y e d   s t a r t   d a y   l a t e r )��  ��  ��   [  ` a ` l  , � b���� b T   , � c c k   1 � d d  e f e r   1 8 g h g I  1 6������
�� .misccurdldt    ��� null��  ��   h o      ����  0 thecurrentdate theCurrentDate f  i j i I  9 P�� k l
�� .sysodlogaskr        TEXT k m   9 : m m � n n � W h e n   i s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r ?   ( S h o u l d   b e   o n   a   M o n d a y ,   y o u   c a n   e n t e r   a   d e l a y e d   s t a r t   d a y   l a t e r ) l �� o��
�� 
dtxt o l  = L p���� p b   = L q r q b   = F s t s n   = B u v u 1   > B��
�� 
dstr v o   = >����  0 thecurrentdate theCurrentDate t 1   B E��
�� 
spac r n   F K w x w 1   G K��
�� 
tstr x o   F G����  0 thecurrentdate theCurrentDate��  ��  ��   j  y z y r   Q \ { | { n   Q X } ~ } 1   T X��
�� 
ttxt ~ 1   Q T��
�� 
rslt | o      ���� 0 thetext theText z  ��  Q   ] � � � � � Z   ` | � ����� � >  ` g � � � o   ` c���� 0 thetext theText � m   c f � � � � �   � k   j x � �  � � � l  j v � � � � r   j v � � � 4   j r�� �
�� 
ldt  � o   n q���� 0 thetext theText � o      ���� 0 firstday firstDay �   a date object    � � � �    a   d a t e   o b j e c t �  ��� �  S   w x��  ��  ��   � R      ������
�� .ascrerr ****      � ****��  ��   � I  � �������
�� .sysobeepnull��� ��� long��  ��  ��  ��  ��   a  � � � l     ��������  ��  ��   �  � � � l  � � ����� � I  � ��� ���
�� .sysottosnull���     TEXT � b   � � � � � b   � � � � � m   � � � � � � � D T h e   f i r s t   d a y   o f   t h e   s e m e s t e r   i s :   � o   � ����� 0 firstday firstDay � m   � � � � � � �  P l e a s e   C o n f i r m��  ��  ��   �  � � � l  � � ����� � I  � ��� � �
�� .sysodlogaskr        TEXT � m   � � � � � � � D T h e   f i r s t   d a y   o f   t h e   s e m e s t e r   i s :   � �� ���
�� 
dtxt � l  � � ����� � b   � � � � � b   � � � � � n   � � � � � 1   � ���
�� 
dstr � o   � ����� 0 firstday firstDay � 1   � ���
�� 
spac � n   � � � � � 1   � ���
�� 
tstr � o   � ����� 0 firstday firstDay��  ��  ��  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l  � � ����� � I  � ��� ���
�� .sysottosnull���     TEXT � m   � � � � � � � J W h e n   i s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r ?��  ��  ��   �  � � � l  �) ����� � T   �) � � k   �$ � �  � � � r   � � � � � I  � �������
�� .misccurdldt    ��� null��  ��   � o      ����  0 thecurrentdate theCurrentDate �  � � � I  � ��� � �
�� .sysodlogaskr        TEXT � m   � � � � � � � J W h e n   i s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r ? � �� ���
�� 
dtxt � l  � � ����� � b   � � � � � b   � � � � � n   � � � � � 1   � ���
�� 
dstr � o   � �����  0 thecurrentdate theCurrentDate � 1   � ���
�� 
spac � n   � � � � � 1   � ��
� 
tstr � o   � ��~�~  0 thecurrentdate theCurrentDate��  ��  ��   �  � � � r   � � � � � n   � � � � � 1   � ��}
�} 
ttxt � 1   � ��|
�| 
rslt � o      �{�{ 0 thetext theText �  ��z � Q   �$ � � � � Z   � � ��y�x � >  � � � � o   � ��w�w 0 thetext theText � m   � � � � � �   � k   � �  � � � l  � � � � r   � � � 4  �v �
�v 
ldt  � o  	�u�u 0 thetext theText � o      �t�t 0 thedate theDate �   a date object    � � � �    a   d a t e   o b j e c t �  ��s �  S  �s  �y  �x   � R      �r�q�p
�r .ascrerr ****      � ****�q  �p   � I $�o�n�m
�o .sysobeepnull��� ��� long�n  �m  �z  ��  ��   �  � � � l     �l�k�j�l  �k  �j   �  � � � l *9 ��i�h � I *9�g ��f
�g .sysottosnull���     TEXT � b  *5 � � � b  *1 � � � m  *- � � � � � B T h e   l a s t   d a y   o f   t h e   s e m e s t e r   i s :   � o  -0�e�e 0 thedate theDate � m  14 � � � � �  P l e a s e   C o n f i r m .�f  �i  �h   �  � � � l :W ��d�c � I :W�b 
�b .sysodlogaskr        TEXT  m  := � B T h e   l a s t   d a y   o f   t h e   s e m e s t e r   i s :   �a�`
�a 
dtxt l @S�_�^ b  @S b  @K	 n  @G

 1  CG�]
�] 
dstr o  @C�\�\ 0 thedate theDate	 1  GJ�[
�[ 
spac n  KR 1  NR�Z
�Z 
tstr o  KN�Y�Y 0 thedate theDate�_  �^  �`  �d  �c   �  l     �X�W�V�X  �W  �V    l     �U�T�S�U  �T  �S    l     �R�R     Ask for delayed start    � ,   A s k   f o r   d e l a y e d   s t a r t  l X_�Q�P I X_�O�N
�O .sysottosnull���     TEXT m  X[ � R I s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r   d e l a y e d ?�N  �Q  �P    l `t�M�L r  `t !  I `p�K"#
�K .gtqpchltns    @   @ ns  " J  `h$$ %&% m  `c'' �((  Y e s& )�J) m  cf** �++  N o�J  # �I,�H
�I 
prmp, m  il-- �.. R I s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r   d e l a y e d ?�H  ! o      �G�G  0 delayedyesorno delayedYesOrNo�M  �L   /0/ l u1�F�E1 r  u232 n  u{454 4  x{�D6
�D 
cobj6 m  yz�C�C 5 o  ux�B�B  0 delayedyesorno delayedYesOrNo3 o      �A�A  0 delayedyesorno delayedYesOrNo�F  �E  0 787 l     �@�?�>�@  �?  �>  8 9:9 l ��;�=�<; Z  ��<=�;�:< = ��>?> o  ���9�9  0 delayedyesorno delayedYesOrNo? m  ��@@ �AA  Y e s= k  ��BB CDC I ���8E�7
�8 .sysottosnull���     TEXTE m  ��FF �GG � I f   t h e   f i r s t   d a y   i s   d e l a y e d   a n d   n o t   o n   a   M o n d a y ,   w h e n   i s   t h e   d e l a y e d   f i r s t   d a y ?�7  D H�6H T  ��II k  ��JJ KLK r  ��MNM I ���5�4�3
�5 .misccurdldt    ��� null�4  �3  N o      �2�2  0 thecurrentdate theCurrentDateL OPO I ���1QR
�1 .sysodlogaskr        TEXTQ m  ��SS �TT � I f   t h e   f i r s t   d a y   i s   d e l a y e d   a n d   n o t   o n   a   M o n d a y ,   w h e n   i s   t h e   d e l a y e d   f i r s t   d a y ?R �0U�/
�0 
dtxtU l ��V�.�-V b  ��WXW b  ��YZY n  ��[\[ 1  ���,
�, 
dstr\ o  ���+�+  0 thecurrentdate theCurrentDateZ 1  ���*
�* 
spacX n  ��]^] 1  ���)
�) 
tstr^ o  ���(�(  0 thecurrentdate theCurrentDate�.  �-  �/  P _`_ r  ��aba n  ��cdc 1  ���'
�' 
ttxtd 1  ���&
�& 
rsltb o      �%�% 0 thetext theText` e�$e Q  ��fghf Z  ��ij�#�"i > ��klk o  ���!�! 0 thetext theTextl m  ��mm �nn  j k  ��oo pqp l ��rstr r  ��uvu 4  ��� w
�  
ldt w o  ���� 0 thetext theTextv o      �� "0 delayedfirstday delayedFirstDays   a date object   t �xx    a   d a t e   o b j e c tq y�y  S  ���  �#  �"  g R      ���
� .ascrerr ****      � ****�  �  h I �����
� .sysobeepnull��� ��� long�  �  �$  �6  �;  �:  �=  �<  : z{z l     ����  �  �  { |}| l �~��~ r  �� [  ���� o  ���� 0 thedate theDate� l ����� ]  ���� 1  ��
� 
days� m  �� �  �  � o      �� 0 thedate theDate�  �  } ��� l     ��
�	�  �
  �	  � ��� l 	���� r  	��� l 	���� n  	��� 1  �
� 
year� o  	�� 0 thedate theDate�  �  � o      �� 0 a  �  �  � ��� l $��� � r  $��� c   ��� l ������ n  ��� m  ��
�� 
mnth� o  ���� 0 thedate theDate��  ��  � m  ��
�� 
nmbr� o      ���� 0 b  �  �   � ��� l %0������ r  %0��� n  %,��� 1  (,��
�� 
day � o  %(���� 0 thedate theDate� o      ���� 0 c  ��  ��  � ��� l     ������  �  set d to "T000000Z"   � ��� & s e t   d   t o   " T 0 0 0 0 0 0 Z "� ��� l 18������ r  18��� m  14�� ��� H F R E Q = W E E K L Y ; I N T E R V A L = 1 ; B Y D A Y = ; U N T I L =� o      ���� 0 e  ��  ��  � ��� l     ��������  ��  ��  � ��� l 9@���� r  9@��� o  9<���� 0 thedate theDate� o      ���� 0 enddate endDate� R L this variable represents the last day of the semester for Microsoft Outlook   � ��� �   t h i s   v a r i a b l e   r e p r e s e n t s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r   f o r   M i c r o s o f t   O u t l o o k� ��� l     ��������  ��  ��  � ��� l AP������ r  AP��� J  AL�� ��� o  AD���� 0 a  � ��� o  DG���� 0 b  � ���� o  GJ���� 0 c  ��  � o      ���� 0 thelist theList��  ��  � ��� l     ��������  ��  ��  � ��� l Qp������ r  Qp��� J  Q[�� ��� 1  QV��
�� 
txdl� ���� m  VY�� ���  0��  � J      �� ��� o      ���� 
0 tid TID� ���� 1  in��
�� 
txdl��  ��  ��  � ��� l q|������ r  q|��� c  qx��� o  qt���� 0 thelist theList� m  tw��
�� 
ctxt� o      ���� "0 thelistasstring theListAsString��  ��  � ��� l }������� r  }���� o  }����� 
0 tid TID� 1  ����
�� 
txdl��  ��  � ��� l     ��������  ��  ��  � ��� l �������� r  ����� J  ���� ��� o  ������ 0 e  � ���� o  ������ "0 thelistasstring theListAsString��  � o      ���� 0 thelist theList��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  � Y S the variable endOfRecurrence represents the last day of classes for Apple Calendar   � ��� �   t h e   v a r i a b l e   e n d O f R e c u r r e n c e   r e p r e s e n t s   t h e   l a s t   d a y   o f   c l a s s e s   f o r   A p p l e   C a l e n d a r� ��� l �������� r  ����� J  ���� ��� 1  ����
�� 
txdl� ���� m  ���� ���  ��  � J      �� � � o      ���� 
0 tid TID  �� 1  ����
�� 
txdl��  ��  ��  �  l ������ r  �� c  �� o  ������ 0 thelist theList m  ����
�� 
ctxt o      ���� "0 endofrecurrence endOfRecurrence��  ��   	
	 l ������ r  �� o  ������ 
0 tid TID 1  ����
�� 
txdl��  ��  
  l     ��������  ��  ��    l ������ r  �� J  ������   o      ���� 0 
classnames 
classNames��  ��    l �}���� O  �} k  �|  I ��������
�� .miscactvnull��� ��� null��  ��    r  ��  n  ��!"! 1  ����
�� 
pnam" 2 ����
�� 
docu  o      ���� "0 listofdocuments listOfDocuments #$# O ��%&% I ����'��
�� .sysottosnull���     TEXT' m  ��(( �)) � C h o o s e   w h i c h   N u m b e r s   d o c u m e n t   h a s   t h e   t i m e t a b l e   a n d   c l a s s   n a m e s .��  &  f  ��$ *+* r  �,-, I ����.��
�� .gtqpchltns    @   @ ns  . o  ������ "0 listofdocuments listOfDocuments��  - o      ����  0 chosendocument chosenDocument+ /0/ r  121 n  
343 4  
��5
�� 
cobj5 m  	���� 4 o  ����  0 chosendocument chosenDocument2 o      ����  0 chosendocument chosenDocument0 676 l ��������  ��  ��  7 8��8 O  |9:9 k  {;; <=< O &>?> I %��@��
�� .sysottosnull���     TEXT@ m  !AA �BB B C h o o s e   t h e   s h e e t   n a m e d   " C l a s s e s " .��  ?  f  = CDC r  '4EFE n  '0GHG 1  ,0��
�� 
pnamH 2 ',��
�� 
NmShF o      ���� 0 listofsheets listOfSheetsD IJI r  5@KLK I 5<��M��
�� .gtqpchltns    @   @ ns  M o  58���� 0 listofsheets listOfSheets��  L o      ���� 0 chosensheet chosenSheetJ NON r  AKPQP n  AGRSR 4  DG��T
�� 
cobjT m  EF���� S o  AD���� 0 chosensheet chosenSheetQ o      ���� 0 chosensheet chosenSheetO UVU O  L�WXW k  W�YY Z[Z O  W�\]\ k  `�^^ _`_ l ``��ab��  a   get year one class names   b �cc 2   g e t   y e a r   o n e   c l a s s   n a m e s` ded r  `yfgf 6 `uhih 4 `f��j
�� 
NmClj m  de���� i E  itklk n  jnmnm 1  jn��
�� 
NMCvn  g  jjl m  osoo �pp 
 C l a s sg o      ���� 0 yearonebegin yearOneBegine qrq r  z�sts n  z�uvu 1  ����
�� 
NMadv n  z�wxw m  }���
�� 
NMCox o  z}���� 0 yearonebegin yearOneBegint o      �� ,0 yearonecolumnaddress yearOneColumnAddressr yzy r  ��{|{ I ���~}�}
�~ .corecnte****       ****} n  ��~~ 2 ���|
�| 
NmCl 4  ���{�
�{ 
NMCo� o  ���z�z ,0 yearonecolumnaddress yearOneColumnAddress�}  | o      �y�y (0 yearonecolumncount yearOneColumnCountz ��� r  ����� [  ����� l ����x�w� n  ����� 1  ���v
�v 
NMad� n  ����� m  ���u
�u 
NMRw� o  ���t�t 0 yearonebegin yearOneBegin�x  �w  � m  ���s�s � o      �r�r 0 yearonebegin yearOneBegin� ��� l ���q�p�o�q  �p  �o  � ��� Y  ���n���m� k  ��� ��� Z  ����l�� = ����� n  ����� 1  ���k
�k 
NMCv� n  ����� 4  ���j�
�j 
NmCl� o  ���i�i 0 i  � 4  ���h�
�h 
NMCo� o  ���g�g ,0 yearonecolumnaddress yearOneColumnAddress� m  ���f
�f 
msng�  S  ���l  � k  ��� ��� l ���e���e  � � �						set end of classNames to "Year 1 " & (value of cell i of column yearOneColumnAddress & " " & value of cell i of column (yearOneColumnAddress + 4))   � ���0 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e s   t o   " Y e a r   1   "   &   ( v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s   &   "   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   ( y e a r O n e C o l u m n A d d r e s s   +   4 ) )� ��d� r  ���� b  ���� b  ����� n  ����� 1  ���c
�c 
NMCv� n  ����� 4  ���b�
�b 
NmCl� o  ���a�a 0 i  � 4  ���`�
�` 
NMCo� l ����_�^� [  ����� o  ���]�] ,0 yearonecolumnaddress yearOneColumnAddress� m  ���\�\ �_  �^  � m  ���� ���    Y e a r   1  � l ���[�Z� n  ���� 1  �Y
�Y 
NMCv� n  ���� 4  ��X�
�X 
NmCl� o  � �W�W 0 i  � 4  ���V�
�V 
NMCo� o  ���U�U ,0 yearonecolumnaddress yearOneColumnAddress�[  �Z  � n      ���  ;  	
� o  	�T�T 0 
classnames 
classNames�d  � ��S� l �R���R  � 9 3					value of cell i of column yearOneColumnAddress   � ��� f 	 	 	 	 	 v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s�S  �n 0 i  � o  ���Q�Q 0 yearonebegin yearOneBegin� o  ���P�P (0 yearonecolumncount yearOneColumnCount�m  � ��� l �O�N�M�O  �N  �M  � ��� l �L���L  �   get year two class names   � ��� 2   g e t   y e a r   t w o   c l a s s   n a m e s� ��� r  ,��� 6 (��� 4 �K�
�K 
NmCl� m  �J�J � E  '��� n  !��� 1  !�I
�I 
NMCv�  g  � m  "&�� ��� 
 C l a s s� o      �H�H 0 yearonebegin yearOneBegin� ��� r  -<��� n  -8��� 1  48�G
�G 
NMad� n  -4��� m  04�F
�F 
NMCo� o  -0�E�E 0 yearonebegin yearOneBegin� o      �D�D ,0 yearonecolumnaddress yearOneColumnAddress� ��� r  =Q��� I =M�C��B
�C .corecnte****       ****� n  =I��� 2 EI�A
�A 
NmCl� 4  =E�@�
�@ 
NMCo� o  AD�?�? ,0 yearonecolumnaddress yearOneColumnAddress�B  � o      �>�> (0 yearonecolumncount yearOneColumnCount� ��� r  Rc��� [  R_��� l R]��=�<� n  R]��� 1  Y]�;
�; 
NMad� n  RY��� m  UY�:
�: 
NMRw� o  RU�9�9 0 yearonebegin yearOneBegin�=  �<  � m  ]^�8�8 � o      �7�7 0 yearonebegin yearOneBegin� ��� l dd�6�5�4�6  �5  �4  � ��� Y  d���3���2� k  r��� ��� Z  r�� �1� = r� n  r� 1  ��0
�0 
NMCv n  r 4  z�/
�/ 
NmCl o  }~�.�. 0 i   4  rz�-	
�- 
NMCo	 o  vy�,�, ,0 yearonecolumnaddress yearOneColumnAddress m  ���+
�+ 
msng   S  ���1   k  ��

  l ���*�*   � �						set end of classNames to "Year 2 " & (value of cell i of column yearOneColumnAddress & " " & value of cell i of column (yearOneColumnAddress + 4))    �0 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e s   t o   " Y e a r   2   "   &   ( v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s   &   "   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   ( y e a r O n e C o l u m n A d d r e s s   +   4 ) ) �) r  �� b  �� b  �� n  �� 1  ���(
�( 
NMCv n  �� 4  ���'
�' 
NmCl o  ���&�& 0 i   4  ���%
�% 
NMCo l ���$�# [  �� o  ���"�" ,0 yearonecolumnaddress yearOneColumnAddress m  ���!�! �$  �#   m  ��   �!!    Y e a r   2   l ��"� �" n  ��#$# 1  ���
� 
NMCv$ n  ��%&% 4  ���'
� 
NmCl' o  ���� 0 i  & 4  ���(
� 
NMCo( o  ���� ,0 yearonecolumnaddress yearOneColumnAddress�   �   n      )*)  ;  ��* o  ���� 0 
classnames 
classNames�)  � +�+ l ���,-�  , 9 3					value of cell i of column yearOneColumnAddress   - �.. f 	 	 	 	 	 v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s�  �3 0 i  � o  gj�� 0 yearonebegin yearOneBegin� o  jm�� (0 yearonecolumncount yearOneColumnCount�2  � /0/ l ������  �  �  0 121 r  ��343 n ��565 I  ���7�� 0 simple_sort  7 8�8 l ��9��9 o  ���� 0 
classnames 
classNames�  �  �  �  6  f  ��4 o      �� 0 
sortedlist 
sortedList2 :;: r  ��<=< o  ���
�
 0 
sortedlist 
sortedList= o      �	�	 0 
classnames 
classNames; >�> l ������  �  �  �  ] 4  W]�?
� 
NmTb? m  [\�� [ @�@ l ���� ���  �   ��  �  X 4  LT��A
�� 
NmShA o  PS���� 0 chosensheet chosenSheetV BCB O ��DED I ����F��
�� .sysottosnull���     TEXTF m  ��GG �HH F C h o o s e   t h e   s h e e t   n a m e d   " T i m e t a b l e " .��  E  f  ��C IJI r  ��KLK n  ��MNM 1  ����
�� 
pnamN 2 ����
�� 
NmShL o      ���� 0 listofsheets listOfSheetsJ OPO r  �QRQ I ���S��
�� .gtqpchltns    @   @ ns  S o  ������ 0 listofsheets listOfSheets��  R o      ���� 0 chosensheet chosenSheetP TUT r  VWV n  XYX 4  ��Z
�� 
cobjZ m  ���� Y o  ���� 0 chosensheet chosenSheetW o      ���� 0 chosensheet chosenSheetU [\[ l ��������  ��  ��  \ ]^] O _`_ I ��a��
�� .sysottosnull���     TEXTa m  bb �cc V G r a b b i n g   d a t a   f r o m   t h e   n u m b e r s   s p r e a d s h e e t .��  `  f  ^ d��d O   {efe O  +zghg k  4yii jkj l 44��lm��  l E ? Find table of room numbers and teachers and put it into a list   m �nn ~   F i n d   t a b l e   o f   r o o m   n u m b e r s   a n d   t e a c h e r s   a n d   p u t   i t   i n t o   a   l i s tk opo r  4Mqrq 6 4Ists 4 4:��u
�� 
NmClu m  89���� t = =Hvwv 1  >B��
�� 
NMCvw m  CGxx �yy  R o o mr o      ���� 0 	roomindex 	roomIndexp z{z r  N]|}| n  NY~~ 1  UY��
�� 
NMad n  NU��� m  QU��
�� 
NMCo� o  NQ���� 0 	roomindex 	roomIndex} o      ���� 0 columnnumber columnNumber{ ��� r  ^m��� n  ^i��� 1  ei��
�� 
NMad� n  ^e��� m  ae��
�� 
NMRw� o  ^a���� 0 	roomindex 	roomIndex� o      ���� 0 	rownumber 	rowNumber� ��� r  n���� I n~�����
�� .corecnte****       ****� n nz��� 2 vz��
�� 
NmCl� 4  nv���
�� 
NMCo� o  ru���� 0 columnnumber columnNumber��  � o      ���� 0 rowcount rowCount� ��� r  ����� J  ������  � o      ���� 0 roomnumbers roomNumbers� ��� Y  ���������� k  ���� ��� Z  ��������� > ����� n  ����� 1  ����
�� 
NMCv� n  ����� 4  �����
�� 
NmCl� o  ������ 0 i  � 4  �����
�� 
NMCo� l �������� \  ����� o  ������ 0 columnnumber columnNumber� m  ������ ��  ��  � m  ����
�� 
msng� r  ����� J  ���� ��� l �������� n  ����� 1  ����
�� 
NMCv� n  ����� 4  �����
�� 
NmCl� o  ������ 0 i  � 4  �����
�� 
NMCo� l �������� \  ����� o  ������ 0 columnnumber columnNumber� m  ������ ��  ��  ��  ��  � ���� c  ����� n  ����� 1  ����
�� 
NMCv� n  ����� 4  �����
�� 
NmCl� o  ������ 0 i  � 4  �����
�� 
NMCo� o  ������ 0 columnnumber columnNumber� m  ����
�� 
ctxt��  � n      ���  ;  ��� o  ������ 0 roomnumbers roomNumbers��  ��  � ���� l ����������  ��  ��  ��  �� 0 i  � [  ����� o  ������ 0 	rownumber 	rowNumber� m  ������ � o  ������ 0 rowcount rowCount��  � ��� l  ��������  ���
				# Create name for all calendars Year X / Class / Teacher Name
				
				set columnNumber to address of column of cell 1 where value of it contains "New Teacher"
				set rowNumber to address of row of cell 1 where value of it contains "New Teacher"
				set className to {}
				
				# Name Year 1 classes
				repeat with i from rowNumber + 1 to count of rows of column columnNumber
					if value of cell i of column columnNumber is not missing value then
						set end of className to "Bear 1 " & value of cell i of column (columnNumber - 2) & " " & value of cell i of column columnNumber
					else
						exit repeat
					end if
				end repeat
				
				set columnNumber to address of column of cell 2 where value of it contains "New Teacher"
				set rowNumber to address of row of cell 2 where value of it contains "New Teacher"
				
				# Name Year 2 classes
				repeat with i from rowNumber + 1 to count of rows of column columnNumber
					if value of cell i of column columnNumber is not missing value then
						set end of className to "Bear 2 " & value of cell i of column (columnNumber - 2) & " " & value of cell i of column columnNumber
					else
						exit repeat
					end if
				end repeat
				   � ���	j 
 	 	 	 	 #   C r e a t e   n a m e   f o r   a l l   c a l e n d a r s   Y e a r   X   /   C l a s s   /   T e a c h e r   N a m e 
 	 	 	 	 
 	 	 	 	 s e t   c o l u m n N u m b e r   t o   a d d r e s s   o f   c o l u m n   o f   c e l l   1   w h e r e   v a l u e   o f   i t   c o n t a i n s   " N e w   T e a c h e r " 
 	 	 	 	 s e t   r o w N u m b e r   t o   a d d r e s s   o f   r o w   o f   c e l l   1   w h e r e   v a l u e   o f   i t   c o n t a i n s   " N e w   T e a c h e r " 
 	 	 	 	 s e t   c l a s s N a m e   t o   { } 
 	 	 	 	 
 	 	 	 	 #   N a m e   Y e a r   1   c l a s s e s 
 	 	 	 	 r e p e a t   w i t h   i   f r o m   r o w N u m b e r   +   1   t o   c o u n t   o f   r o w s   o f   c o l u m n   c o l u m n N u m b e r 
 	 	 	 	 	 i f   v a l u e   o f   c e l l   i   o f   c o l u m n   c o l u m n N u m b e r   i s   n o t   m i s s i n g   v a l u e   t h e n 
 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e   t o   " B e a r   1   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   ( c o l u m n N u m b e r   -   2 )   &   "   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   c o l u m n N u m b e r 
 	 	 	 	 	 e l s e 
 	 	 	 	 	 	 e x i t   r e p e a t 
 	 	 	 	 	 e n d   i f 
 	 	 	 	 e n d   r e p e a t 
 	 	 	 	 
 	 	 	 	 s e t   c o l u m n N u m b e r   t o   a d d r e s s   o f   c o l u m n   o f   c e l l   2   w h e r e   v a l u e   o f   i t   c o n t a i n s   " N e w   T e a c h e r " 
 	 	 	 	 s e t   r o w N u m b e r   t o   a d d r e s s   o f   r o w   o f   c e l l   2   w h e r e   v a l u e   o f   i t   c o n t a i n s   " N e w   T e a c h e r " 
 	 	 	 	 
 	 	 	 	 #   N a m e   Y e a r   2   c l a s s e s 
 	 	 	 	 r e p e a t   w i t h   i   f r o m   r o w N u m b e r   +   1   t o   c o u n t   o f   r o w s   o f   c o l u m n   c o l u m n N u m b e r 
 	 	 	 	 	 i f   v a l u e   o f   c e l l   i   o f   c o l u m n   c o l u m n N u m b e r   i s   n o t   m i s s i n g   v a l u e   t h e n 
 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e   t o   " B e a r   2   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   ( c o l u m n N u m b e r   -   2 )   &   "   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   c o l u m n N u m b e r 
 	 	 	 	 	 e l s e 
 	 	 	 	 	 	 e x i t   r e p e a t 
 	 	 	 	 	 e n d   i f 
 	 	 	 	 e n d   r e p e a t 
 	 	 	 	� ��� l ��������  � !  Grab data from spreadsheet   � ��� 6   G r a b   d a t a   f r o m   s p r e a d s h e e t� ��� r  ����� J  ������  � o      ���� 0 lessons  � ���� Y  �y�������� Y  t�������� Z  o������� > #��� n  ��� 1  ��
�� 
NMCv� n  ��� 4  ���
�� 
NmCl� o  ���� 0 j  � 4  ���
�� 
NMCo� o  ���� 0 i  � m  "��
�� 
msng� r  &k��� J  &f�� ��� n  &5��� 1  15��
�� 
NMCv� n  &1��� 4  ,1���
�� 
NmCl� m  /0���� � 4  &,���
�� 
NMCo� o  *+���� 0 i  � ��� n  5D��� 1  @D��
�� 
NMCv� n  5@��� 4  ;@���
�� 
NmCl� m  >?���� � 4  5;���
�� 
NMCo� o  9:���� 0 i  � ��� n  DS��� 1  OS��
�� 
NMCv� n  DO��� 4  JO�� 
�� 
NmCl  o  MN���� 0 j  � 4  DJ��
�� 
NMCo m  HI���� � �� n  Sb 1  ^b��
�� 
NMCv n  S^ 4  Y^��
�� 
NmCl o  \]���� 0 j   4  SY��
�� 
NMCo o  WX���� 0 i  ��  � n      	
	  ;  ij
 o  fi���� 0 lessons  ��  ��  �� 0 j  � m  ���� � m  ���� "��  �� 0 i  � m  ������ � m  ������ ��  ��  h 4  +1��
�� 
NmTb m  /0���� f 4   (��
�� 
NmSh o  $'���� 0 chosensheet chosenSheet��  : 4  �
� 
docu o  �~�~  0 chosendocument chosenDocument��   m  ���                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  ��  ��    l     �}�|�{�}  �|  �{    l     �z�z   7 1 convert year 1 or year 2 from chinese to english    � b   c o n v e r t   y e a r   1   o r   y e a r   2   f r o m   c h i n e s e   t o   e n g l i s h  l ~��y�x r  ~� J  ~��w�w   o      �v�v 0 lessons3  �y  �x    l ���u�t X  ���s Z  �� !�r"  E  ��#$# n  ��%&% 4  ���q'
�q 
cobj' m  ���p�p & o  ���o�o 
0 lesson  $ m  ��(( �)) Y'N ! r  ��*+* J  ��,, -.- n  ��/0/ 4  ���n1
�n 
cobj1 m  ���m�m 0 o  ���l�l 
0 lesson  . 232 m  ��44 �55  Y e a r   13 676 n  ��898 4  ���k:
�k 
cobj: m  ���j�j 9 o  ���i�i 
0 lesson  7 ;�h; n  ��<=< 4  ���g>
�g 
cobj> m  ���f�f = o  ���e�e 
0 lesson  �h  + n      ?@?  ;  ��@ o  ���d�d 0 lessons3  �r  " r  ��ABA J  ��CC DED n  ��FGF 4  ���cH
�c 
cobjH m  ���b�b G o  ���a�a 
0 lesson  E IJI m  ��KK �LL  Y e a r   2J MNM n  ��OPO 4  ���`Q
�` 
cobjQ m  ���_�_ P o  ���^�^ 
0 lesson  N R�]R n  ��STS 4  ���\U
�\ 
cobjU m  ���[�[ T o  ���Z�Z 
0 lesson  �]  B n      VWV  ;  ��W o  ���Y�Y 0 lessons3  �s 
0 lesson   o  ���X�X 0 lessons  �u  �t   XYX l     �W�V�U�W  �V  �U  Y Z[Z l     �T\]�T  \   add time to each event   ] �^^ .   a d d   t i m e   t o   e a c h   e v e n t[ _`_ l ��a�S�Ra r  ��bcb J  ���Q�Q  c o      �P�P 0 lessons4  �S  �R  ` ded l ��f�O�Nf X  ��g�Mhg k  ��ii jkj r  �lml n  � non 4  � �Lp
�L 
cobjp m  ���K�K o o  ���J�J 
0 lesson  m o      �I�I 0 	inputtime 	inputTimek qrq r  "sts b  uvu b  wxw n  yzy 4  �H{
�H 
cwor{ m  �G�G z o  �F�F 0 	inputtime 	inputTimex m  || �}}  :v n  ~~ 4  �E�
�E 
cwor� m  �D�D  o  �C�C 0 	inputtime 	inputTimet o      �B�B 0 thetime theTimer ��� Z  #���A�� ?  #3��� c  #/��� n  #+��� 4  &+�@�
�@ 
cwor� m  )*�?�? � l #&��>�=� o  #&�<�< 0 thetime theTime�>  �=  � m  +.�;
�; 
nmbr� m  /2�:�: � r  6[��� c  6W��� b  6S��� b  6O��� b  6F��� \  6B��� l 6>��9�8� n  6>��� 4  9>�7�
�7 
cwor� m  <=�6�6 � o  69�5�5 0 thetime theTime�9  �8  � m  >A�4�4 � m  BE�� ���  :� n  FN��� 4  IN�3�
�3 
cwor� m  LM�2�2 � o  FI�1�1 0 thetime theTime� m  OR�� ���  P M� m  SV�0
�0 
ctxt� o      �/�/ 0 thetime2 theTime2�A  � r  ^��� c  ^{��� b  ^w��� b  ^s��� b  ^j��� n  ^f��� 4  af�.�
�. 
cwor� m  de�-�- � o  ^a�,�, 0 thetime theTime� m  fi�� ���  :� n  jr��� 4  mr�+�
�+ 
cwor� m  pq�*�* � o  jm�)�) 0 thetime theTime� m  sv�� ���  A M� m  wz�(
�( 
ctxt� o      �'�' 0 thetime2 theTime2� ��&� r  ����� J  ���� ��� n  ����� 4  ���%�
�% 
cobj� m  ���$�$ � o  ���#�# 
0 lesson  � ��� n  ����� 4  ���"�
�" 
cobj� m  ���!�! � o  ��� �  
0 lesson  � ��� o  ���� 0 thetime2 theTime2� ��� n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  �  � n      ���  ;  ��� o  ���� 0 lessons4  �&  �M 
0 lesson  h o  ���� 0 lessons3  �O  �N  e ��� l     ����  �  �  � ��� l     ����  � , & Combine Year, class, and teacher name   � ��� L   C o m b i n e   Y e a r ,   c l a s s ,   a n d   t e a c h e r   n a m e� ��� l ������ r  ����� J  ����  � o      �� 0 lessons5  �  �  � ��� l ������ X  ������ r  ����� J  ���� ��� b  ����� b  ����� n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  � m  ���� ���   � n  ����� 4  ���
�
�
 
cobj� m  ���	�	 � o  ���� 
0 lesson  � ��� n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  � ��� n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  �  � n      ���  ;  ��� o  ��� �  0 lessons5  � 
0 lesson  � o  ������ 0 lessons4  �  �  � ��� l     ��������  ��  ��  � ��� l     �� ��    ( " add date to each item in the list    � D   a d d   d a t e   t o   e a c h   i t e m   i n   t h e   l i s t�  l ������ r  �� J  ������   o      ���� 0 lessons6  ��  ��   	 l ��
����
 X  ���� Z  ���� E  � n  �� 4  ����
�� 
cobj m  ������  o  ������ 
0 lesson   m  � �  M o n d a y r    J    n  	 4  	��
�� 
cobj m  ����  o  ���� 
0 lesson   �� b  	 !  b  	"#" n  	$%$ 1  ��
�� 
dstr% o  	���� 0 firstday firstDay# m  && �''    a t  ! n  ()( 4  ��*
�� 
cobj* m  ���� ) o  ���� 
0 lesson  ��   n      +,+  ;  , o  ���� 0 lessons6   -.- E  #+/0/ n  #'121 4  $'��3
�� 
cobj3 m  %&���� 2 o  #$���� 
0 lesson  0 m  '*44 �55  T u e s d a y. 676 r  .O898 J  .J:: ;<; n  .2=>= 4  /2��?
�� 
cobj? m  01���� > o  ./���� 
0 lesson  < @��@ b  2HABA b  2CCDC n  2?EFE 1  ;?��
�� 
dstrF l 2;G����G [  2;HIH o  25���� 0 firstday firstDayI l 5:J����J ]  5:KLK 1  58��
�� 
daysL m  89���� ��  ��  ��  ��  D m  ?BMM �NN    a t  B n  CGOPO 4  DG��Q
�� 
cobjQ m  EF���� P o  CD���� 
0 lesson  ��  9 n      RSR  ;  MNS o  JM���� 0 lessons6  7 TUT E  RZVWV n  RVXYX 4  SV��Z
�� 
cobjZ m  TU���� Y o  RS���� 
0 lesson  W m  VY[[ �\\  W e d n e s d a yU ]^] r  ]~_`_ J  ]yaa bcb n  ]aded 4  ^a��f
�� 
cobjf m  _`���� e o  ]^���� 
0 lesson  c g��g b  awhih b  arjkj n  anlml 1  jn��
�� 
dstrm l ajn����n [  ajopo o  ad���� 0 firstday firstDayp l diq����q ]  dirsr 1  dg��
�� 
dayss m  gh���� ��  ��  ��  ��  k m  nqtt �uu    a t  i n  rvvwv 4  sv��x
�� 
cobjx m  tu���� w o  rs���� 
0 lesson  ��  ` n      yzy  ;  |}z o  y|���� 0 lessons6  ^ {|{ E  ��}~} n  ��� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ~ m  ���� ���  T h u r s d a y| ��� r  ����� J  ���� ��� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ���� b  ����� b  ����� n  ����� 1  ����
�� 
dstr� l �������� [  ����� o  ������ 0 firstday firstDay� l �������� ]  ����� 1  ����
�� 
days� m  ������ ��  ��  ��  ��  � m  ���� ���    a t  � n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  � n      ���  ;  ��� o  ������ 0 lessons6  � ��� E  ����� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � m  ���� ���  F r i d a y� ���� r  ����� J  ���� ��� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ���� b  ����� b  ����� n  ����� 1  ����
�� 
dstr� l �������� [  ����� o  ������ 0 firstday firstDay� l �������� ]  ����� 1  ����
�� 
days� m  ������ ��  ��  ��  ��  � m  ���� ���    a t  � n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  � n      ���  ;  ��� o  ������ 0 lessons6  ��  ��  �� 
0 lesson   o  ������ 0 lessons5  ��  ��  	 ��� l     ��������  ��  ��  � ��� l     ������  � ) # convert each item to a date object   � ��� F   c o n v e r t   e a c h   i t e m   t o   a   d a t e   o b j e c t� ��� l �������� r  ����� J  ������  � o      ���� 0 lessons7  ��  ��  � ��� l �	,������ X  �	,����� k  		'�� ��� r  		��� n  		��� 4  		���
�� 
cobj� m  		���� � o  		���� 
0 lesson  � o      ���� 0 
datestring 
dateString� ��� r  		��� 4  		���
�� 
ldt � o  		���� 0 
datestring 
dateString� o      ���� 0 thedate theDate� ���� r  		'��� J  		"�� ��� n  		��� 4  		���
�� 
cobj� m  		���� � o  		���� 
0 lesson  � ���� o  		 ���� 0 thedate theDate��  � n      ���  ;  	%	&� o  	"	%�� 0 lessons7  ��  �� 
0 lesson  � o  ���~�~ 0 lessons6  ��  ��  � ��� l     �}�|�{�}  �|  �{  � ��� l     �z���z  � #  Add room numbers to the list   � ��� :   A d d   r o o m   n u m b e r s   t o   t h e   l i s t� ��� l 	-	3��y�x� r  	-	3��� J  	-	/�w�w  � o      �v�v 0 lessons8  �y  �x  � ��� l 	4	� �u�t  X  	4	��s Z  	H	��r E  	H	P n  	H	L	 4  	I	L�q

�q 
cobj
 m  	J	K�p�p 	 o  	H	I�o�o 
0 lesson   m  	L	O �  / r  	S	e J  	S	`  n  	S	W 4  	T	W�n
�n 
cobj m  	U	V�m�m  o  	S	T�l�l 
0 lesson    n  	W	[ 4  	X	[�k
�k 
cobj m  	Y	Z�j�j  o  	W	X�i�i 
0 lesson   �h m  	[	^ �  M u l t i p l e�h   n        ;  	c	d o  	`	c�g�g 0 lessons8  �r   X  	h	��f  Z  	|	�!"�e�d! E  	|	�#$# n  	|	�%&% 4  	}	��c'
�c 
cobj' m  	~	�b�b & o  	|	}�a�a 
0 lesson  $ n  	�	�()( 4  	�	��`*
�` 
cobj* m  	�	��_�_ ) o  	�	��^�^ 0 room  " r  	�	�+,+ J  	�	�-- ./. n  	�	�010 4  	�	��]2
�] 
cobj2 m  	�	��\�\ 1 o  	�	��[�[ 
0 lesson  / 343 n  	�	�565 4  	�	��Z7
�Z 
cobj7 m  	�	��Y�Y 6 o  	�	��X�X 
0 lesson  4 8�W8 n  	�	�9:9 4  	�	��V;
�V 
cobj; m  	�	��U�U : o  	�	��T�T 0 room  �W  , n      <=<  ;  	�	�= o  	�	��S�S 0 lessons8  �e  �d  �f 0 room    o  	k	n�R�R 0 roomnumbers roomNumbers�s 
0 lesson   o  	7	:�Q�Q 0 lessons7  �u  �t  � >?> l     �P�O�N�P  �O  �N  ? @A@ l      �MBC�M  B � �
set classNames to {}
repeat with lesson in lessons8
	set end of classNames to item 1 of lesson
end repeat

set classNames2 to addr(classNames)
   C �DD  
 s e t   c l a s s N a m e s   t o   { } 
 r e p e a t   w i t h   l e s s o n   i n   l e s s o n s 8 
 	 s e t   e n d   o f   c l a s s N a m e s   t o   i t e m   1   o f   l e s s o n 
 e n d   r e p e a t 
 
 s e t   c l a s s N a m e s 2   t o   a d d r ( c l a s s N a m e s ) 
A EFE l 	�	�G�L�KG r  	�	�HIH J  	�	��J�J  I o      �I�I 0 lessons9  �L  �K  F JKJ l     �HLM�H  L &  set classNamesTeacherFirst to {}   M �NN @ s e t   c l a s s N a m e s T e a c h e r F i r s t   t o   { }K OPO l 	�
�Q�G�FQ X  	�
�R�ESR k  	�
�TT UVU Z  	�
�WX�DYW E  	�	�Z[Z n  	�	�\]\ 4  	�	��C^
�C 
cobj^ m  	�	��B�B ] o  	�	��A�A 
0 lesson  [ m  	�	�__ �``  /X k  	�
Aaa bcb r  	�	�ded n  	�	�fgf 4  	�	��@h
�@ 
cworh m  	�	��?�? g n  	�	�iji 4  	�	��>k
�> 
cobjk m  	�	��=�= j o  	�	��<�< 
0 lesson  e o      �;�; 0 teacherfirst teacherFirstc lml r  	�	�non n  	�	�pqp 4  	�	��:r
�: 
cworr m  	�	��9�9 q n  	�	�sts 4  	�	��8u
�8 
cobju m  	�	��7�7 t o  	�	��6�6 
0 lesson  o o      �5�5 0 teachersecond teacherSecondm vwv r  	�	�xyx n  	�	�z{z 4  	�	��4|
�4 
cwor| m  	�	��3�3 { n  	�	�}~} 4  	�	��2
�2 
cobj m  	�	��1�1 ~ o  	�	��0�0 
0 lesson  y o      �/�/ 0 yearword yearWordw ��� r  	�
��� n  	�
��� 4  

�.�
�. 
cwor� m  

�-�- � n  	�
��� 4  	�
�,�
�, 
cobj� m  
 
�+�+ � o  	�	��*�* 
0 lesson  � o      �)�) 0 
yearnumber 
yearNumber� ��� r  

��� n  

��� 4  

�(�
�( 
cwor� m  

�'�' � n  

��� 4  

�&�
�& 
cobj� m  

�%�% � o  

�$�$ 
0 lesson  � o      �#�# 0 	classname 	className� ��"� r  

A��� b  

=��� b  

9��� b  

5��� b  

1��� b  

-��� b  

)��� b  

%��� b  

!��� o  

�!�! 0 teacherfirst teacherFirst� m  

 �� ���  /� o  
!
$� �  0 teachersecond teacherSecond� m  
%
(�� ���   � o  
)
,�� 0 yearword yearWord� m  
-
0�� ���   � o  
1
4�� 0 
yearnumber 
yearNumber� m  
5
8�� ���   � o  
9
<�� 0 	classname 	className� o      �� 0 	eventname 	eventName�"  �D  Y k  
D
��� ��� r  
D
S��� n  
D
O��� 4  
H
O��
� 
cwor� m  
K
N�� � n  
D
H��� 4  
E
H��
� 
cobj� m  
F
G�� � o  
D
E�� 
0 lesson  � o      �� 0 teacher  � ��� r  
T
a��� n  
T
]��� 4  
X
]��
� 
cwor� m  
[
\�� � n  
T
X��� 4  
U
X��
� 
cobj� m  
V
W�� � o  
T
U�� 
0 lesson  � o      �� 0 yearword yearWord� ��� r  
b
o��� n  
b
k��� 4  
f
k��
� 
cwor� m  
i
j�� � n  
b
f��� 4  
c
f��
� 
cobj� m  
d
e�� � o  
b
c�� 
0 lesson  � o      �
�
 0 
yearnumber 
yearNumber� ��� r  
p
}��� n  
p
y��� 4  
t
y�	�
�	 
cwor� m  
w
x�� � n  
p
t��� 4  
q
t��
� 
cobj� m  
r
s�� � o  
p
q�� 
0 lesson  � o      �� 0 	classname 	className� ��� r  
~
���� b  
~
���� b  
~
���� b  
~
���� b  
~
���� b  
~
���� b  
~
���� o  
~
��� 0 teacher  � m  
�
��� ���   � o  
�
��� 0 yearword yearWord� m  
�
��� ���   � o  
�
�� �  0 
yearnumber 
yearNumber� m  
�
��� ���   � o  
�
����� 0 	classname 	className� o      ���� 0 	eventname 	eventName�  V ���� r  
�
���� J  
�
��� ��� o  
�
����� 0 	eventname 	eventName� ��� n  
�
���� 4  
�
����
�� 
cobj� m  
�
����� � o  
�
����� 
0 lesson  � ���� n  
�
���� 4  
�
����
�� 
cobj� m  
�
����� � o  
�
����� 
0 lesson  ��  � n      ���  ;  
�
�� o  
�
����� 0 lessons9  ��  �E 
0 lesson  S o  	�	����� 0 lessons8  �G  �F  P ��� l     ��������  ��  ��  �    l     ����   . ( Creates the calendars in the chosen app    � P   C r e a t e s   t h e   c a l e n d a r s   i n   t h e   c h o s e n   a p p  l 
������ Z  
��	��
 = 
�
� o  
�
����� &0 chosencalendarapp chosenCalendarApp m  
�
� �  A p p l e   C a l e n d a r	 k  
��  l 
�
�����   &   Deletes previous work calendars    � @   D e l e t e s   p r e v i o u s   w o r k   c a l e n d a r s  O  
�
� k  
�
�  I 
�
�������
�� .miscactvnull��� ��� null��  ��    O 
�
� I 
�
��� ��
�� .sysottosnull���     TEXT  m  
�
�!! �"" n D e l e t i n g   p r e v i o u s   C a l e n d a r s   w i t h   t h e   w o r d   " Y e a r "   i n   i t .��    f  
�
� #$# l 
�
���������  ��  ��  $ %��% O  
�
�&'& I 
�
�������
�� .coredelonull���     obj ��  ��  ' l 
�
�(����( 6 
�
�)*) 2 
�
���
�� 
wres* E  
�
�+,+ 1  
�
���
�� 
pnam, m  
�
�-- �..  Y e a r��  ��  ��   m  
�
�//�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��   010 l 
�
���������  ��  ��  1 232 l 
�
���������  ��  ��  3 454 l 
�
���67��  6 3 - add calendars and events to the Calendar app   7 �88 Z   a d d   c a l e n d a r s   a n d   e v e n t s   t o   t h e   C a l e n d a r   a p p5 9:9 I 
�
���;��
�� .sysottosnull���     TEXT; m  
�
�<< �== $ C r e a t i n g   C a l e n d a r s��  : >?> O  
�R@A@ k  QBB CDC I 	������
�� .miscactvnull��� ��� null��  ��  D EFE Y  
EG��HI��G Z  @JK����J H  +LL l *M����M I *��N��
�� .coredoexnull���     ****N l &O����O 4  &��P
�� 
wresP l %Q����Q n  %RSR 4  !$��T
�� 
cobjT o  "#���� 0 i  S o  !���� 0 
classnames 
classNames��  ��  ��  ��  ��  ��  ��  K I .<����U
�� .wrbtaec2null��� ��� null��  U ��V��
�� 
wtnmV n  28WXW 4  58��Y
�� 
cobjY o  67���� 0 i  X o  25���� 0 
classnames 
classNames��  ��  ��  �� 0 i  H m  ���� I I ��Z��
�� .corecnte****       ****Z o  ���� 0 
classnames 
classNames��  ��  F [\[ O FR]^] I JQ��_��
�� .sysottosnull���     TEXT_ m  JM`` �aa  C r e a t i n g   E v e n t s��  ^  f  FG\ b��b X  SQc��dc Z  gLef��ge I gu��h��
�� .coredoexnull���     ****h 4  gq��i
�� 
wresi l kpj����j n  kpklk 4  lo��m
�� 
cobjm m  mn���� l o  kl���� 
0 lesson  ��  ��  ��  f O  x�non I ������p
�� .corecrel****      � null��  p ��qr
�� 
koclq m  ����
�� 
wrevr ��s��
�� 
prdts K  ��tt ��uv
�� 
wr11u n  ��wxw 4  ����y
�� 
cobjy m  ������ x o  ������ 
0 lesson  v ��z{
�� 
wr14z n  ��|}| 4  ����~
�� 
cobj~ m  ������ } o  ������ 
0 lesson  { ���
�� 
wr1s n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ����
�� 
wr5s� [  ����� l �������� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  ��  � l �������� ]  ����� 1  ����
�� 
min � m  ������ P��  ��  � �����
�� 
wr15� o  ������ "0 endofrecurrence endOfRecurrence��  ��  o 4  x����
�� 
wres� l |������� n  |���� 4  }����
�� 
cobj� m  ~���� � o  |}���� 
0 lesson  ��  ��  ��  g k  �L�� ��� r  ����� n  ����� 2 ����
�� 
cwor� n  ����� 4  ����
� 
cobj� m  ���~�~ � o  ���}�} 
0 lesson  � o      �|�| 0 problem  � ��� r  ���� b  ����� b  ����� b  ����� b  ����� b  ����� b  ����� n  ����� 4  ���{�
�{ 
cobj� m  ���z�z � o  ���y�y 0 problem  � m  ���� ���   � n  ����� 4  ���x�
�x 
cobj� m  ���w�w � o  ���v�v 0 problem  � m  ���� ���   � n  ����� 4  ���u�
�u 
cobj� m  ���t�t � o  ���s�s 0 problem  � m  ���� ���   � n  ����� 4  ���r�
�r 
cobj� m  ���q�q � o  ���p�p 0 problem  � o      �o�o "0 neweventsummary newEventSummary� ��n� O  L��� I K�m�l�
�m .corecrel****      � null�l  � �k��
�k 
kocl� m  �j
�j 
wrev� �i��h
�i 
prdt� K  E�� �g��
�g 
wr11� n  ��� 4  �f�
�f 
cobj� m  �e�e � o  �d�d 
0 lesson  � �c��
�c 
wr14� n  !%��� 4  "%�b�
�b 
cobj� m  #$�a�a � o  !"�`�` 
0 lesson  � �_��
�_ 
wr1s� n  (,��� 4  ),�^�
�^ 
cobj� m  *+�]�] � o  ()�\�\ 
0 lesson  � �[��
�[ 
wr5s� [  /;��� l /3��Z�Y� n  /3��� 4  03�X�
�X 
cobj� m  12�W�W � o  /0�V�V 
0 lesson  �Z  �Y  � l 3:��U�T� ]  3:��� 1  36�S
�S 
min � m  69�R�R P�U  �T  � �Q��P
�Q 
wr15� o  >A�O�O "0 endofrecurrence endOfRecurrence�P  �h  � 4  
�N�
�N 
wres� o  	�M�M "0 neweventsummary newEventSummary�n  �� 
0 lesson  d o  VY�L�L 0 lessons9  ��  A m  
����                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  ? ��� l SS�K�J�I�K  �J  �I  � ��� l SS�H�G�F�H  �G  �F  � ��� l SS�E���E  �   Sets the time zone   � ��� &   S e t s   t h e   t i m e   z o n e� ��� O S_��� I W^�D��C
�D .sysottosnull���     TEXT� m  WZ�� ��� ^ C h a n g i n g   t h e   t i m e   z o n e   t o   t h e   C h i n e s e   t i m e   z o n e�C  �  f  ST� ��� l ``�B�A�@�B  �A  �@  � ��� O  `���� r  f���� 6 f~��� n  fo   1  ko�?
�? 
pnam 2 fk�>
�> 
wres� E  r} 1  sw�=
�= 
pnam m  x| �  Y e a r� o      �<�< 0 thecalendars theCalendars� m  `c�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  �  l ���;�:�9�;  �:  �9   	
	 r  �� I ���8�7 z�6 
�6 .!Cls!fstnull��� ��� null�8  �7   o      �5�5 0 thestore theStore
  l ���4�3�2�4  �3  �2    r  �� o  ���1�1 0 firstday firstDay o      �0�0 0 	startdate 	startDate  r  �� [  �� o  ���/�/ 0 	startdate 	startDate ]  �� 1  ���.
�. 
days m  ���-�-  o      �,�, 0 enddate endDate  l ���+�*�)�+  �*  �)    r  �� !  I ��"#$" z�( 
�( .!CLs!fcAnull���     ****# o  ���'�' 0 thecalendars theCalendars$ �&%&
�& 
!CtY% m  ��'' z�% 
�% !Tct!TtL& �$(�#
�$ 
!Cst( o  ���"�" 0 thestore theStore�#  ! o      �!�! 0 thecalendars theCalendars )*) l ��� ���   �  �  * +,+ r  �-.- I �/�0/ z� 
� .!CLs!feSnull��� ��� null�  0 �12
� 
!Csd1 o  ���� 0 	startdate 	startDate2 �34
� 
!Ced3 o  ���� 0 enddate endDate4 �56
� 
!Csc5 o  ���� 0 thecalendars theCalendars6 �7�
� 
!Cst7 o  ���� 0 thestore theStore�  . o      �� 0 	theevents 	theEvents, 898 l ����  �  �  9 :;: Y  c<�=>�< k  ^?? @A@ r  >BCB I :D�ED z� 
� .!CLs!mo2null��� ��� null�  E �
FG
�
 
!CevF n  '-HIH 4  *-�	J
�	 
cobjJ o  +,�� 0 i  I o  '*�� 0 	theevents 	theEventsG �K�
� 
!CtzK m  03LL �MM  A s i a / S h a n g h a i�  C o      �� 0 thenewevent theNewEventA N�N I ?^O�PO z� 
� .!CLs!updnull��� ��� null�  P � QR
�  
!CevQ o  NQ���� 0 thenewevent theNewEventR ��S��
�� 
!CstS o  TW���� 0 thestore theStore��  �  � 0 i  = m  ���� > I ��T��
�� .corecnte****       ****T o  ���� 0 	theevents 	theEvents��  �  ; UVU l dd��������  ��  ��  V WXW l dd��������  ��  ��  X YZY l dd��[\��  [ + % remove days because of delayed start   \ �]] J   r e m o v e   d a y s   b e c a u s e   o f   d e l a y e d   s t a r tZ ^_^ Z  d/`a����` = dkbcb o  dg����  0 delayedyesorno delayedYesOrNoc m  gjdd �ee  Y e sa k  n+ff ghg r  n�iji I nk����k z�� 
�� .!Cls!fstnull��� ��� null��  ��  j o      ���� 0 thestore theStoreh lml l ����������  ��  ��  m non r  ��pqp I ��rstr z�� 
�� .!CLs!fcAnull���     ****s J  ������  t ��uv
�� 
!CtYu m  ��ww z�� 
�� !Tct!TtCv ��x��
�� 
!Cstx o  ������ 0 thestore theStore��  q o      ���� 0 thecalendars theCalendarso yzy r  ��{|{ o  ������ 0 firstday firstDay| o      ���� 0 startingdate startingDatez }~} r  ��� \  ����� o  ������ "0 delayedfirstday delayedFirstDay� l �������� ]  ����� 1  ����
�� 
days� m  ������ ��  ��  � o      ���� 0 
endingdate 
endingDate~ ��� r  ����� I ������� z�� 
�� .!CLs!feSnull��� ��� null��  � ����
�� 
!Csd� o  ������ 0 startingdate startingDate� ����
�� 
!Ced� o  ������ 0 
endingdate 
endingDate� ����
�� 
!Csc� o  ������ 0 thecalendars theCalendars� �����
�� 
!Cst� o  ������ 0 thestore theStore��  � o      ���� 0 	theevents 	theEvents� ��� l ����������  ��  ��  � ��� X  �)����� I $����� z�� 
�� .!CLs!reEnull��� ��� null��  � ����
�� 
!Cev� o  ���� 0 theevent theEvent� ����
�� 
!Cst� o  ���� 0 thestore theStore� �����
�� 
!Cft� m  ��
�� boovfals��  �� 0 theevent theEvent� o  ������ 0 	theevents 	theEvents� ���� l **��������  ��  ��  ��  ��  ��  _ ��� l 00��������  ��  ��  � ��� l 00������  � k e This turns off and on the iCloud Calendar in preferences so the calendars gets uploaded to the cloud   � ��� �   T h i s   t u r n s   o f f   a n d   o n   t h e   i C l o u d   C a l e n d a r   i n   p r e f e r e n c e s   s o   t h e   c a l e n d a r s   g e t s   u p l o a d e d   t o   t h e   c l o u d� ��� I 07�����
�� .sysottosnull���     TEXT� m  03�� ��� J M o v i n g   C a l e n d a r s   f r o m   l o c a l   t o   i C l o u d��  � ��� O  8\��� k  >[�� ��� I >C������
�� .miscactvnull��� ��� null��  ��  � ��� I DI�����
�� .sysodelanull��� ��� nmbr� m  DE���� ��  � ��� I JO������
�� .aevtquitnull��� ��� null��  ��  � ��� I PU������
�� .miscactvnull��� ��� null��  ��  � ���� I V[�����
�� .sysodelanull��� ��� nmbr� m  VW���� ��  ��  � m  8;���                                                                                  sprf  alis    `  Macintosh HD                   BD ����System Preferences.app                                         ����            ����  
 cu             Applications  -/:System:Applications:System Preferences.app/   .  S y s t e m   P r e f e r e n c e s . a p p    M a c i n t o s h   H D  *System/Applications/System Preferences.app  / ��  � ��� l ]]��������  ��  ��  � ��� O  ]���� O  c���� O  n���� O  w���� k  ���� ��� I ��������
�� .prcsclicnull��� ��� uiel��  ��  � ���� I �������
�� .sysodelanull��� ��� nmbr� m  ������ ��  ��  � 4  w}���
�� 
butT� m  {|���� � 4  nt���
�� 
cwin� m  rs���� � 4  ck���
�� 
prcs� m  gj�� ��� $ S y s t e m   P r e f e r e n c e s� m  ]`���                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � ��� l ����������  ��  ��  � ��� O  ���� O  ���� O  ���� O  ���� O  ���� O  ���� O  ���� O  ���� O  ���� k  ��� ��� O ����� I �������
�� .sysottosnull���     TEXT� m  ���� ��� 8 U n c h e c k i n g   i C l o u d   C a l e n d a r s .��  �  f  ��� ��� l ������ I ��������
�� .prcsclicnull��� ��� uiel��  ��  � #  turn off icloud for calendar   � ��� :   t u r n   o f f   i c l o u d   f o r   c a l e n d a r� ��� I �������
�� .sysodelanull��� ��� nmbr� m  ������ ��  � ��� O �	��� I ��	 ��
�� .sysottosnull���     TEXT	  m  		 �		 4 C h e c k i n g   i C l o u d   C a l e n d a r s .��  �  f  ��� 			 l 
				 I 
������
�� .prcsclicnull��� ��� uiel��  ��  	 "  turn on icloud for calendar   	 �		 8   t u r n   o n   i c l o u d   f o r   c a l e n d a r	 		��		 I ��	
��
�� .sysodelanull��� ��� nmbr	
 m  ���� ��  ��  � 4  ���	
� 
uiel	 m  ���~�~ � 4  ���}	
�} 
uiel	 m  ���|�| � 4  ���{	
�{ 
crow	 m  ���z�z � 4  ���y	
�y 
tabB	 m  ���x�x � 4  ���w	
�w 
scra	 m  ���v�v � 4  ���u	
�u 
sgrp	 m  ���t�t � 4  ���s	
�s 
cwin	 m  ���r�r � 4  ���q	
�q 
prcs	 m  ��		 �		 $ S y s t e m   P r e f e r e n c e s� m  ��		�                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � 			 l  �p		�p  	 � �
	tell application "System Events"
		tell process "System Preferences"
			tell window 1
				tell sheet 1
					tell button 2
						click #click the merge button
						delay 90
					end tell
				end tell
			end tell
		end tell
	end tell
	   	 �		� 
 	 t e l l   a p p l i c a t i o n   " S y s t e m   E v e n t s " 
 	 	 t e l l   p r o c e s s   " S y s t e m   P r e f e r e n c e s " 
 	 	 	 t e l l   w i n d o w   1 
 	 	 	 	 t e l l   s h e e t   1 
 	 	 	 	 	 t e l l   b u t t o n   2 
 	 	 	 	 	 	 c l i c k   # c l i c k   t h e   m e r g e   b u t t o n 
 	 	 	 	 	 	 d e l a y   9 0 
 	 	 	 	 	 e n d   t e l l 
 	 	 	 	 e n d   t e l l 
 	 	 	 e n d   t e l l 
 	 	 e n d   t e l l 
 	 e n d   t e l l 
 		 			 O  3			 k  %2		 	 	!	  I %*�o�n�m
�o .miscactvnull��� ��� null�n  �m  	! 	"�l	" I +2�k	#�j
�k .sysodelanull��� ��� nmbr	# m  +.�i�i Z�j  �l  	 m  "	$	$�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  	 	%	&	% l 44�h�g�f�h  �g  �f  	& 	'	(	' l 44�e	)	*�e  	) ( " Publishes the Calendars as Public   	* �	+	+ D   P u b l i s h e s   t h e   C a l e n d a r s   a s   P u b l i c	( 	,	-	, I 4;�d	.�c
�d .sysottosnull���     TEXT	. m  47	/	/ �	0	0 \ P u b l i s h i n g   t h e   c a l e n d a r s   b y   m a k i n g   t h e m   p u b l i c�c  	- 	1	2	1 O  <�	3	4	3 O  B�	5	6	5 k  M�	7	7 	8	9	8 l MM�b�a�`�b  �a  �`  	9 	:	;	: l  MM�_	<	=�_  	<��
				keystroke "f" using command down # search bar
				delay 1
				keystroke tab # moves focus to calendar list
				key code 126 using {option down} # moves focus to first calendar / option up arrow
				
				key code 125 # press down arrow
				
				tell menu "Edit" of menu bar item "Edit" of menu bar 1
					click # click Edit Menu
				end tell
				
				delay 1
				
				tell UI element 15 of menu "Edit" of menu bar item "Edit" of menu bar 1
					click # click Share Calendar under Edit Menu
				end tell
				
				delay 1
				
				tell button "Done" of pop over 1 of row 3 of outline 1 of scroll area 1 of splitter group 1 of splitter group 1 of window 1
					click
				end tell
				
				delay 1
				   	= �	>	>| 
 	 	 	 	 k e y s t r o k e   " f "   u s i n g   c o m m a n d   d o w n   #   s e a r c h   b a r 
 	 	 	 	 d e l a y   1 
 	 	 	 	 k e y s t r o k e   t a b   #   m o v e s   f o c u s   t o   c a l e n d a r   l i s t 
 	 	 	 	 k e y   c o d e   1 2 6   u s i n g   { o p t i o n   d o w n }   #   m o v e s   f o c u s   t o   f i r s t   c a l e n d a r   /   o p t i o n   u p   a r r o w 
 	 	 	 	 
 	 	 	 	 k e y   c o d e   1 2 5   #   p r e s s   d o w n   a r r o w 
 	 	 	 	 
 	 	 	 	 t e l l   m e n u   " E d i t "   o f   m e n u   b a r   i t e m   " E d i t "   o f   m e n u   b a r   1 
 	 	 	 	 	 c l i c k   #   c l i c k   E d i t   M e n u 
 	 	 	 	 e n d   t e l l 
 	 	 	 	 
 	 	 	 	 d e l a y   1 
 	 	 	 	 
 	 	 	 	 t e l l   U I   e l e m e n t   1 5   o f   m e n u   " E d i t "   o f   m e n u   b a r   i t e m   " E d i t "   o f   m e n u   b a r   1 
 	 	 	 	 	 c l i c k   #   c l i c k   S h a r e   C a l e n d a r   u n d e r   E d i t   M e n u 
 	 	 	 	 e n d   t e l l 
 	 	 	 	 
 	 	 	 	 d e l a y   1 
 	 	 	 	 
 	 	 	 	 t e l l   b u t t o n   " D o n e "   o f   p o p   o v e r   1   o f   r o w   3   o f   o u t l i n e   1   o f   s c r o l l   a r e a   1   o f   s p l i t t e r   g r o u p   1   o f   s p l i t t e r   g r o u p   1   o f   w i n d o w   1 
 	 	 	 	 	 c l i c k 
 	 	 	 	 e n d   t e l l 
 	 	 	 	 
 	 	 	 	 d e l a y   1 
 	 	 	 		; 	?	@	? l MM�^�]�\�^  �]  �\  	@ 	A	B	A l MZ	C	D	E	C I MZ�[	F	G
�[ .prcskprsnull���     ctxt	F m  MP	H	H �	I	I  f	G �Z	J�Y
�Z 
faal	J m  SV�X
�X eMdsKcmd�Y  	D   search bar   	E �	K	K    s e a r c h   b a r	B 	L	M	L I [`�W	N�V
�W .sysodelanull��� ��� nmbr	N m  [\�U�U �V  	M 	O	P	O l aa�T�S�R�T  �S  �R  	P 	Q	R	Q l ah	S	T	U	S I ah�Q	V�P
�Q .prcskprsnull���     ctxt	V 1  ad�O
�O 
tab �P  	T #  moves focus to calendar list   	U �	W	W :   m o v e s   f o c u s   t o   c a l e n d a r   l i s t	R 	X	Y	X I in�N	Z�M
�N .sysodelanull��� ��� nmbr	Z m  ij�L�L �M  	Y 	[	\	[ l oo�K�J�I�K  �J  �I  	\ 	]	^	] l o~	_	`	a	_ I o~�H	b	c
�H .prcskcodnull���     ****	b m  or�G�G ~	c �F	d�E
�F 
faal	d J  uz	e	e 	f�D	f m  ux�C
�C eMdsKopt�D  �E  	` 6 0 moves focus to first calendar / option up arrow   	a �	g	g `   m o v e s   f o c u s   t o   f i r s t   c a l e n d a r   /   o p t i o n   u p   a r r o w	^ 	h	i	h I ��B	j�A
�B .sysodelanull��� ��� nmbr	j m  ��@�@ �A  	i 	k	l	k l ���?�>�=�?  �>  �=  	l 	m	n	m r  ��	o	p	o J  ���<�<  	p o      �;�; 0 calendarnames calendarNames	n 	q�:	q Y  ��	r�9	s	t�8	r k  ��	u	u 	v	w	v Q  ��	x	y�7	x k  ��	z	z 	{	|	{ r  ��	}	~	} n  ��		�	 1  ���6
�6 
valL	� n  ��	�	�	� 4  ���5	�
�5 
txtf	� m  ���4�4 	� n  ��	�	�	� 4  ���3	�
�3 
uiel	� m  ���2�2 	� n  ��	�	�	� 4  ���1	�
�1 
crow	� o  ���0�0 0 i  	� n  ��	�	�	� 4  ���/	�
�/ 
outl	� m  ���.�. 	� n  ��	�	�	� 4  ���-	�
�- 
scra	� m  ���,�, 	� n  ��	�	�	� 4  ���+	�
�+ 
splg	� m  ���*�* 	� n  ��	�	�	� 4  ���)	�
�) 
splg	� m  ���(�( 	� 4  ���'	�
�' 
cwin	� m  ���&�& 	~ o      �%�% $0 calendarnametext calendarNameText	| 	��$	� Z  ��	�	��#�"	� E  ��	�	�	� o  ���!�! $0 calendarnametext calendarNameText	� m  ��	�	� �	�	�  Y e a r	� k  ��	�	� 	�	�	� r  ��	�	�	� o  ��� �  $0 calendarnametext calendarNameText	� n      	�	�	�  ;  ��	� o  ���� 0 calendarnames calendarNames	� 	�	�	� l �	�	�	�	� I ��	�	�
� .prcskprsnull���     ctxt	� m  �	�	� �	�	�  f	� �	��
� 
faal	� m  �
� eMdsKcmd�  	�   search bar   	� �	�	�    s e a r c h   b a r	� 	�	�	� I �	��
� .sysodelanull��� ��� nmbr	� m  �� �  	� 	�	�	� l ����  �  �  	� 	�	�	� l 	�	�	�	� I �	��
� .prcskprsnull���     ctxt	� 1  �
� 
tab �  	� #  moves focus to calendar list   	� �	�	� :   m o v e s   f o c u s   t o   c a l e n d a r   l i s t	� 	�	�	� l �	�	��  	�  						delay 1   	� �	�	�  	 	 	 	 	 	 d e l a y   1	� 	�	�	� l ����  �  �  	� 	�	�	� l "	�	�	�	� I "�	��
� .prcskcodnull���     ****	� m  �� }�  	�   down arrow   	� �	�	�    d o w n   a r r o w	� 	�	�	� I #(�
	��	
�
 .sysodelanull��� ��� nmbr	� m  #$�� �	  	� 	�	�	� l ))����  �  �  	� 	�	�	� O  )F	�	�	� l @E	�	�	�	� I @E���
� .prcsclicnull��� ��� uiel�  �  	�   click Edit Menu   	� �	�	�     c l i c k   E d i t   M e n u	� n  )=	�	�	� 4  6=�	�
� 
menE	� m  9<	�	� �	�	�  E d i t	� n  )6	�	�	� 4  /6� 	�
�  
mbri	� m  25	�	� �	�	�  E d i t	� 4  )/��	�
�� 
mbar	� m  -.���� 	� 	�	�	� l GG��	�	���  	�  						delay 1   	� �	�	�  	 	 	 	 	 	 d e l a y   1	� 	�	�	� O  Gk	�	�	� l ej	�	�	�	� I ej������
�� .prcsclicnull��� ��� uiel��  ��  	� + % click Share Calendar under Edit Menu   	� �	�	� J   c l i c k   S h a r e   C a l e n d a r   u n d e r   E d i t   M e n u	� n  Gb	�	�	� 4  [b��	�
�� 
uiel	� m  ^a���� 	� n  G[	�	�	� 4  T[��	�
�� 
menE	� m  WZ	�	� �	�	�  E d i t	� n  GT	�	�	� 4  MT��	�
�� 
mbri	� m  PS	�	� �	�	�  E d i t	� 4  GM��	�
�� 
mbar	� m  KL���� 	� 	�	�	� l ll��	�
 ��  	�  						delay 1   
  �

  	 	 	 	 	 	 d e l a y   1	� 


 O  l�


 l ��



 I ��������
�� .prcsclicnull��� ��� uiel��  ��  
 ( " click the share calendar checkbox   
 �
	
	 D   c l i c k   t h e   s h a r e   c a l e n d a r   c h e c k b o x
 n  l�




 4  ����

�� 
chbx
 m  ������ 
 n  l�


 4  ����

�� 
popv
 m  ������ 
 n  l�


 4  ����

�� 
crow
 o  ������ 0 i  
 n  l�


 4  ����

�� 
outl
 m  ������ 
 n  l�


 4  |���

�� 
scra
 m  ����� 
 n  l|


 4  w|��

�� 
splg
 m  z{���� 
 n  lw


 4  rw��

�� 
splg
 m  uv���� 
 4  lr��

�� 
cwin
 m  pq���� 
 
 
!
  l ����
"
#��  
"  						delay 1   
# �
$
$  	 	 	 	 	 	 d e l a y   1
! 
%
&
% O  ��
'
(
' I ��������
�� .prcsclicnull��� ��� uiel��  ��  
( n  ��
)
*
) 4  ����
+
�� 
butT
+ m  ��
,
, �
-
-  D o n e
* n  ��
.
/
. 4  ����
0
�� 
popv
0 m  ������ 
/ n  ��
1
2
1 4  ����
3
�� 
crow
3 o  ������ 0 i  
2 n  ��
4
5
4 4  ����
6
�� 
outl
6 m  ������ 
5 n  ��
7
8
7 4  ����
9
�� 
scra
9 m  ������ 
8 n  ��
:
;
: 4  ����
<
�� 
splg
< m  ������ 
; n  ��
=
>
= 4  ����
?
�� 
splg
? m  ������ 
> 4  ����
@
�� 
cwin
@ m  ������ 
& 
A
B
A l ����
C
D��  
C 7 1					keystroke "f" using command down # seach bar   
D �
E
E b 	 	 	 	 	 k e y s t r o k e   " f "   u s i n g   c o m m a n d   d o w n   #   s e a c h   b a r
B 
F
G
F I ����
H��
�� .sysodelanull��� ��� nmbr
H m  ������ ��  
G 
I
J
I l ��
K
L
M
K I ����
N��
�� .prcskprsnull���     ctxt
N 1  ����
�� 
tab ��  
L #  moves focus to calendar list   
M �
O
O :   m o v e s   f o c u s   t o   c a l e n d a r   l i s t
J 
P��
P I ����
Q��
�� .sysodelanull��� ��� nmbr
Q m  ������ ��  ��  �#  �"  �$  	y R      ������
�� .ascrerr ****      � ****��  ��  �7  	w 
R��
R l ����������  ��  ��  ��  �9 0 i  	s m  ������ 	t I ����
S��
�� .corecnte****       ****
S n ��
T
U
T 2 ����
�� 
crow
U n  ��
V
W
V 4  ����
X
�� 
outl
X m  ������ 
W n  ��
Y
Z
Y 4  ����
[
�� 
scra
[ m  ������ 
Z n  ��
\
]
\ 4  ����
^
�� 
splg
^ m  ������ 
] n  ��
_
`
_ 4  ����
a
�� 
splg
a m  ������ 
` 4  ����
b
�� 
cwin
b m  ������ ��  �8  �:  	6 4  BJ��
c
�� 
prcs
c m  FI
d
d �
e
e  C a l e n d a r	4 m  <?
f
f�                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  	2 
g
h
g l ����������  ��  ��  
h 
i
j
i l ����
k
l��  
k N H Grabs the published calendars and writes the URLs to a numbers document   
l �
m
m �   G r a b s   t h e   p u b l i s h e d   c a l e n d a r s   a n d   w r i t e s   t h e   U R L s   t o   a   n u m b e r s   d o c u m e n t
j 
n
o
n O  �

p
q
p I 	������
�� .miscactvnull��� ��� null��  ��  
q m  �
r
r�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  
o 
s
t
s I ��
u��
�� .sysottosnull���     TEXT
u m  
v
v �
w
w X G r a b b i n g   t h e   U R L s   f o r   t h e s e   p u b l i c   c a l e n d a r s��  
t 
x
y
x l ��������  ��  ��  
y 
z
{
z r  
|
}
| J  ����  
} o      ���� 0 calendarurls calendarURLs
{ 
~

~ O  �
�
�
� O   �
�
�
� k  +�
�
� 
�
�
� I +8��
�
�
�� .prcskprsnull���     ctxt
� m  +.
�
� �
�
�  f
� ��
���
�� 
faal
� m  14��
�� eMdsKcmd��  
� 
�
�
� I 9>��
���
�� .sysodelanull��� ��� nmbr
� m  9:���� ��  
� 
�
�
� I ?F��
���
�� .prcskprsnull���     ctxt
� 1  ?B��
�� 
tab ��  
� 
�
�
� I GL��
���
�� .sysodelanull��� ��� nmbr
� m  GH����  ��  
� 
�
�
� I M\��
�
�
�� .prcskcodnull���     ****
� m  MP���� ~
� ��
���
�� 
faal
� J  SX
�
� 
���
� m  SV��
�� eMdsKopt��  ��  
� 
�
�
� I ]b��
���
�� .sysodelanull��� ��� nmbr
� m  ]^����  ��  
� 
�
�
� l cc��������  ��  ��  
� 
���
� Y  c�
���
�
���
� k  s�
�
� 
�
�
� I sz��
���
�� .prcskcodnull���     ****
� m  sv�� }��  
� 
�
�
� I {��~
��}
�~ .sysodelanull��� ��� nmbr
� m  {|�|�|  �}  
� 
�
�
� I ���{
�
�
�{ .prcskprsnull���     ctxt
� m  ��
�
� �
�
�  i
� �z
��y
�z 
faal
� m  ���x
�x eMdsKcmd�y  
� 
�
�
� I ���w
��v
�w .sysodelanull��� ��� nmbr
� m  ���u�u  �v  
� 
�
�
� O  ��
�
�
� O  ��
�
�
� k  ��
�
� 
�
�
� r  ��
�
�
� n  ��
�
�
� 1  ���t
�t 
valL
� 4  ���s
�
�s 
txtf
� m  ���r�r 
� o      �q�q 0 calendarname calendarName
� 
��p
� O  ��
�
�
� O  ��
�
�
� r  ��
�
�
� 1  ���o
�o 
valL
� o      �n�n 0 theurl theURL
� 4  ���m
�
�m 
txta
� m  ���l�l 
� 4  ���k
�
�k 
scra
� m  ���j�j �p  
� 4  ���i
�
�i 
sheE
� m  ���h�h 
� 4  ���g
�
�g 
cwin
� m  ���f�f 
� 
�
�
� I ���e
��d
�e .prcskprsnull���     ctxt
� o  ���c
�c 
ret �d  
� 
�
�
� I ���b
��a
�b .sysodelanull��� ��� nmbr
� m  ���`�`  �a  
� 
��_
� r  ��
�
�
� J  ��
�
� 
�
�
� o  ���^�^ 0 calendarname calendarName
� 
��]
� o  ���\�\ 0 theurl theURL�]  
� n      
�
�
�  ;  ��
� o  ���[�[ 0 calendarurls calendarURLs�_  �� 0 i  
� m  fg�Z�Z 
� I gn�Y
��X
�Y .corecnte****       ****
� o  gj�W�W 0 calendarnames calendarNames�X  ��  ��  
� 4   (�V
�
�V 
prcs
� m  $'
�
� �
�
�  C a l e n d a r
� m  
�
��                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  
 
�
�
� l ���U�T�S�U  �T  �S  
� 
�
�
� I � �R
��Q
�R .sysottosnull���     TEXT
� m  ��
�
� �
�
� p P a s t i n g   p u b l i c   c a l e n d a r   l i n k s   i n t o   a   N u m b e r s   s p r e a d s h e e t�Q  
� 
�
�
� O  �
�
�
� k  �
�
� 
�
�
� I �P�O�N
�P .miscactvnull��� ��� null�O  �N  
� 
�
�
� I �M�L
�
�M .corecrel****      � null�L  
� �K
��J
�K 
kocl
� m  �I
�I 
docu�J  
� 
��H
� O  �
�
�
� O  "�
�
�
� O  +�
�
�
� k  4�
�
� 
�
�
� O  4[
� 
� k  =Z  r  =K m  =@ �  C a l e n d a r   N a m e n      	 1  FJ�G
�G 
NMCv	 4  @F�F

�F 
NmCl
 m  DE�E�E  �D r  LZ m  LO �  U R L n       1  UY�C
�C 
NMCv 4  OU�B
�B 
NmCl m  ST�A�A �D    4  4:�@
�@ 
NMRw m  89�?�? 
�  l \\�>�=�<�>  �=  �<   �; Y  \��:�9 k  l�  r  ly n  lu  4  ru�8!
�8 
cobj! m  st�7�7   n  lr"#" 4  or�6$
�6 
cobj$ o  pq�5�5 0 i  # o  lo�4�4 0 calendarurls calendarURLs o      �3�3 0 calendarname calendarName %&% r  z�'(' n  z�)*) 4  ���2+
�2 
cobj+ m  ���1�1 * n  z�,-, 4  }��0.
�0 
cobj. o  ~�/�/ 0 i  - o  z}�.�. 0 calendarurls calendarURLs( o      �-�- 0 theurl theURL& /0/ Z  ��12�,�+1 H  ��33 l ��4�*�)4 I ���(5�'
�( .coredoexnull���     ****5 4  ���&6
�& 
NMRw6 l ��7�%�$7 [  ��898 o  ���#�# 0 i  9 m  ���"�" �%  �$  �'  �*  �)  2 I ���!� :
�! .corecrel****      � null�   : �;�
� 
kocl; m  ���
� 
NMRw�  �,  �+  0 <=< O  ��>?> k  ��@@ ABA r  ��CDC o  ���� 0 calendarname calendarNameD n      EFE 1  ���
� 
NMCvF 4  ���G
� 
NmClG m  ���� B H�H r  ��IJI o  ���� 0 theurl theURLJ n      KLK 1  ���
� 
NMCvL 4  ���M
� 
NmClM m  ���� �  ? 4  ���N
� 
NMRwN l ��O��O [  ��PQP o  ���� 0 i  Q m  ���� �  �  = R�R l ������  �  �  �  �: 0 i   m  _`�
�
  I `g�	S�
�	 .corecnte****       ****S o  `c�� 0 calendarurls calendarURLs�  �9  �;  
� 4  +1�T
� 
NmTbT m  /0�� 
� 4  "(�U
� 
NmShU m  &'�� 
� 4  �V
� 
docuV m  �� �H  
� m  WW�                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  
� XYX l ��� �����   ��  ��  Y Z��Z l ����������  ��  ��  ��  ��  
 k  ��[[ \]\ l ����^_��  ^ 9 3 Outlook --> Incomplete code, needs repeat function   _ �`` f   O u t l o o k   - - >   I n c o m p l e t e   c o d e ,   n e e d s   r e p e a t   f u n c t i o n] aba l ����cd��  c B < Creates all the events on the first day but does not repeat   d �ee x   C r e a t e s   a l l   t h e   e v e n t s   o n   t h e   f i r s t   d a y   b u t   d o e s   n o t   r e p e a tb fgf O  �qhih k  �pjj klk I ��������
�� .miscactvnull��� ��� null��  ��  l mnm l ����������  ��  ��  n opo I ���q��
�� .coredelonull���     obj q l �r����r 6 �sts 2 ����
�� 
cCFot E  � uvu 1  ����
�� 
pnamv m  ��ww �xx  Y e a r��  ��  ��  p yzy l ��������  ��  ��  z {|{ r  #}~} 6 � 4 ���
�� 
cCFo� m  
���� � E  ��� n  ��� m  ��
�� 
emad� 2 ��
�� 
cAct� m  �� ���  w i n t e c~ o      ���� 0 wintecaccount wintecAccount| ���� O  $p��� k  *o�� ��� l **��������  ��  ��  � ��� Y  *r�������� Z  :m������� H  :K�� l :J������ I :J�����
�� .coredoexnull���     ****� l :F������ 4  :F���
�� 
cCFo� l >E������ n  >E��� 4  AD���
�� 
cobj� o  BC���� 0 i  � o  >A���� 0 
classnames 
classNames��  ��  ��  ��  ��  ��  ��  � I Ni�����
�� .corecrel****      � null��  � ����
�� 
kocl� m  RU��
�� 
cCFo� �����
�� 
prdt� K  Xc�� �����
�� 
pnam� n  [a��� 4  ^a���
�� 
cobj� o  _`���� 0 i  � o  [^���� 0 
classnames 
classNames��  ��  ��  ��  �� 0 i  � m  -.���� � I .5�����
�� .corecnte****       ****� o  .1���� 0 
classnames 
classNames��  ��  � ��� l ss��������  ��  ��  � ��� X  sm����� Z  �h������ I �������
�� .coredoexnull���     ****� 4  �����
�� 
cCFo� l �������� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  ��  ��  � O  ����� I �������
�� .corecrel****      � null��  � ����
�� 
kocl� m  ����
�� 
cEvt� �����
�� 
prdt� K  ���� ����
�� 
subj� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ����
�� 
iloc� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ����
�� 
offs� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ����
�� 
endT� [  ����� l �������� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  ��  � l �������� ]  ����� 1  ����
�� 
min � m  ������ P��  ��  � �����
�� 
pHsR� m  ����
�� boovfals��  ��  � 4  �����
�� 
cCFo� l �������� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  ��  ��  � k  �h�� ��� r  ����� n  ����� 2 ����
�� 
cwor� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � o      ���� 0 problem  � ��� r  ���� b  ���� b  ���� b  ���� b  ���� b  ���� b  ����� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 0 problem  � m  ���� ���   � n  ���� 4  ����
�� 
cobj� m   ���� � o  ������ 0 problem  � m     �   � n   4  
��
�� 
cobj m  ����  o  
���� 0 problem  � m   �   � n   4  ��	
�� 
cobj	 m  ��  o  �~�~ 0 problem  � o      �}�} "0 neweventsummary newEventSummary� 
�|
 O   h I +g�{�z
�{ .corecrel****      � null�z   �y
�y 
kocl m  /2�x
�x 
cEvt �w�v
�w 
prdt K  5a �u
�u 
subj n  8< 4  9<�t
�t 
cobj m  :;�s�s  o  89�r�r 
0 lesson   �q
�q 
iloc n  ?C 4  @C�p
�p 
cobj m  AB�o�o  o  ?@�n�n 
0 lesson   �m
�m 
offs n  FJ 4  GJ�l 
�l 
cobj  m  HI�k�k  o  FG�j�j 
0 lesson   �i!"
�i 
endT! [  MY#$# l MQ%�h�g% n  MQ&'& 4  NQ�f(
�f 
cobj( m  OP�e�e ' o  MN�d�d 
0 lesson  �h  �g  $ l QX)�c�b) ]  QX*+* 1  QT�a
�a 
min + m  TW�`�` P�c  �b  " �_,�^
�_ 
pHsR, m  \]�]
�] boovfals�^  �v   4   (�\-
�\ 
cCFo- o  $'�[�[ "0 neweventsummary newEventSummary�|  �� 
0 lesson  � o  vy�Z�Z 0 lessons9  � .�Y. l nn�X�W�V�X  �W  �V  �Y  � o  $'�U�U 0 wintecaccount wintecAccount��  i m  ��//�                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  g 010 l rr�T�S�R�T  �S  �R  1 2�Q2 O  r�343 k  x�55 676 r  x~898 J  xz�P�P  9 o      �O�O 0 	allevents 	allEvents7 :;: r  �<=< 6 �>?> 4 ��N@
�N 
cCFo@ m  ���M�M ? E  ��ABA n  ��CDC m  ���L
�L 
emadD 2 ���K
�K 
cActB m  ��EE �FF  w i n t e c= o      �J�J 0 wintecaccount wintecAccount; GHG O  ��IJI k  ��KK LML r  ��NON 6 ��PQP 2 ���I
�I 
cCFoQ E  ��RSR 1  ���H
�H 
pnamS m  ��TT �UU  Y e a rO o      �G�G "0 winteccalendars wintecCalendarsM V�FV X  ��W�EXW k  ��YY Z[Z r  ��\]\ n  ��^_^ 2 ���D
�D 
cEvt_ o  ���C�C 0 	acalendar 	aCalendar] o      �B�B 0 	theevents 	theEvents[ `�A` X  ��a�@ba r  ��cdc o  ���?�? 0 anevent anEventd n      efe  ;  ��f o  ���>�> 0 	allevents 	allEvents�@ 0 anevent anEventb o  ���=�= 0 	theevents 	theEvents�A  �E 0 	acalendar 	aCalendarX o  ���<�< "0 winteccalendars wintecCalendars�F  J o  ���;�; 0 wintecaccount wintecAccountH g�:g X   �h�9ih k  �jj klk n  mnm 1  �8
�8 
subjn o  �7�7 0 anevent anEventl opo r  $qrq n   sts 1   �6
�6 
offst o  �5�5 0 anevent anEventr o      �4�4 0 startday startDayp uvu r  %4wxw c  %0yzy n  %,{|{ m  (,�3
�3 
wkdy| o  %(�2�2 0 startday startDayz m  ,/�1
�1 
ctxtx o      �0�0 0 startday startDayv }~} l 55�/�.�-�/  �.  �-  ~ � l 55�,���,  � L F next edit -- change start date into a variable instead of a hard date   � ��� �   n e x t   e d i t   - -   c h a n g e   s t a r t   d a t e   i n t o   a   v a r i a b l e   i n s t e a d   o f   a   h a r d   d a t e� ��� Z  5�����+� = 5<��� o  58�*�* 0 startday startDay� m  8;�� ���  M o n d a y� r  ?u��� K  ?o�� �)��
�) 
ReDw� K  BH�� �(��'
�( 
rDMo� m  EF�&
�& boovtrue�'  � �%��
�% 
ReED� K  K[�� �$��
�$ 
pEnT� m  NQ�#
�# eEDTeEDt� �"��!
�" 
peDa� o  TW� �  0 enddate endDate�!  � ���
� 
tStr� m  ^a�� ldt     �a� � ���
� 
ReOi� m  de�� � ���
� 
ReTy� m  hk�
� eRPTeRwp�  � n      ��� m  pt�
� 
rREc� o  op�� 0 anevent anEvent� ��� = x��� o  x{�� 0 startday startDay� m  {~�� ���  T u e s d a y� ��� r  ����� K  ���� ���
� 
ReDw� K  ���� ���
� 
rDTu� m  ���
� boovtrue�  � ���
� 
ReED� K  ���� ���
� 
pEnT� m  ���
� eEDTeEDt� ���
� 
peDa� o  ���� 0 enddate endDate�  � ���
� 
tStr� m  ���� ldt     �c0�� ���
� 
ReOi� m  ���
�
 � �	��
�	 
ReTy� m  ���
� eRPTeRwp�  � n      ��� m  ���
� 
rREc� o  ���� 0 anevent anEvent� ��� = ����� o  ���� 0 startday startDay� m  ���� ���  W e d n e s d a y� ��� r  ����� K  ���� ���
� 
ReDw� K  ���� ���
� 
rDWe� m  ��� 
�  boovtrue�  � ����
�� 
ReED� K  ���� ����
�� 
pEnT� m  ����
�� eEDTeEDt� �����
�� 
peDa� o  ������ 0 enddate endDate��  � ����
�� 
tStr� m  ���� ldt     �d� � ����
�� 
ReOi� m  ������ � �����
�� 
ReTy� m  ����
�� eRPTeRwp��  � n      ��� m  ����
�� 
rREc� o  ������ 0 anevent anEvent� ��� = ���� o  ����� 0 startday startDay� m  �� ���  T h u r s d a y� ��� r  >��� K  8�� ����
�� 
ReDw� K  �� �����
�� 
rDTh� m  ��
�� boovtrue��  � ����
�� 
ReED� K  $�� ����
�� 
pEnT� m  ��
�� eEDTeEDt� �����
�� 
peDa� o   ���� 0 enddate endDate��  � ����
�� 
tStr� m  '*�� ldt     �eӀ� ����
�� 
ReOi� m  -.���� � �����
�� 
ReTy� m  14��
�� eRPTeRwp��  � n      ��� m  9=��
�� 
rREc� o  89���� 0 anevent anEvent� ��� = AH��� o  AD���� 0 startday startDay� m  DG�� ���  F r i d a y� ���� r  K�   K  K{ ��
�� 
ReDw K  NT ����
�� 
rDFr m  QR��
�� boovtrue��   ��
�� 
ReED K  Wg		 ��

�� 
pEnT
 m  Z]��
�� eEDTeEDt ����
�� 
peDa o  `c���� 0 enddate endDate��   ��
�� 
tStr m  jm ldt     �g%  ��
�� 
ReOi m  pq����  ����
�� 
ReTy m  tw��
�� eRPTeRwp��   n       m  |���
�� 
rREc o  {|���� 0 anevent anEvent��  �+  � �� l ����������  ��  ��  ��  �9 0 anevent anEventi o  ���� 0 	allevents 	allEvents�:  4 m  ru�                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  �Q  ��  ��    l     ��������  ��  ��    l ������ I ������
�� .sysottosnull���     TEXT m  �� �  F i n i s h e d .��  ��  ��     l     ��������  ��  ��    !"! i   " %#$# I      ��%���� 0 simple_sort  % &��& o      ���� 0 my_list  ��  ��  $ k     u'' ()( r     *+* J     ����  + l     ,����, o      ���� 0 
index_list  ��  ��  ) -.- r    	/0/ J    ����  0 l     1����1 o      ���� 0 sorted_list  ��  ��  . 232 U   
 r454 k    m66 787 r    9:9 m    ;; �<<  : l     =����= o      ���� 0 low_item  ��  ��  8 >?> Y    c@��AB��@ Z   ) ^CD����C H   ) -EE E  ) ,FGF l  ) *H����H o   ) *���� 0 
index_list  ��  ��  G o   * +���� 0 i  D k   0 ZII JKJ r   0 8LML c   0 6NON n   0 4PQP 4   1 4��R
�� 
cobjR o   2 3���� 0 i  Q o   0 1���� 0 my_list  O m   4 5��
�� 
ctxtM o      ���� 0 	this_item  K S��S Z   9 ZTUV��T =  9 <WXW l  9 :Y����Y o   9 :���� 0 low_item  ��  ��  X m   : ;ZZ �[[  U k   ? F\\ ]^] r   ? B_`_ o   ? @���� 0 	this_item  ` l     a����a o      ���� 0 low_item  ��  ��  ^ b��b r   C Fcdc o   C D���� 0 i  d l     e����e o      ���� 0 low_item_index  ��  ��  ��  V fgf A I Lhih o   I J���� 0 	this_item  i l  J Kj����j o   J K���� 0 low_item  ��  ��  g k��k k   O Vll mnm r   O Ropo o   O P���� 0 	this_item  p l     q����q o      ���� 0 low_item  ��  ��  n r��r r   S Vsts o   S T���� 0 i  t l     u����u o      ���� 0 low_item_index  ��  ��  ��  ��  ��  ��  ��  ��  �� 0 i  A m    ���� B l   $v����v n    $wxw m   ! #��
�� 
nmbrx n   !yzy 2   !�
� 
cobjz o    �~�~ 0 my_list  ��  ��  ��  ? {|{ r   d h}~} l  d e�}�| o   d e�{�{ 0 low_item  �}  �|  ~ l     ��z�y� n      ���  ;   f g� o   e f�x�x 0 sorted_list  �z  �y  | ��w� r   i m��� l  i j��v�u� o   i j�t�t 0 low_item_index  �v  �u  � l     ��s�r� n      ���  ;   k l� l  j k��q�p� o   j k�o�o 0 
index_list  �q  �p  �s  �r  �w  5 l   ��n�m� l   ��l�k� n    ��� m    �j
�j 
nmbr� n   ��� 2   �i
�i 
cobj� o    �h�h 0 my_list  �l  �k  �n  �m  3 ��g� L   s u�� l  s t��f�e� o   s t�d�d 0 sorted_list  �f  �e  �g  " ��� l     �c�b�a�c  �b  �a  � ��`� l     �_�^�]�_  �^  �]  �`       �\�����\  � �[�Z�Y
�[ 
pimr�Z 0 simple_sort  
�Y .aevtoappnull  �   � ****� �X��X �  ���� �W �V
�W 
vers�V  � �U��T
�U 
cobj� ��   �S
�S 
osax�T  � �R��
�R 
cobj� ��   �Q 
�Q 
scpt� �P �O
�P 
vers�O  � �N$�M�L���K�N 0 simple_sort  �M �J��J �  �I�I 0 my_list  �L  � �H�G�F�E�D�C�B�H 0 my_list  �G 0 
index_list  �F 0 sorted_list  �E 0 low_item  �D 0 i  �C 0 	this_item  �B 0 low_item_index  � �A�@;�?Z
�A 
cobj
�@ 
nmbr
�? 
ctxt�K vjvE�OjvE�O g��-�,Ekh�E�O Hk��-�,Ekh �� /��/�&E�O��  �E�O�E�Y �� �E�O�E�Y hY h[OY��O��6FO��6F[OY��O�� �>��=�<���;
�> .aevtoappnull  �   � ****� k    ���  *��  0��  8��  >��  P��  Z��  `��  ���  ���  ���  ���  ���  ��� �� �� /�� 9�� |�� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� �� 	�� �� �� �� �� _�� d�� ��� ��� �� �� ��� ��� ��� ��� E�� O�� �� �:�:  �=  �<  � �9�8�7�6�5�4�3�9 0 i  �8 0 j  �7 
0 lesson  �6 0 room  �5 0 theevent theEvent�4 0 	acalendar 	aCalendar�3 0 anevent anEvent�P .�2 4�1 < H K�0 N�/�.�- ^�,�+ m�*�)�(�'�&�%�$ ��#�"�!� � � � � � � �� � �'*-�@FSm����������������������
�	��(�A������� o�������������������� ����Gbx��������������������(4K������|�����������������&4M[t������������_���������������������/!��-��<������`���������������������������������� ������ �������� ������������ ����L���� ��d  ������  ����������������������	�����������	��	/
d	H��������������������������	�	�������	���	�	�	�����
,
v��
�
�
�����������
�/��w�������������~�}�|� �{ET�z�y�x��w�v�u�t�s�r�q��p�o�n�m��l���k���j���i
�2 .sysottosnull���     TEXT
�1 .sysodlogaskr        TEXT
�0 
prmp
�/ .gtqpchltns    @   @ ns  �. &0 chosencalendarapp chosenCalendarApp
�- 
cobj
�, .misccurdldt    ��� null�+  0 thecurrentdate theCurrentDate
�* 
dtxt
�) 
dstr
�( 
spac
�' 
tstr
�& 
rslt
�% 
ttxt�$ 0 thetext theText
�# 
ldt �" 0 firstday firstDay�!  �   
� .sysobeepnull��� ��� long� 0 thedate theDate�  0 delayedyesorno delayedYesOrNo� "0 delayedfirstday delayedFirstDay
� 
days
� 
year� 0 a  
� 
mnth
� 
nmbr� 0 b  
� 
day � 0 c  � 0 e  � 0 enddate endDate� 0 thelist theList
� 
txdl� 
0 tid TID
� 
ctxt� "0 thelistasstring theListAsString� "0 endofrecurrence endOfRecurrence� 0 
classnames 
classNames
�
 .miscactvnull��� ��� null
�	 
docu
� 
pnam� "0 listofdocuments listOfDocuments�  0 chosendocument chosenDocument
� 
NmSh� 0 listofsheets listOfSheets� 0 chosensheet chosenSheet
� 
NmTb
� 
NmCl�  
�  
NMCv�� 0 yearonebegin yearOneBegin
�� 
NMCo
�� 
NMad�� ,0 yearonecolumnaddress yearOneColumnAddress
�� .corecnte****       ****�� (0 yearonecolumncount yearOneColumnCount
�� 
NMRw
�� 
msng�� �� 0 simple_sort  �� 0 
sortedlist 
sortedList�� 0 	roomindex 	roomIndex�� 0 columnnumber columnNumber�� 0 	rownumber 	rowNumber�� 0 rowcount rowCount�� 0 roomnumbers roomNumbers�� 0 lessons  �� �� "�� 0 lessons3  
�� 
kocl�� 0 lessons4  �� 0 	inputtime 	inputTime
�� 
cwor�� �� 0 thetime theTime�� �� 0 thetime2 theTime2�� 0 lessons5  �� 0 lessons6  �� 0 lessons7  �� 0 
datestring 
dateString�� 0 lessons8  �� 0 lessons9  �� 0 teacherfirst teacherFirst�� 0 teachersecond teacherSecond�� 0 yearword yearWord�� 0 
yearnumber 
yearNumber�� 0 	classname 	className�� 0 	eventname 	eventName�� 0 teacher  
�� 
wres
�� .coredelonull���     obj 
�� .coredoexnull���     ****
�� 
wtnm
�� .wrbtaec2null��� ��� null
�� 
wrev
�� 
prdt
�� 
wr11
�� 
wr14
�� 
wr1s
�� 
wr5s
�� 
min �� P
�� 
wr15�� 

�� .corecrel****      � null�� 0 problem  �� "0 neweventsummary newEventSummary�� 0 thecalendars theCalendars
�� 
scpt
�� .!Cls!fstnull��� ��� null�� 0 thestore theStore�� 0 	startdate 	startDate
�� 
!CtY
�� !Tct!TtL
�� 
!Cst
�� .!CLs!fcAnull���     ****
�� 
!Csd
�� 
!Ced
�� 
!Csc�� 
�� .!CLs!feSnull��� ��� null�� 0 	theevents 	theEvents
�� 
!Cev
�� 
!Ctz
�� .!CLs!mo2null��� ��� null�� 0 thenewevent theNewEvent
�� .!CLs!updnull��� ��� null
�� !Tct!TtC�� 0 startingdate startingDate�� 0 
endingdate 
endingDate
�� 
!Cft�� 
�� .!CLs!reEnull��� ��� null
�� .sysodelanull��� ��� nmbr
�� .aevtquitnull��� ��� null
�� 
prcs
�� 
cwin
�� 
butT
�� .prcsclicnull��� ��� uiel
�� 
sgrp
�� 
scra
�� 
tabB
�� 
crow
�� 
uiel�� Z
�� 
faal
�� eMdsKcmd
�� .prcskprsnull���     ctxt
�� 
tab �� ~
�� eMdsKopt
�� .prcskcodnull���     ****�� 0 calendarnames calendarNames
�� 
splg
�� 
outl
�� 
txtf
�� 
valL�� $0 calendarnametext calendarNameText�� }
�� 
mbar
�� 
mbri
�� 
menE
�� 
popv
�� 
chbx�� 0 calendarurls calendarURLs
�� 
sheE�� 0 calendarname calendarName
�� 
txta�� 0 theurl theURL
�� 
ret 
�� 
cCFo
�� 
cAct
�� 
emad�� 0 wintecaccount wintecAccount
�� 
cEvt
�� 
subj
� 
iloc
�~ 
offs
�} 
endT
�| 
pHsR�{ 0 	allevents 	allEvents�z "0 winteccalendars wintecCalendars�y 0 startday startDay
�x 
wkdy
�w 
ReDw
�v 
rDMo
�u 
ReED
�t 
pEnT
�s eEDTeEDt
�r 
peDa
�q 
tStr
�p 
ReOi
�o 
ReTy
�n eRPTeRwp
�m 
rREc
�l 
rDTu
�k 
rDWe
�j 
rDTh
�i 
rDFr�;��j O�j O�j O��lv��l 	E�O��k/E�O�j O ahZ*j E�O�a �a ,_ %�a ,%l O_ a ,E` O !_ a  *a _ /E` OY hW X  *j [OY��Oa _ %a %j Oa a _ a ,_ %_ a ,%l Oa  j O chZ*j E�Oa !a �a ,_ %�a ,%l O_ a ,E` O !_ a " *a _ /E` #OY hW X  *j [OY��Oa $_ #%a %%j Oa &a _ #a ,_ %_ #a ,%l Oa 'j Oa (a )lv�a *l 	E` +O_ +�k/E` +O_ +a ,  qa -j O chZ*j E�Oa .a �a ,_ %�a ,%l O_ a ,E` O !_ a / *a _ /E` 0OY hW X  *j [OY��Y hO_ #_ 1k E` #O_ #a 2,E` 3O_ #a 4,a 5&E` 6O_ #a 7,E` 8Oa 9E` :O_ #E` ;O_ 3_ 6_ 8mvE` <O*a =,a >lvE[�k/E` ?Z[�l/*a =,FZO_ <a @&E` AO_ ?*a =,FO_ :_ AlvE` <O*a =,a BlvE[�k/E` ?Z[�l/*a =,FZO_ <a @&E` CO_ ?*a =,FOjvE` DOa E�*j FO*a G-a H,E` IO) 	a Jj UO_ Ij 	E` KO_ K�k/E` KO*a G_ K/c) 	a Lj UO*a M-a H,E` NO_ Nj 	E` OO_ O�k/E` OO*a M_ O/�*a Pk/~*a Qk/a R[a S,\Za T@1E` UO_ Ua V,a W,E` XO*a V_ X/a Q-j YE` ZO_ Ua [,a W,kE` UO `_ U_ Zkh  *a V_ X/a Q�/a S,a \  Y 2*a V_ Xa ]/a Q�/a S,a ^%*a V_ X/a Q�/a S,%_ D6FOP[OY��O*a Ql/a R[a S,\Za _@1E` UO_ Ua V,a W,E` XO*a V_ X/a Q-j YE` ZO_ Ua [,a W,kE` UO `_ U_ Zkh  *a V_ X/a Q�/a S,a \  Y 2*a V_ Xa ]/a Q�/a S,a `%*a V_ X/a Q�/a S,%_ D6FOP[OY��O)_ Dk+ aE` bO_ bE` DOPUOPUO) 	a cj UO*a M-a H,E` NO_ Nj 	E` OO_ O�k/E` OO) 	a dj UO*a M_ O/Q*a Pk/G*a Qk/a R[a S,\Za e81E` fO_ fa V,a W,E` gO_ fa [,a W,E` hO*a V_ g/a Q-j YE` iOjvE` jO c_ hk_ ikh  *a V_ gk/a Q�/a S,a \ 4*a V_ gk/a Q�/a S,*a V_ g/a Q�/a S,a @&lv_ j6FY hOP[OY��OjvE` kO �la lkh   qa ]a mkh *a V�/a Q�/a S,a \ J*a V�/a Ql/a S,*a V�/a Qm/a S,*a Vk/a Q�/a S,*a V�/a Q�/a S,a ]v_ k6FY h[OY��[OY��UUUUOjvE` nO Z_ k[a o�l Ykh ��l/a p ��k/a q��m/��a ]/a ]v_ n6FY ��k/a r��m/��a ]/a ]v_ n6F[OY��OjvE` sO �_ n[a o�l Ykh ��m/E` tO_ ta ua ]/a v%_ ta ua w/%E` xO_ xa uk/a 5&a y *_ xa uk/a ya z%_ xa ul/%a {%a @&E` |Y #_ xa uk/a }%_ xa ul/%a ~%a @&E` |O��k/��l/_ |��a ]/a ]v_ s6F[OY�\OjvE` O 6_ s[a o�l Ykh ��l/a �%��a ]/%��k/��m/mv_ 6F[OY��OjvE` �O _ [a o�l Ykh ��l/a �  ��k/_ a ,a �%��m/%lv_ �6FY ���l/a � &��k/_ _ 1k a ,a �%��m/%lv_ �6FY ���l/a � &��k/_ _ 1l a ,a �%��m/%lv_ �6FY c��l/a � &��k/_ _ 1m a ,a �%��m/%lv_ �6FY 4��l/a � (��k/_ _ 1a ] a ,a �%��m/%lv_ �6FY h[OY�OjvE` �O <_ �[a o�l Ykh ��l/E` �O*a _ �/E` #O��k/_ #lv_ �6F[OY��OjvE` �O t_ �[a o�l Ykh ��k/a � ��k/��l/a �mv_ �6FY > ;_ j[a o�l Ykh ��k/��k/ ��k/��l/��l/mv_ �6FY h[OY��[OY��OjvE` �O_ �[a o�l Ykh ��k/a � v��k/a ua ]/E` �O��k/a ua w/E` �O��k/a uk/E` �O��k/a ul/E` �O��k/a um/E` �O_ �a �%_ �%a �%_ �%a �%_ �%a �%_ �%E` �Y [��k/a ua ]/E` �O��k/a uk/E` �O��k/a ul/E` �O��k/a um/E` �O_ �a �%_ �%a �%_ �%a �%_ �%E` �O_ ���l/��m/mv_ �6F[OY�O�a � %a � 2*j FO) 	a �j UO*a �-a R[a H,\Za �@1 *j �UUOa �j Oa �O*j FO :k_ Dj Ykh  *a �_ D�/E/j � *a �_ D�/l �Y h[OY��O) 	a �j UO �_ �[a o�l Ykh *a ���k/E/j � Q*a ���k/E/ @*a oa �a �a ���k/a ���m/a ���l/a ���l/_ �a � a �_ Ca �a ] �UY ���k/a u-E` �O_ ��k/a �%_ ��l/%a �%_ ��m/%a �%_ ��a ]/%E` �O*a �_ �/ @*a oa �a �a ���k/a ���m/a ���l/a ���l/_ �a � a �_ Ca �a ] �U[OY�UO) 	a �j UOa � *a �-a H,a R[a H,\Za �@1E` �UO)a �a �/ *j �UE` �O_ E` �O_ �_ 1a w E` ;O)a �a �/ _ �a �a �a �_ �a ] �UE` �O)a �a �/ !*a �_ �a �_ ;a �_ �a �_ �a � �UE` �O Zk_ �j Ykh  )a �a �/ *a �_ ��/a �a �a ] �UE` �O)a �a �/ *a �_ �a �_ �a ] �U[OY��O_ +a �  �)a �a �/ *j �UE` �O)a �a �/ jva �a �a �_ �a ] �UE` �O_ E` �O_ 0_ 1k E` �O)a �a �/ !*a �_ �a �_ �a �_ �a �_ �a � �UE` �O 9_ �[a o�l Ykh )a �a �/ *a Ѥa �_ �a �fa � �U[OY��OPY hOa �j Oa � *j FOmj �O*j �O*j FOlj �UOa � -*a �a �/ !*a �k/ *a �k/ *j �Omj �UUUUOa � �*a �a �/ }*a �k/ s*a �k/ i*a �k/ _*a �k/ U*a �a w/ I*a �k/ ?*a �k/ 5) 	a �j UO*j �Oa wj �O) 	a �j UO*j �Okj �UUUUUUUUUOa � *j FOa �j �UOa �j Oa ��*a �a �/�a �a �a �l �Okj �O_ �j �Okj �Oa �a �a �kvl �Okj �OjvE` Onl*a �k/ak/ak/a �k/ak/a �-j Ykh  6*a �k/ak/ak/a �k/ak/a �/a �k/ak/a,E`O_a �__ 6FOaa �a �l �Okj �O_ �j �Oaj �Okj �O*a	k/a
a/aa/ *j �UO*a	k/a
a/aa/a �a l/ *j �UO*a �k/ak/ak/a �k/ak/a �/ak/ak/ *j �UO*a �k/ak/ak/a �k/ak/a �/ak/a �a/ *j �UOkj �O_ �j �Okj �Y hW X  hOP[OY��UUOa � *j FUOaj OjvE`Oa � �*a �a/ �aa �a �l �Okj �O_ �j �Ojj �Oa �a �a �kvl �Ojj �O �k_ j Ykh  aj �Ojj �Oaa �a �l �Ojj �O*a �k/ 8*ak/ .*ak/a,E`O*a �k/ *ak/ *a,E`UUUUO_j �Ojj �O__lv_6F[OY�|UUOaj Oa E �*j FO*a oa Gl �O*a Gk/ �*a Mk/ �*a Pk/ �*a [k/ a*a Qk/a S,FOa*a Ql/a S,FUO {k_j Ykh  _�/�k/E`O_�/�l/E`O*a [�k/j � *a oa [l �Y hO*a [�k/ _*a Qk/a S,FO_*a Ql/a S,FUOP[OY��UUUUOPY�a �*j FO*a!-a R[a H,\Za"@1j �O*a!k/a R[a#-a$,\Za%@1E`&O_&G Gk_ Dj Ykh  *a!_ D�/E/j �  *a oa!a �a H_ D�/la ] �Y h[OY��O �_ �[a o�l Ykh *a!��k/E/j � O*a!��k/E/ >*a oa'a �a(��k/a)��m/a*��l/a+��l/_ �a � a,fa �a ] �UY ���k/a u-E` �O_ ��k/a-%_ ��l/%a.%_ ��m/%a/%_ ��a ]/%E` �O*a!_ �/ >*a oa'a �a(��k/a)��m/a*��l/a+��l/_ �a � a,fa �a ] �U[OY�OPUUOa jvE`0O*a!k/a R[a#-a$,\Za1@1E`&O_& ]*a!-a R[a H,\Za2@1E`3O A_3[a o�l Ykh �a'-E` �O _ �[a o�l Ykh �_06F[OY��[OY��UO�_0[a o�l Ykh �a(,EO�a*,E`4O_4a5,a @&E`4O_4a6  ;a7a8ela9a:a;a<_ ;a ]a=a>a?ka@aAa ��aB,FY_4aC  ;a7aDela9a:a;a<_ ;a ]a=aEa?ka@aAa ��aB,FY �_4aF  ;a7aGela9a:a;a<_ ;a ]a=aHa?ka@aAa ��aB,FY �_4aI  ;a7aJela9a:a;a<_ ;a ]a=aKa?ka@aAa ��aB,FY F_4aL  ;a7aMela9a:a;a<_ ;a ]a=aNa?ka@aAa ��aB,FY hOP[OY��UOaOj ascr  ��ޭ