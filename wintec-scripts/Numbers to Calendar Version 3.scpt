FasdUAS 1.101.10   ��   ��    k             l        	  x     �� 
 ��   
 1      ��
�� 
ascr  �� ��
�� 
minv  m         �    2 . 4��       Yosemite (10.10) or later    	 �   4   Y o s e m i t e   ( 1 0 . 1 0 )   o r   l a t e r      x    �� ����    2  	 ��
�� 
osax��        l          x    "��  ��    4  ai�� 
�� 
scpt  m  eh   �    C a l e n d a r L i b   E C  �� ��
�� 
minv  m         �   
 1 . 1 . 5��    * $ put this at the top of your scripts     �     H   p u t   t h i s   a t   t h e   t o p   o f   y o u r   s c r i p t s   ! " ! l     ��������  ��  ��   "  # $ # l     �� % &��   % � � This script converts an Excel timetable to Apple Calendar. Must open the excel timetable in Numbers first, then run this script    & � ' '    T h i s   s c r i p t   c o n v e r t s   a n   E x c e l   t i m e t a b l e   t o   A p p l e   C a l e n d a r .   M u s t   o p e n   t h e   e x c e l   t i m e t a b l e   i n   N u m b e r s   f i r s t ,   t h e n   r u n   t h i s   s c r i p t $  ( ) ( l     ��������  ��  ��   )  * + * l     ,���� , I    �� -��
�� .sysottosnull���     TEXT - m      . . � / / � T h i s   s c r i p t   w i l l   f i r s t   d e l e t e   a n y   c a l e n d a r s   t h a t   c o n t a i n   t h e   w o r d   ' Y e a r '   i n   t h e   c a l e n d a r   n a m e .��  ��  ��   +  0 1 0 l    2���� 2 I   �� 3��
�� .sysodlogaskr        TEXT 3 m     4 4 � 5 5 � T h i s   s c r i p t   w i l l   f i r s t   d e l e t e   a n y   c a l e n d a r s   t h a t   c o n t a i n   t h e   w o r d   ' Y e a r '   i n   t h e   c a l e n d a r   n a m e .��  ��  ��   1  6 7 6 l     ��������  ��  ��   7  8 9 8 l     �� : ;��   : h bsay "Which calendar app do you want to create the events in? Apple Calendar or Microsoft Outlook?"    ; � < < � s a y   " W h i c h   c a l e n d a r   a p p   d o   y o u   w a n t   t o   c r e a t e   t h e   e v e n t s   i n ?   A p p l e   C a l e n d a r   o r   M i c r o s o f t   O u t l o o k ? " 9  = > = l     �� ? @��   ? � �set chosenCalendarApp to choose from list {"Apple Calendar", "Microsoft Outlook"} with prompt "Which calendar app do you want the calendars created in?"    @ � A A0 s e t   c h o s e n C a l e n d a r A p p   t o   c h o o s e   f r o m   l i s t   { " A p p l e   C a l e n d a r " ,   " M i c r o s o f t   O u t l o o k " }   w i t h   p r o m p t   " W h i c h   c a l e n d a r   a p p   d o   y o u   w a n t   t h e   c a l e n d a r s   c r e a t e d   i n ? " >  B C B l     �� D E��   D : 4set chosenCalendarApp to item 1 of chosenCalendarApp    E � F F h s e t   c h o s e n C a l e n d a r A p p   t o   i t e m   1   o f   c h o s e n C a l e n d a r A p p C  G H G l     ��������  ��  ��   H  I J I l    K���� K r     L M L m     N N � O O  A p p l e   C a l e n d a r M o      ���� &0 chosencalendarapp chosenCalendarApp��  ��   J  P Q P l     ��������  ��  ��   Q  R S R l    T���� T I   �� U��
�� .sysottosnull���     TEXT U m     V V � W W � W h e n   i s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r ?   ( S h o u l d   b e   o n   a   M o n d a y ,   y o u   c a n   e n t e r   a   d e l a y e d   s t a r t   d a y   l a t e r )��  ��  ��   S  X Y X l   l Z���� Z T    l [ [ k    g \ \  ] ^ ] r    " _ ` _ I    ������
�� .misccurdldt    ��� null��  ��   ` o      ����  0 thecurrentdate theCurrentDate ^  a b a I  # 2�� c d
�� .sysodlogaskr        TEXT c m   # $ e e � f f � W h e n   i s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r ?   ( S h o u l d   b e   o n   a   M o n d a y ,   y o u   c a n   e n t e r   a   d e l a y e d   s t a r t   d a y   l a t e r ) d �� g��
�� 
dtxt g l  % . h���� h b   % . i j i b   % * k l k n   % ( m n m 1   & (��
�� 
dstr n o   % &����  0 thecurrentdate theCurrentDate l 1   ( )��
�� 
spac j n   * - o p o 1   + -��
�� 
tstr p o   * +����  0 thecurrentdate theCurrentDate��  ��  ��   b  q r q r   3 : s t s n   3 6 u v u 1   4 6��
�� 
ttxt v 1   3 4��
�� 
rslt t o      ���� 0 thetext theText r  w�� w Q   ; g x y z x Z   > Z { |���� { >  > E } ~ } o   > A���� 0 thetext theText ~ m   A D   � � �   | k   H V � �  � � � l  H T � � � � r   H T � � � 4   H P�� �
�� 
ldt  � o   L O���� 0 thetext theText � o      ���� 0 firstday firstDay �   a date object    � � � �    a   d a t e   o b j e c t �  ��� �  S   U V��  ��  ��   y R      ������
�� .ascrerr ****      � ****��  ��   z I  b g������
�� .sysobeepnull��� ��� long��  ��  ��  ��  ��   Y  � � � l     ��������  ��  ��   �  � � � l  m | ����� � I  m |�� ���
�� .sysottosnull���     TEXT � b   m x � � � b   m t � � � m   m p � � � � � D T h e   f i r s t   d a y   o f   t h e   s e m e s t e r   i s :   � o   p s���� 0 firstday firstDay � m   t w � � � � �  P l e a s e   C o n f i r m��  ��  ��   �  � � � l  } � ����� � I  } ��� � �
�� .sysodlogaskr        TEXT � m   } � � � � � � D T h e   f i r s t   d a y   o f   t h e   s e m e s t e r   i s :   � �� ���
�� 
dtxt � l  � � ����� � b   � � � � � b   � � � � � n   � � � � � 1   � ���
�� 
dstr � o   � ����� 0 firstday firstDay � 1   � ���
�� 
spac � n   � � � � � 1   � ���
�� 
tstr � o   � ����� 0 firstday firstDay��  ��  ��  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l  � � ����� � I  � ��� ���
�� .sysottosnull���     TEXT � m   � � � � � � � J W h e n   i s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r ?��  ��  ��   �  � � � l  � � ����� � T   � � � � k   � � � �  � � � r   � � � � � I  � �������
�� .misccurdldt    ��� null��  ��   � o      ����  0 thecurrentdate theCurrentDate �  � � � I  � ��� � �
�� .sysodlogaskr        TEXT � m   � � � � � � � J W h e n   i s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r ? � �� ���
�� 
dtxt � l  � � ����� � b   � � � � � b   � � � � � n   � � � � � 1   � ���
�� 
dstr � o   � �����  0 thecurrentdate theCurrentDate � 1   � ���
�� 
spac � n   � � � � � 1   � ���
�� 
tstr � o   � �����  0 thecurrentdate theCurrentDate��  ��  ��   �  � � � r   � � � � � n   � � � � � 1   � ���
�� 
ttxt � 1   � ���
�� 
rslt � o      ���� 0 thetext theText �  ��� � Q   � � � � � � Z   � � � ����� � >  � � � � � o   � ��� 0 thetext theText � m   � � � � � � �   � k   � � � �  � � � l  � � � � � � r   � � � � � 4   � ��~ �
�~ 
ldt  � o   � ��}�} 0 thetext theText � o      �|�| 0 thedate theDate �   a date object    � � � �    a   d a t e   o b j e c t �  ��{ �  S   � ��{  ��  ��   � R      �z�y�x
�z .ascrerr ****      � ****�y  �x   � I  � ��w�v�u
�w .sysobeepnull��� ��� long�v  �u  ��  ��  ��   �  � � � l     �t�s�r�t  �s  �r   �  � � � l  � ��q�p � I  ��o ��n
�o .sysottosnull���     TEXT � b   � � � � � b   � � � � � m   � � � � � � � B T h e   l a s t   d a y   o f   t h e   s e m e s t e r   i s :   � o   � ��m�m 0 thedate theDate � m   � � � � � � �  P l e a s e   C o n f i r m .�n  �q  �p   �  � � � l  ��l�k � I �j � �
�j .sysodlogaskr        TEXT � m   � � � � � B T h e   l a s t   d a y   o f   t h e   s e m e s t e r   i s :   � �i ��h
�i 
dtxt � l  ��g�f � b   � � � b     n   1  �e
�e 
dstr o  �d�d 0 thedate theDate 1  �c
�c 
spac � n   1  �b
�b 
tstr o  �a�a 0 thedate theDate�g  �f  �h  �l  �k   �  l     �`�_�^�`  �_  �^   	 l     �]�\�[�]  �\  �[  	 

 l     �Z�Z     Ask for delayed start    � ,   A s k   f o r   d e l a y e d   s t a r t  l !�Y�X I !�W�V
�W .sysottosnull���     TEXT m   � R I s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r   d e l a y e d ?�V  �Y  �X    l "8�U�T r  "8 I "4�S
�S .gtqpchltns    @   @ ns   J  "*  m  "% �    Y e s !�R! m  %("" �##  N o�R   �Q$�P
�Q 
prmp$ m  -0%% �&& R I s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r   d e l a y e d ?�P   o      �O�O  0 delayedyesorno delayedYesOrNo�U  �T   '(' l 9E)�N�M) r  9E*+* n  9A,-, 4  <A�L.
�L 
cobj. m  ?@�K�K - o  9<�J�J  0 delayedyesorno delayedYesOrNo+ o      �I�I  0 delayedyesorno delayedYesOrNo�N  �M  ( /0/ l     �H�G�F�H  �G  �F  0 121 l F�3�E�D3 Z  F�45�C�B4 = FM676 o  FI�A�A  0 delayedyesorno delayedYesOrNo7 m  IL88 �99  Y e s5 k  P�:: ;<; I PW�@=�?
�@ .sysottosnull���     TEXT= m  PS>> �?? � I f   t h e   f i r s t   d a y   i s   d e l a y e d   a n d   n o t   o n   a   M o n d a y ,   w h e n   i s   t h e   d e l a y e d   f i r s t   d a y ?�?  < @�>@ T  X�AA k  ]�BB CDC r  ]dEFE I ]b�=�<�;
�= .misccurdldt    ��� null�<  �;  F o      �:�:  0 thecurrentdate theCurrentDateD GHG I ev�9IJ
�9 .sysodlogaskr        TEXTI m  ehKK �LL � I f   t h e   f i r s t   d a y   i s   d e l a y e d   a n d   n o t   o n   a   M o n d a y ,   w h e n   i s   t h e   d e l a y e d   f i r s t   d a y ?J �8M�7
�8 
dtxtM l irN�6�5N b  irOPO b  inQRQ n  ilSTS 1  jl�4
�4 
dstrT o  ij�3�3  0 thecurrentdate theCurrentDateR 1  lm�2
�2 
spacP n  nqUVU 1  oq�1
�1 
tstrV o  no�0�0  0 thecurrentdate theCurrentDate�6  �5  �7  H WXW r  w~YZY n  wz[\[ 1  xz�/
�/ 
ttxt\ 1  wx�.
�. 
rsltZ o      �-�- 0 thetext theTextX ]�,] Q  �^_`^ Z  ��ab�+�*a > ��cdc o  ���)�) 0 thetext theTextd m  ��ee �ff  b k  ��gg hih l ��jklj r  ��mnm 4  ���(o
�( 
ldt o o  ���'�' 0 thetext theTextn o      �&�& "0 delayedfirstday delayedFirstDayk   a date object   l �pp    a   d a t e   o b j e c ti q�%q  S  ���%  �+  �*  _ R      �$�#�"
�$ .ascrerr ****      � ****�#  �"  ` I ���!� �
�! .sysobeepnull��� ��� long�   �  �,  �>  �C  �B  �E  �D  2 rsr l     ����  �  �  s tut l ��v��v r  ��wxw [  ��yzy o  ���� 0 thedate theDatez l ��{��{ ]  ��|}| 1  ���
� 
days} m  ���� �  �  x o      �� 0 thedate theDate�  �  u ~~ l     ����  �  �   ��� l ������ r  ����� l ������ n  ����� 1  ���
� 
year� o  ���� 0 thedate theDate�  �  � o      �
�
 0 a  �  �  � ��� l ����	�� r  ����� c  ����� l ������ n  ����� m  ���
� 
mnth� o  ���� 0 thedate theDate�  �  � m  ���
� 
nmbr� o      �� 0 b  �	  �  � ��� l ����� � r  ����� n  ����� 1  ����
�� 
day � o  ������ 0 thedate theDate� o      ���� 0 c  �  �   � ��� l     ������  �  set d to "T000000Z"   � ��� & s e t   d   t o   " T 0 0 0 0 0 0 Z "� ��� l �������� r  ����� m  ���� ��� H F R E Q = W E E K L Y ; I N T E R V A L = 1 ; B Y D A Y = ; U N T I L =� o      ���� 0 e  ��  ��  � ��� l     ��������  ��  ��  � ��� l ������ r  ����� o  ������ 0 thedate theDate� o      ���� 0 enddate endDate� R L this variable represents the last day of the semester for Microsoft Outlook   � ��� �   t h i s   v a r i a b l e   r e p r e s e n t s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r   f o r   M i c r o s o f t   O u t l o o k� ��� l     ��������  ��  ��  � ��� l �
������ r  �
��� J  ��� ��� o  ������ 0 a  � ��� o  ����� 0 b  � ���� o  ���� 0 c  ��  � o      ���� 0 thelist theList��  ��  � ��� l     ��������  ��  ��  � ��� l .������ r  .��� J  �� ��� 1  ��
�� 
txdl� ���� m  �� ���  0��  � J      �� ��� o      ���� 
0 tid TID� ���� 1  ',��
�� 
txdl��  ��  ��  � ��� l /:������ r  /:��� c  /6��� o  /2���� 0 thelist theList� m  25��
�� 
ctxt� o      ���� "0 thelistasstring theListAsString��  ��  � ��� l ;D������ r  ;D��� o  ;>���� 
0 tid TID� 1  >C��
�� 
txdl��  ��  � ��� l     ��������  ��  ��  � ��� l EQ������ r  EQ��� J  EM�� ��� o  EH���� 0 e  � ���� o  HK���� "0 thelistasstring theListAsString��  � o      ���� 0 thelist theList��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  � Y S the variable endOfRecurrence represents the last day of classes for Apple Calendar   � ��� �   t h e   v a r i a b l e   e n d O f R e c u r r e n c e   r e p r e s e n t s   t h e   l a s t   d a y   o f   c l a s s e s   f o r   A p p l e   C a l e n d a r� ��� l Ru������ r  Ru��� J  R\�� ��� 1  RW��
�� 
txdl� ���� m  WZ�� ���  ��  � J      �� ��� o      ���� 
0 tid TID� ���� 1  ns��
�� 
txdl��  ��  ��  � ��� l v������� r  v���� c  v}� � o  vy���� 0 thelist theList  m  y|��
�� 
ctxt� o      ���� "0 endofrecurrence endOfRecurrence��  ��  �  l ������ r  �� o  ������ 
0 tid TID 1  ����
�� 
txdl��  ��    l     ��������  ��  ��   	 l ��
����
 r  �� J  ������   o      ���� 0 
classnames 
classNames��  ��  	  l �E���� O  �E k  �D  I ��������
�� .miscactvnull��� ��� null��  ��    r  �� n  �� 1  ����
�� 
pnam 2 ����
�� 
docu o      ���� "0 listofdocuments listOfDocuments  O �� I ������
�� .sysottosnull���     TEXT m  ��   �!! � C h o o s e   w h i c h   N u m b e r s   d o c u m e n t   h a s   t h e   t i m e t a b l e   a n d   c l a s s   n a m e s .��    f  �� "#" r  ��$%$ I ����&��
�� .gtqpchltns    @   @ ns  & o  ������ "0 listofdocuments listOfDocuments��  % o      ����  0 chosendocument chosenDocument# '(' r  ��)*) n  ��+,+ 4  ����-
�� 
cobj- m  ������ , o  ������  0 chosendocument chosenDocument* o      ����  0 chosendocument chosenDocument( ./. l ����������  ��  ��  / 0��0 O  �D121 k  �C33 454 O ��676 I ����8��
�� .sysottosnull���     TEXT8 m  ��99 �:: B C h o o s e   t h e   s h e e t   n a m e d   " C l a s s e s " .��  7  f  ��5 ;<; r  ��=>= n  ��?@? 1  ����
�� 
pnam@ 2 ����
�� 
NmSh> o      ���� 0 listofsheets listOfSheets< ABA r  �CDC I � ��E��
�� .gtqpchltns    @   @ ns  E o  ������ 0 listofsheets listOfSheets��  D o      ���� 0 chosensheet chosenSheetB FGF r  HIH n  JKJ 4  ��L
�� 
cobjL m  ���� K o  ���� 0 chosensheet chosenSheetI o      ���� 0 chosensheet chosenSheetG MNM O  �OPO k  �QQ RSR O  �TUT k  &�VV WXW l &&��YZ��  Y   get year one class names   Z �[[ 2   g e t   y e a r   o n e   c l a s s   n a m e sX \]\ r  &?^_^ 6 &;`a` 4 &,��b
�� 
NmClb m  *+���� a E  /:cdc n  04efe 1  04��
�� 
NMCvf  g  00d m  59gg �hh 
 C l a s s_ o      ���� 0 yearonebegin yearOneBegin] iji r  @Oklk n  @Kmnm 1  GK��
�� 
NMadn n  @Gopo m  CG��
�� 
NMCop o  @C���� 0 yearonebegin yearOneBeginl o      ���� ,0 yearonecolumnaddress yearOneColumnAddressj qrq r  Pdsts I P`��u��
�� .corecnte****       ****u n  P\vwv 2 X\��
�� 
NmClw 4  PX��x
�� 
NMCox o  TW���� ,0 yearonecolumnaddress yearOneColumnAddress��  t o      ���� (0 yearonecolumncount yearOneColumnCountr yzy r  ev{|{ [  er}~} l ep��� n  ep��� 1  lp�~
�~ 
NMad� n  el��� m  hl�}
�} 
NMRw� o  eh�|�| 0 yearonebegin yearOneBegin��  �  ~ m  pq�{�{ | o      �z�z 0 yearonebegin yearOneBeginz ��� l ww�y�x�w�y  �x  �w  � ��� Y  w���v���u� k  ���� ��� Z  �����t�� = ����� n  ����� 1  ���s
�s 
NMCv� n  ����� 4  ���r�
�r 
NmCl� o  ���q�q 0 i  � 4  ���p�
�p 
NMCo� o  ���o�o ,0 yearonecolumnaddress yearOneColumnAddress� m  ���n
�n 
msng�  S  ���t  � k  ���� ��� l ���m���m  � � �						set end of classNames to "Year 1 " & (value of cell i of column yearOneColumnAddress & " " & value of cell i of column (yearOneColumnAddress + 4))   � ���0 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e s   t o   " Y e a r   1   "   &   ( v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s   &   "   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   ( y e a r O n e C o l u m n A d d r e s s   +   4 ) )� ��l� r  ����� b  ����� b  ����� n  ����� 1  ���k
�k 
NMCv� n  ����� 4  ���j�
�j 
NmCl� o  ���i�i 0 i  � 4  ���h�
�h 
NMCo� l ����g�f� [  ����� o  ���e�e ,0 yearonecolumnaddress yearOneColumnAddress� m  ���d�d �g  �f  � m  ���� ���    Y e a r   1  � l ����c�b� n  ����� 1  ���a
�a 
NMCv� n  ����� 4  ���`�
�` 
NmCl� o  ���_�_ 0 i  � 4  ���^�
�^ 
NMCo� o  ���]�] ,0 yearonecolumnaddress yearOneColumnAddress�c  �b  � n      ���  ;  ��� o  ���\�\ 0 
classnames 
classNames�l  � ��[� l ���Z���Z  � 9 3					value of cell i of column yearOneColumnAddress   � ��� f 	 	 	 	 	 v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s�[  �v 0 i  � o  z}�Y�Y 0 yearonebegin yearOneBegin� o  }��X�X (0 yearonecolumncount yearOneColumnCount�u  � ��� l ���W�V�U�W  �V  �U  � ��� l ���T���T  �   get year two class names   � ��� 2   g e t   y e a r   t w o   c l a s s   n a m e s� ��� r  ����� 6 ����� 4 ���S�
�S 
NmCl� m  ���R�R � E  ����� n  ����� 1  ���Q
�Q 
NMCv�  g  ��� m  ���� ��� 
 C l a s s� o      �P�P 0 yearonebegin yearOneBegin� ��� r  ���� n  ����� 1  ���O
�O 
NMad� n  ����� m  ���N
�N 
NMCo� o  ���M�M 0 yearonebegin yearOneBegin� o      �L�L ,0 yearonecolumnaddress yearOneColumnAddress� ��� r  ��� I �K��J
�K .corecnte****       ****� n  ��� 2 �I
�I 
NmCl� 4  �H�
�H 
NMCo� o  
�G�G ,0 yearonecolumnaddress yearOneColumnAddress�J  � o      �F�F (0 yearonecolumncount yearOneColumnCount� ��� r  )��� [  %��� l #��E�D� n  #��� 1  #�C
�C 
NMad� n  ��� m  �B
�B 
NMRw� o  �A�A 0 yearonebegin yearOneBegin�E  �D  � m  #$�@�@ � o      �?�? 0 yearonebegin yearOneBegin� ��� l **�>�=�<�>  �=  �<  � ��� Y  *���;���:� k  8��� ��� Z  8����9�� = 8M��� n  8I��� 1  EI�8
�8 
NMCv� n  8E��� 4  @E�7 
�7 
NmCl  o  CD�6�6 0 i  � 4  8@�5
�5 
NMCo o  <?�4�4 ,0 yearonecolumnaddress yearOneColumnAddress� m  IL�3
�3 
msng�  S  PQ�9  � k  T�  l TT�2�2   � �						set end of classNames to "Year 2 " & (value of cell i of column yearOneColumnAddress & " " & value of cell i of column (yearOneColumnAddress + 4))    �0 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e s   t o   " Y e a r   2   "   &   ( v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s   &   "   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   ( y e a r O n e C o l u m n A d d r e s s   +   4 ) ) �1 r  T�	
	 b  T b  Tm n  Ti 1  ei�0
�0 
NMCv n  Te 4  `e�/
�/ 
NmCl o  cd�.�. 0 i   4  T`�-
�- 
NMCo l X_�,�+ [  X_ o  X[�*�* ,0 yearonecolumnaddress yearOneColumnAddress m  [^�)�) �,  �+   m  il �    Y e a r   2   l m~�(�' n  m~ 1  z~�&
�& 
NMCv n  mz 4  uz�%
�% 
NmCl o  xy�$�$ 0 i   4  mu�# 
�# 
NMCo  o  qt�"�" ,0 yearonecolumnaddress yearOneColumnAddress�(  �'  
 n      !"!  ;  ��" o  ��!�! 0 
classnames 
classNames�1  � #� # l ���$%�  $ 9 3					value of cell i of column yearOneColumnAddress   % �&& f 	 	 	 	 	 v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s�   �; 0 i  � o  -0�� 0 yearonebegin yearOneBegin� o  03�� (0 yearonecolumncount yearOneColumnCount�:  � '(' l ������  �  �  ( )*) r  ��+,+ n ��-.- I  ���/�� 0 simple_sort  / 0�0 l ��1��1 o  ���� 0 
classnames 
classNames�  �  �  �  .  f  ��, o      �� 0 
sortedlist 
sortedList* 232 r  ��454 o  ���� 0 
sortedlist 
sortedList5 o      �� 0 
classnames 
classNames3 6�6 l ������  �  �  �  U 4  #�7
� 
NmTb7 m  !"�� S 8�
8 l ���	���	  �  �  �
  P 4  �9
� 
NmSh9 o  �� 0 chosensheet chosenSheetN :;: O ��<=< I ���>�
� .sysottosnull���     TEXT> m  ��?? �@@ F C h o o s e   t h e   s h e e t   n a m e d   " T i m e t a b l e " .�  =  f  ��; ABA r  ��CDC n  ��EFE 1  ���
� 
pnamF 2 ���
� 
NmShD o      � �  0 listofsheets listOfSheetsB GHG r  ��IJI I ����K��
�� .gtqpchltns    @   @ ns  K o  ������ 0 listofsheets listOfSheets��  J o      ���� 0 chosensheet chosenSheetH LML r  ��NON n  ��PQP 4  ����R
�� 
cobjR m  ������ Q o  ������ 0 chosensheet chosenSheetO o      ���� 0 chosensheet chosenSheetM STS l ����������  ��  ��  T UVU O ��WXW I ����Y��
�� .sysottosnull���     TEXTY m  ��ZZ �[[ V G r a b b i n g   d a t a   f r o m   t h e   n u m b e r s   s p r e a d s h e e t .��  X  f  ��V \��\ O  �C]^] O  �B_`_ k  �Aaa bcb l ����de��  d E ? Find table of room numbers and teachers and put it into a list   e �ff ~   F i n d   t a b l e   o f   r o o m   n u m b e r s   a n d   t e a c h e r s   a n d   p u t   i t   i n t o   a   l i s tc ghg r  �iji 6 �klk 4 ���m
�� 
NmClm m   ���� l = non 1  
��
�� 
NMCvo m  pp �qq  R o o mj o      ���� 0 	roomindex 	roomIndexh rsr r  %tut n  !vwv 1  !��
�� 
NMadw n  xyx m  ��
�� 
NMCoy o  ���� 0 	roomindex 	roomIndexu o      ���� 0 columnnumber columnNumbers z{z r  &5|}| n  &1~~ 1  -1��
�� 
NMad n  &-��� m  )-��
�� 
NMRw� o  &)���� 0 	roomindex 	roomIndex} o      ���� 0 	rownumber 	rowNumber{ ��� r  6J��� I 6F�����
�� .corecnte****       ****� n 6B��� 2 >B��
�� 
NmCl� 4  6>���
�� 
NMCo� o  :=���� 0 columnnumber columnNumber��  � o      ���� 0 rowcount rowCount� ��� r  KQ��� J  KM����  � o      ���� 0 roomnumbers roomNumbers� ��� Y  R��������� k  b��� ��� Z  b�������� > by��� n  bu��� 1  qu��
�� 
NMCv� n  bq��� 4  lq���
�� 
NmCl� o  op���� 0 i  � 4  bl���
�� 
NMCo� l fk������ \  fk��� o  fi���� 0 columnnumber columnNumber� m  ij���� ��  ��  � m  ux��
�� 
msng� r  |���� J  |��� ��� l |������� n  |���� 1  ����
�� 
NMCv� n  |���� 4  �����
�� 
NmCl� o  ������ 0 i  � 4  |����
�� 
NMCo� l �������� \  ����� o  ������ 0 columnnumber columnNumber� m  ������ ��  ��  ��  ��  � ���� c  ����� n  ����� 1  ����
�� 
NMCv� n  ����� 4  �����
�� 
NmCl� o  ������ 0 i  � 4  �����
�� 
NMCo� o  ������ 0 columnnumber columnNumber� m  ����
�� 
ctxt��  � n      ���  ;  ��� o  ������ 0 roomnumbers roomNumbers��  ��  � ���� l ����������  ��  ��  ��  �� 0 i  � [  UZ��� o  UX���� 0 	rownumber 	rowNumber� m  XY���� � o  Z]���� 0 rowcount rowCount��  � ��� l  ��������  ���
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
 	 	 	 	� ��� l ��������  � !  Grab data from spreadsheet   � ��� 6   G r a b   d a t a   f r o m   s p r e a d s h e e t� ��� r  ����� J  ������  � o      ���� 0 lessons  � ���� Y  �A�������� Y  �<�������� Z  �7������� > ����� n  ����� 1  ����
�� 
NMCv� n  ����� 4  �����
�� 
NmCl� o  ������ 0 j  � 4  �����
�� 
NMCo� o  ������ 0 i  � m  ����
�� 
msng� r  �3��� J  �.�� ��� n  ����� 1  ����
�� 
NMCv� n  ����� 4  �����
�� 
NmCl� m  ������ � 4  �����
�� 
NMCo� o  ������ 0 i  � ��� n  ���� 1  ��
�� 
NMCv� n  ���� 4  ���
�� 
NmCl� m  ���� � 4  ����
�� 
NMCo� o  ���� 0 i  � ��� n  ��� 1  ��
�� 
NMCv� n  ��� 4  ���
�� 
NmCl� o  ���� 0 j  � 4  ���
�� 
NMCo� m  ���� � ���� n  *��� 1  &*��
�� 
NMCv� n  &��� 4  !&���
�� 
NmCl� o  $%���� 0 j  � 4  !�� 
�� 
NMCo  o   ���� 0 i  ��  � n        ;  12 o  .1���� 0 lessons  ��  ��  �� 0 j  � m  ������ � m  ������ "��  �� 0 i  � m  ������ � m  ������ ��  ��  ` 4  ����
�� 
NmTb m  ������ ^ 4  ����
�� 
NmSh o  ������ 0 chosensheet chosenSheet��  2 4  ����
�� 
docu o  ������  0 chosendocument chosenDocument��   m  ���                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  ��  ��    l     ��������  ��  ��   	
	 l     ����   7 1 convert year 1 or year 2 from chinese to english    � b   c o n v e r t   y e a r   1   o r   y e a r   2   f r o m   c h i n e s e   t o   e n g l i s h
  l FL���� r  FL J  FH��   o      �~�~ 0 lessons3  ��  ��    l M��}�| X  M��{ Z  c��z E  cm n  ci 4  di�y
�y 
cobj m  gh�x�x  o  cd�w�w 
0 lesson   m  il   �!! Y'N  r  p�"#" J  p�$$ %&% n  pv'(' 4  qv�v)
�v 
cobj) m  tu�u�u ( o  pq�t�t 
0 lesson  & *+* m  vy,, �--  Y e a r   1+ ./. n  y010 4  z�s2
�s 
cobj2 m  }~�r�r 1 o  yz�q�q 
0 lesson  / 3�p3 n  �454 4  ���o6
�o 
cobj6 m  ���n�n 5 o  ��m�m 
0 lesson  �p  # n      787  ;  ��8 o  ���l�l 0 lessons3  �z   r  ��9:9 J  ��;; <=< n  ��>?> 4  ���k@
�k 
cobj@ m  ���j�j ? o  ���i�i 
0 lesson  = ABA m  ��CC �DD  Y e a r   2B EFE n  ��GHG 4  ���hI
�h 
cobjI m  ���g�g H o  ���f�f 
0 lesson  F J�eJ n  ��KLK 4  ���dM
�d 
cobjM m  ���c�c L o  ���b�b 
0 lesson  �e  : n      NON  ;  ��O o  ���a�a 0 lessons3  �{ 
0 lesson   o  PS�`�` 0 lessons  �}  �|   PQP l     �_�^�]�_  �^  �]  Q RSR l     �\TU�\  T   add time to each event   U �VV .   a d d   t i m e   t o   e a c h   e v e n tS WXW l ��Y�[�ZY r  ��Z[Z J  ���Y�Y  [ o      �X�X 0 lessons4  �[  �Z  X \]\ l ��^�W�V^ X  ��_�U`_ k  �|aa bcb r  ��ded n  ��fgf 4  ���Th
�T 
cobjh m  ���S�S g o  ���R�R 
0 lesson  e o      �Q�Q 0 	inputtime 	inputTimec iji r  ��klk b  ��mnm b  ��opo n  ��qrq 4  ���Ps
�P 
cwors m  ���O�O r o  ���N�N 0 	inputtime 	inputTimep m  ��tt �uu  :n n  ��vwv 4  ���Mx
�M 
cworx m  ���L�L w o  ���K�K 0 	inputtime 	inputTimel o      �J�J 0 thetime theTimej yzy Z  �[{|�I}{ ?  �~~ c  ���� n  ���� 4  �H�
�H 
cwor� m  �G�G � l ���F�E� o  ��D�D 0 thetime theTime�F  �E  � m  
�C
�C 
nmbr m  �B�B | r  7��� c  3��� b  /��� b  +��� b  "��� \  ��� l ��A�@� n  ��� 4  �?�
�? 
cwor� m  �>�> � o  �=�= 0 thetime theTime�A  �@  � m  �<�< � m  !�� ���  :� n  "*��� 4  %*�;�
�; 
cwor� m  ()�:�: � o  "%�9�9 0 thetime theTime� m  +.�� ���  P M� m  /2�8
�8 
ctxt� o      �7�7 0 thetime2 theTime2�I  } r  :[��� c  :W��� b  :S��� b  :O��� b  :F��� n  :B��� 4  =B�6�
�6 
cwor� m  @A�5�5 � o  :=�4�4 0 thetime theTime� m  BE�� ���  :� n  FN��� 4  IN�3�
�3 
cwor� m  LM�2�2 � o  FI�1�1 0 thetime theTime� m  OR�� ���  A M� m  SV�0
�0 
ctxt� o      �/�/ 0 thetime2 theTime2z ��.� r  \|��� J  \w�� ��� n  \b��� 4  ]b�-�
�- 
cobj� m  `a�,�, � o  \]�+�+ 
0 lesson  � ��� n  bh��� 4  ch�*�
�* 
cobj� m  fg�)�) � o  bc�(�( 
0 lesson  � ��� o  hk�'�' 0 thetime2 theTime2� ��&� n  ks��� 4  ls�%�
�% 
cobj� m  or�$�$ � o  kl�#�# 
0 lesson  �&  � n      ���  ;  z{� o  wz�"�" 0 lessons4  �.  �U 
0 lesson  ` o  ���!�! 0 lessons3  �W  �V  ] ��� l     � ���   �  �  � ��� l     ����  � , & Combine Year, class, and teacher name   � ��� L   C o m b i n e   Y e a r ,   c l a s s ,   a n d   t e a c h e r   n a m e� ��� l ������ r  ����� J  ����  � o      �� 0 lessons5  �  �  � ��� l ������ X  ������ r  ����� J  ���� ��� b  ����� b  ����� n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  � m  ���� ���   � n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  � ��� n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  � ��� n  ����� 4  ����
� 
cobj� m  ���
�
 � o  ���	�	 
0 lesson  �  � n      ���  ;  ��� o  ���� 0 lessons5  � 
0 lesson  � o  ���� 0 lessons4  �  �  � ��� l     ����  �  �  � ��� l     ����  � ( " add date to each item in the list   � ��� D   a d d   d a t e   t o   e a c h   i t e m   i n   t h e   l i s t� ��� l ������ r  ����� J  ��� �   � o      ���� 0 lessons6  �  �  �    l ������ X  ���� Z  ���� E  ��	 n  ��

 4  ����
�� 
cobj m  ������  o  ������ 
0 lesson  	 m  �� �  M o n d a y r  � J  �  n  �� 4  ����
�� 
cobj m  ������  o  ������ 
0 lesson   �� b  � b  � n  �  1  � ��
�� 
dstr o  ������ 0 firstday firstDay m    �    a t   n  
 !  4  
��"
�� 
cobj" m  	���� ! o  ���� 
0 lesson  ��   n      #$#  ;  $ o  ���� 0 lessons6   %&% E  '(' n  )*) 4  ��+
�� 
cobj+ m  ���� * o  ���� 
0 lesson  ( m  ,, �--  T u e s d a y& ./. r  "E010 J  "@22 343 n  "(565 4  #(��7
�� 
cobj7 m  &'���� 6 o  "#���� 
0 lesson  4 8��8 b  (>9:9 b  (7;<; n  (3=>= 1  13��
�� 
dstr> l (1?����? [  (1@A@ o  (+���� 0 firstday firstDayA l +0B����B ]  +0CDC 1  +.��
�� 
daysD m  ./���� ��  ��  ��  ��  < m  36EE �FF    a t  : n  7=GHG 4  8=��I
�� 
cobjI m  ;<���� H o  78���� 
0 lesson  ��  1 n      JKJ  ;  CDK o  @C���� 0 lessons6  / LML E  HRNON n  HNPQP 4  IN��R
�� 
cobjR m  LM���� Q o  HI���� 
0 lesson  O m  NQSS �TT  W e d n e s d a yM UVU r  UxWXW J  UsYY Z[Z n  U[\]\ 4  V[��^
�� 
cobj^ m  YZ���� ] o  UV���� 
0 lesson  [ _��_ b  [q`a` b  [jbcb n  [fded 1  df��
�� 
dstre l [df����f [  [dghg o  [^���� 0 firstday firstDayh l ^ci����i ]  ^cjkj 1  ^a��
�� 
daysk m  ab���� ��  ��  ��  ��  c m  fill �mm    a t  a n  jpnon 4  kp��p
�� 
cobjp m  no���� o o  jk���� 
0 lesson  ��  X n      qrq  ;  vwr o  sv���� 0 lessons6  V sts E  {�uvu n  {�wxw 4  |���y
�� 
cobjy m  ����� x o  {|���� 
0 lesson  v m  ��zz �{{  T h u r s d a yt |}| r  ��~~ J  ���� ��� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ���� b  ����� b  ����� n  ����� 1  ����
�� 
dstr� l �������� [  ����� o  ������ 0 firstday firstDay� l �������� ]  ����� 1  ����
�� 
days� m  ������ ��  ��  ��  ��  � m  ���� ���    a t  � n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��   n      ���  ;  ��� o  ������ 0 lessons6  } ��� E  ����� n  ����� 4  �����
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
0 lesson   o  ������ 0 lessons5  ��  ��   ��� l     ��������  ��  ��  � ��� l     ������  � ) # convert each item to a date object   � ��� F   c o n v e r t   e a c h   i t e m   t o   a   d a t e   o b j e c t� ��� l �������� r  ����� J  ������  � o      ���� 0 lessons7  ��  ��  � ��� l �	4������ X  �	4����� k  		/�� ��� r  		��� n  		��� 4  		���
�� 
cobj� m  		���� � o  		���� 
0 lesson  � o      ���� 0 
datestring 
dateString� ��� r  		��� 4  		���
�� 
ldt � o  		���� 0 
datestring 
dateString� o      ���� 0 thedate theDate� ���� r  		/��� J  		*�� ��� n  		%��� 4  	 	%���
�� 
cobj� m  	#	$���� � o  		 ���� 
0 lesson  � ���� o  	%	(���� 0 thedate theDate��  � n      ���  ;  	-	.� o  	*	-���� 0 lessons7  ��  �� 
0 lesson  � o  ������ 0 lessons6  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  � #  Add room numbers to the list   � ��� :   A d d   r o o m   n u m b e r s   t o   t h e   l i s t� ��� l 	5	;������ r  	5	;��� J  	5	7��  � o      �~�~ 0 lessons8  ��  ��  � ��� l 	<	���}�|� X  	<	���{�� Z  	R	����z�� E  	R	\��� n  	R	X   4  	S	X�y
�y 
cobj m  	V	W�x�x  o  	R	S�w�w 
0 lesson  � m  	X	[ �  /� r  	_	u J  	_	p 	 n  	_	e

 4  	`	e�v
�v 
cobj m  	c	d�u�u  o  	_	`�t�t 
0 lesson  	  n  	e	k 4  	f	k�s
�s 
cobj m  	i	j�r�r  o  	e	f�q�q 
0 lesson   �p m  	k	n �  M u l t i p l e�p   n        ;  	s	t o  	p	s�o�o 0 lessons8  �z  � X  	x	��n Z  	�	��m�l E  	�	� n  	�	� 4  	�	��k
�k 
cobj m  	�	��j�j  o  	�	��i�i 
0 lesson   n  	�	� !  4  	�	��h"
�h 
cobj" m  	�	��g�g ! o  	�	��f�f 0 room   r  	�	�#$# J  	�	�%% &'& n  	�	�()( 4  	�	��e*
�e 
cobj* m  	�	��d�d ) o  	�	��c�c 
0 lesson  ' +,+ n  	�	�-.- 4  	�	��b/
�b 
cobj/ m  	�	��a�a . o  	�	��`�` 
0 lesson  , 0�_0 n  	�	�121 4  	�	��^3
�^ 
cobj3 m  	�	��]�] 2 o  	�	��\�\ 0 room  �_  $ n      454  ;  	�	�5 o  	�	��[�[ 0 lessons8  �m  �l  �n 0 room   o  	{	~�Z�Z 0 roomnumbers roomNumbers�{ 
0 lesson  � o  	?	B�Y�Y 0 lessons7  �}  �|  � 676 l     �X�W�V�X  �W  �V  7 898 l      �U:;�U  : � �
set classNames to {}
repeat with lesson in lessons8
	set end of classNames to item 1 of lesson
end repeat

set classNames2 to addr(classNames)
   ; �<<  
 s e t   c l a s s N a m e s   t o   { } 
 r e p e a t   w i t h   l e s s o n   i n   l e s s o n s 8 
 	 s e t   e n d   o f   c l a s s N a m e s   t o   i t e m   1   o f   l e s s o n 
 e n d   r e p e a t 
 
 s e t   c l a s s N a m e s 2   t o   a d d r ( c l a s s N a m e s ) 
9 =>= l 	�	�?�T�S? r  	�	�@A@ J  	�	��R�R  A o      �Q�Q 0 lessons9  �T  �S  > BCB l     �PDE�P  D &  set classNamesTeacherFirst to {}   E �FF @ s e t   c l a s s N a m e s T e a c h e r F i r s t   t o   { }C GHG l 	�
�I�O�NI X  	�
�J�MKJ k  	�
�LL MNM Z  	�
�OP�LQO E  	�	�RSR n  	�	�TUT 4  	�	��KV
�K 
cobjV m  	�	��J�J U o  	�	��I�I 
0 lesson  S m  	�	�WW �XX  /P k  	�
kYY Z[Z r  	�
\]\ n  	�	�^_^ 4  	�	��H`
�H 
cwor` m  	�	��G�G _ n  	�	�aba 4  	�	��Fc
�F 
cobjc m  	�	��E�E b o  	�	��D�D 
0 lesson  ] o      �C�C 0 teacherfirst teacherFirst[ ded r  

fgf n  

hih 4  

�Bj
�B 
cworj m  

�A�A i n  

klk 4  

�@m
�@ 
cobjm m  

�?�? l o  

�>�> 
0 lesson  g o      �=�= 0 teachersecond teacherSeconde non r  

#pqp n  

rsr 4  

�<t
�< 
cwort m  

�;�; s n  

uvu 4  

�:w
�: 
cobjw m  

�9�9 v o  

�8�8 
0 lesson  q o      �7�7 0 yearword yearWordo xyx r  
$
3z{z n  
$
/|}| 4  
*
/�6~
�6 
cwor~ m  
-
.�5�5 } n  
$
*� 4  
%
*�4�
�4 
cobj� m  
(
)�3�3 � o  
$
%�2�2 
0 lesson  { o      �1�1 0 
yearnumber 
yearNumbery ��� r  
4
C��� n  
4
?��� 4  
:
?�0�
�0 
cwor� m  
=
>�/�/ � n  
4
:��� 4  
5
:�.�
�. 
cobj� m  
8
9�-�- � o  
4
5�,�, 
0 lesson  � o      �+�+ 0 	classname 	className� ��*� r  
D
k��� b  
D
g��� b  
D
c��� b  
D
_��� b  
D
[��� b  
D
W��� b  
D
S��� b  
D
O��� b  
D
K��� o  
D
G�)�) 0 teacherfirst teacherFirst� m  
G
J�� ���  /� o  
K
N�(�( 0 teachersecond teacherSecond� m  
O
R�� ���   � o  
S
V�'�' 0 yearword yearWord� m  
W
Z�� ���   � o  
[
^�&�& 0 
yearnumber 
yearNumber� m  
_
b�� ���   � o  
c
f�%�% 0 	classname 	className� o      �$�$ 0 	eventname 	eventName�*  �L  Q k  
n
��� ��� r  
n
��� n  
n
{��� 4  
t
{�#�
�# 
cwor� m  
w
z�"�" � n  
n
t��� 4  
o
t�!�
�! 
cobj� m  
r
s� �  � o  
n
o�� 
0 lesson  � o      �� 0 teacher  � ��� r  
�
���� n  
�
���� 4  
�
���
� 
cwor� m  
�
��� � n  
�
���� 4  
�
���
� 
cobj� m  
�
��� � o  
�
��� 
0 lesson  � o      �� 0 yearword yearWord� ��� r  
�
���� n  
�
���� 4  
�
���
� 
cwor� m  
�
��� � n  
�
���� 4  
�
���
� 
cobj� m  
�
��� � o  
�
��� 
0 lesson  � o      �� 0 
yearnumber 
yearNumber� ��� r  
�
���� n  
�
���� 4  
�
���
� 
cwor� m  
�
��� � n  
�
���� 4  
�
���
� 
cobj� m  
�
��� � o  
�
��� 
0 lesson  � o      �� 0 	classname 	className� ��� r  
�
���� b  
�
���� b  
�
���� b  
�
���� b  
�
���� b  
�
���� b  
�
���� o  
�
��
�
 0 teacher  � m  
�
��� ���   � o  
�
��	�	 0 yearword yearWord� m  
�
��� ���   � o  
�
��� 0 
yearnumber 
yearNumber� m  
�
��� ���   � o  
�
��� 0 	classname 	className� o      �� 0 	eventname 	eventName�  N ��� r  
�
���� J  
�
��� ��� o  
�
��� 0 	eventname 	eventName� ��� n  
�
���� 4  
�
���
� 
cobj� m  
�
��� � o  
�
��� 
0 lesson  � �� � n  
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
0 lesson  �   � n      ���  ;  
�
�� o  
�
����� 0 lessons9  �  �M 
0 lesson  K o  	�	����� 0 lessons8  �O  �N  H ��� l     ��������  ��  ��  � ��� l     ������  � . ( Creates the calendars in the chosen app   � ��� P   C r e a t e s   t h e   c a l e n d a r s   i n   t h e   c h o s e n   a p p� ��� l 
������� Z  
� ��  = 
�
� o  
�
����� &0 chosencalendarapp chosenCalendarApp m  
�
� �  A p p l e   C a l e n d a r k  
�D 	 l 
�
���
��  
 &   Deletes previous work calendars    � @   D e l e t e s   p r e v i o u s   w o r k   c a l e n d a r s	  O  
�+ k  
�*  I 
�
�������
�� .miscactvnull��� ��� null��  ��    O   I ����
�� .sysottosnull���     TEXT m   � n D e l e t i n g   p r e v i o u s   C a l e n d a r s   w i t h   t h e   w o r d   " Y e a r "   i n   i t .��    f     l ��������  ��  ��   �� O  * I $)������
�� .coredelonull���     obj ��  ��   l ! ����  6 !!"! 2 ��
�� 
wres" E   #$# 1  ��
�� 
pnam$ m  %% �&&  Y e a r��  ��  ��   m  
�
�''�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��   ()( l ,,��������  ��  ��  ) *+* l ,,��������  ��  ��  + ,-, l ,,��./��  . 3 - add calendars and events to the Calendar app   / �00 Z   a d d   c a l e n d a r s   a n d   e v e n t s   t o   t h e   C a l e n d a r   a p p- 121 I ,3��3��
�� .sysottosnull���     TEXT3 m  ,/44 �55 $ C r e a t i n g   C a l e n d a r s��  2 676 O  4�898 k  :�:: ;<; I :?������
�� .miscactvnull��� ��� null��  ��  < =>= Y  @?��@A��? Z  PzBC����B H  PcDD l PbE����E I Pb��F��
�� .coredoexnull���     ****F l P^G����G 4  P^��H
�� 
wresH l T]I����I n  T]JKJ 4  W\��L
�� 
cobjL o  Z[���� 0 i  K o  TW���� 0 
classnames 
classNames��  ��  ��  ��  ��  ��  ��  C I fv����M
�� .wrbtaec2null��� ��� null��  M ��N��
�� 
wtnmN n  jrOPO 4  mr��Q
�� 
cobjQ o  pq���� 0 i  P o  jm���� 0 
classnames 
classNames��  ��  ��  �� 0 i  @ m  CD���� A I DK��R��
�� .corecnte****       ****R o  DG���� 0 
classnames 
classNames��  ��  > STS O ��UVU I ����W��
�� .sysottosnull���     TEXTW m  ��XX �YY  C r e a t i n g   E v e n t s��  V  f  ��T Z��Z X  ��[��\[ Z  ��]^��_] I ����`��
�� .coredoexnull���     ****` 4  ����a
�� 
wresa l ��b����b n  ��cdc 4  ����e
�� 
cobje m  ������ d o  ������ 
0 lesson  ��  ��  ��  ^ O  �fgf I �����h
�� .corecrel****      � null��  h ��ij
�� 
kocli m  ����
�� 
wrevj ��k��
�� 
prdtk K  �ll ��mn
�� 
wr11m n  ��opo 4  ����q
�� 
cobjq m  ������ p o  ������ 
0 lesson  n ��rs
�� 
wr14r n  ��tut 4  ����v
�� 
cobjv m  ������ u o  ������ 
0 lesson  s ��wx
�� 
wr1sw n  ��yzy 4  ����{
�� 
cobj{ m  ������ z o  ������ 
0 lesson  x ��|}
�� 
wr5s| [  ��~~ l �������� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  ��   l �������� ]  ����� 1  ����
�� 
min � m  ������ P��  ��  } �����
�� 
wr15� o  ����� "0 endofrecurrence endOfRecurrence��  ��  g 4  �����
�� 
wres� l �������� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  ��  ��  _ k  ��� ��� r  ��� n  ��� 2 ��
�� 
cwor� n  ��� 4  ���
�� 
cobj� m  ���� � o  ���� 
0 lesson  � o      ���� 0 problem  � ��� r  S��� b  O��� b  D��� b  @��� b  7��� b  3��� b  *��� n  &��� 4  !&���
�� 
cobj� m  $%���� � o  !���� 0 problem  � m  &)�� ���   � n  *2��� 4  -2���
�� 
cobj� m  01�� � o  *-�~�~ 0 problem  � m  36�� ���   � n  7?��� 4  :?�}�
�} 
cobj� m  =>�|�| � o  7:�{�{ 0 problem  � m  @C�� ���   � n  DN��� 4  GN�z�
�z 
cobj� m  JM�y�y � o  DG�x�x 0 problem  � o      �w�w "0 neweventsummary newEventSummary� ��v� O  T���� I _��u�t�
�u .corecrel****      � null�t  � �s��
�s 
kocl� m  cf�r
�r 
wrev� �q��p
�q 
prdt� K  i��� �o��
�o 
wr11� n  lr��� 4  mr�n�
�n 
cobj� m  pq�m�m � o  lm�l�l 
0 lesson  � �k��
�k 
wr14� n  u{��� 4  v{�j�
�j 
cobj� m  yz�i�i � o  uv�h�h 
0 lesson  � �g��
�g 
wr1s� n  ~���� 4  ��f�
�f 
cobj� m  ���e�e � o  ~�d�d 
0 lesson  � �c��
�c 
wr5s� [  ����� l ����b�a� n  ����� 4  ���`�
�` 
cobj� m  ���_�_ � o  ���^�^ 
0 lesson  �b  �a  � l ����]�\� ]  ����� 1  ���[
�[ 
min � m  ���Z�Z P�]  �\  � �Y��X
�Y 
wr15� o  ���W�W "0 endofrecurrence endOfRecurrence�X  �p  � 4  T\�V�
�V 
wres� o  X[�U�U "0 neweventsummary newEventSummary�v  �� 
0 lesson  \ o  ���T�T 0 lessons9  ��  9 m  47���                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  7 ��� l ���S�R�Q�S  �R  �Q  � ��� l ���P�O�N�P  �O  �N  � ��� l ���M���M  �   Sets the time zone   � ��� &   S e t s   t h e   t i m e   z o n e� ��� O ����� I ���L��K
�L .sysottosnull���     TEXT� m  ���� ��� ^ C h a n g i n g   t h e   t i m e   z o n e   t o   t h e   C h i n e s e   t i m e   z o n e�K  �  f  ��� ��� l ���J�I�H�J  �I  �H  � ��� O  ����� r  ����� 6 ����� n  ����� 1  ���G
�G 
pnam� 2 ���F
�F 
wres� E  ����� 1  ���E
�E 
pnam� m  ���� ���  Y e a r� o      �D�D 0 thecalendars theCalendars� m  �����                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  � � � l ���C�B�A�C  �B  �A     r  �� I ���@�? z�> 
�> .!Cls!fstnull��� ��� null�@  �?   o      �=�= 0 thestore theStore  l ���<�;�:�<  �;  �:   	 r  ��

 o  ���9�9 0 firstday firstDay o      �8�8 0 	startdate 	startDate	  r  � [  � o  ���7�7 0 	startdate 	startDate ]  � 1  ��6
�6 
days m  �5�5  o      �4�4 0 enddate endDate  l �3�2�1�3  �2  �1    r  1 I - z�0 
�0 .!CLs!fcAnull���     **** o  �/�/ 0 thecalendars theCalendars �.
�. 
!CtY m    z�- 
�- !Tct!TtL �, �+
�, 
!Cst  o  #&�*�* 0 thestore theStore�+   o      �)�) 0 thecalendars theCalendars !"! l 22�(�'�&�(  �'  �&  " #$# r  2a%&% I 2]'�%(' z�$ 
�$ .!CLs!feSnull��� ��� null�%  ( �#)*
�# 
!Csd) o  AD�"�" 0 	startdate 	startDate* �!+,
�! 
!Ced+ o  GJ� �  0 enddate endDate, �-.
� 
!Csc- o  MP�� 0 thecalendars theCalendars. �/�
� 
!Cst/ o  SV�� 0 thestore theStore�  & o      �� 0 	theevents 	theEvents$ 010 l bb����  �  �  1 232 Y  b�4�56�4 k  r�77 898 r  r�:;: I r�<�=< z� 
� .!CLs!mo2null��� ��� null�  = �>?
� 
!Cev> n  ��@A@ 4  ���B
� 
cobjB o  ���� 0 i  A o  ���� 0 	theevents 	theEvents? �C�
� 
!CtzC m  ��DD �EE  A s i a / S h a n g h a i�  ; o      �� 0 thenewevent theNewEvent9 F�F I ��G�
HG z�	 
�	 .!CLs!updnull��� ��� null�
  H �IJ
� 
!CevI o  ���� 0 thenewevent theNewEventJ �K�
� 
!CstK o  ���� 0 thestore theStore�  �  � 0 i  5 m  ef�� 6 I fm�L�
� .corecnte****       ****L o  fi� �  0 	theevents 	theEvents�  �  3 MNM l ����������  ��  ��  N OPO l ����������  ��  ��  P QRQ l ����ST��  S + % remove days because of delayed start   T �UU J   r e m o v e   d a y s   b e c a u s e   o f   d e l a y e d   s t a r tR VWV Z  ��XY����X = ��Z[Z o  ������  0 delayedyesorno delayedYesOrNo[ m  ��\\ �]]  Y e sY k  ��^^ _`_ r  ��aba I ��c����c z�� 
�� .!Cls!fstnull��� ��� null��  ��  b o      ���� 0 thestore theStore` ded l ����������  ��  ��  e fgf r  �hih I � jklj z�� 
�� .!CLs!fcAnull���     ****k J  ������  l ��mn
�� 
!CtYm m  ��oo z�� 
�� !Tct!TtCn ��p��
�� 
!Cstp o  ������ 0 thestore theStore��  i o      ���� 0 thecalendars theCalendarsg qrq r  sts o  ���� 0 firstday firstDayt o      ���� 0 startingdate startingDater uvu r  wxw \  yzy o  ���� "0 delayedfirstday delayedFirstDayz l {����{ ]  |}| 1  ��
�� 
days} m  ���� ��  ��  x o      ���� 0 
endingdate 
endingDatev ~~ r  J��� I F����� z�� 
�� .!CLs!feSnull��� ��� null��  � ����
�� 
!Csd� o  *-���� 0 startingdate startingDate� ����
�� 
!Ced� o  03���� 0 
endingdate 
endingDate� ����
�� 
!Csc� o  69���� 0 thecalendars theCalendars� �����
�� 
!Cst� o  <?���� 0 thestore theStore��  � o      ���� 0 	theevents 	theEvents ��� l KK��������  ��  ��  � ��� X  K������ I a������ z�� 
�� .!CLs!reEnull��� ��� null��  � ����
�� 
!Cev� o  pq���� 0 theevent theEvent� ����
�� 
!Cst� o  tw���� 0 thestore theStore� �����
�� 
!Cft� m  z{��
�� boovfals��  �� 0 theevent theEvent� o  NQ���� 0 	theevents 	theEvents� ���� l ����������  ��  ��  ��  ��  ��  W ��� l ����������  ��  ��  � ��� l ��������  � k e This turns off and on the iCloud Calendar in preferences so the calendars gets uploaded to the cloud   � ��� �   T h i s   t u r n s   o f f   a n d   o n   t h e   i C l o u d   C a l e n d a r   i n   p r e f e r e n c e s   s o   t h e   c a l e n d a r s   g e t s   u p l o a d e d   t o   t h e   c l o u d� ��� I �������
�� .sysottosnull���     TEXT� m  ���� ��� J M o v i n g   C a l e n d a r s   f r o m   l o c a l   t o   i C l o u d��  � ��� O  ����� k  ���� ��� I ��������
�� .miscactvnull��� ��� null��  ��  � ��� I �������
�� .sysodelanull��� ��� nmbr� m  ������ ��  � ��� I ��������
�� .aevtquitnull��� ��� null��  ��  � ��� I ��������
�� .miscactvnull��� ��� null��  ��  � ���� I �������
�� .sysodelanull��� ��� nmbr� m  ������ ��  ��  � m  �����                                                                                  sprf  alis    `  Macintosh HD                   BD ����System Preferences.app                                         ����            ����  
 cu             Applications  -/:System:Applications:System Preferences.app/   .  S y s t e m   P r e f e r e n c e s . a p p    M a c i n t o s h   H D  *System/Applications/System Preferences.app  / ��  � ��� l ����������  ��  ��  � ��� O  ����� O  ����� O  ����� O  ����� k  ���� ��� I ��������
�� .prcsclicnull��� ��� uiel��  ��  � ���� I �������
�� .sysodelanull��� ��� nmbr� m  ������ ��  ��  � 4  �����
�� 
butT� m  ������ � 4  �����
�� 
cwin� m  ������ � 4  �����
�� 
prcs� m  ���� ��� $ S y s t e m   P r e f e r e n c e s� m  �����                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � ��� l ����������  ��  ��  � ��� O  �|��� O  �{��� O  �z��� O  y��� O  x��� O  w��� O  #v��� O  .u��� O  7t��� k  @s�� ��� O @L��� I DK�����
�� .sysottosnull���     TEXT� m  DG�� ��� 8 U n c h e c k i n g   i C l o u d   C a l e n d a r s .��  �  f  @A� ��� l MR���� I MR������
�� .prcsclicnull��� ��� uiel��  ��  � #  turn off icloud for calendar   � ��� :   t u r n   o f f   i c l o u d   f o r   c a l e n d a r� ��� I SZ�����
�� .sysodelanull��� ��� nmbr� m  SV���� ��  � ��� O [g��� I _f�����
�� .sysottosnull���     TEXT� m  _b�� ��� 4 C h e c k i n g   i C l o u d   C a l e n d a r s .��  �  f  [\� ��� l hm���� I hm������
�� .prcsclicnull��� ��� uiel��  ��  � "  turn on icloud for calendar   � �	 	  8   t u r n   o n   i c l o u d   f o r   c a l e n d a r� 	��	 I ns��	��
�� .sysodelanull��� ��� nmbr	 m  no���� ��  ��  � 4  7=��	
�� 
uiel	 m  ;<���� � 4  .4��	
�� 
uiel	 m  23���� � 4  #+��	
�� 
crow	 m  '*���� � 4   ��	
�� 
tabB	 m  ���� � 4  �	
� 
scra	 m  �~�~ � 4  �}	
�} 
sgrp	 m  �|�| � 4  ��{		
�{ 
cwin		 m  �z�z � 4  ���y	

�y 
prcs	
 m  ��		 �		 $ S y s t e m   P r e f e r e n c e s� m  ��		�                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � 			 l  }}�x		�x  	 � �
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
	   	 �		� 
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
 		 			 O  }�			 k  ��		 			 I ���w�v�u
�w .miscactvnull��� ��� null�v  �u  	 	�t	 I ���s	�r
�s .sysodelanull��� ��� nmbr	 m  ���q�q Z�r  �t  	 m  }�		�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  	 			 l ���p�o�n�p  �o  �n  	 		 	 l ���m	!	"�m  	! ( " Publishes the Calendars as Public   	" �	#	# D   P u b l i s h e s   t h e   C a l e n d a r s   a s   P u b l i c	  	$	%	$ I ���l	&�k
�l .sysottosnull���     TEXT	& m  ��	'	' �	(	( \ P u b l i s h i n g   t h e   c a l e n d a r s   b y   m a k i n g   t h e m   p u b l i c�k  	% 	)	*	) O  �[	+	,	+ O  �Z	-	.	- k  �Y	/	/ 	0	1	0 l ���j�i�h�j  �i  �h  	1 	2	3	2 l  ���g	4	5�g  	4��
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
				   	5 �	6	6| 
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
 	 	 	 		3 	7	8	7 l ���f�e�d�f  �e  �d  	8 	9	:	9 l ��	;	<	=	; I ���c	>	?
�c .prcskprsnull���     ctxt	> m  ��	@	@ �	A	A  f	? �b	B�a
�b 
faal	B m  ���`
�` eMdsKcmd�a  	<   search bar   	= �	C	C    s e a r c h   b a r	: 	D	E	D I ���_	F�^
�_ .sysodelanull��� ��� nmbr	F m  ���]�] �^  	E 	G	H	G l ���\�[�Z�\  �[  �Z  	H 	I	J	I l ��	K	L	M	K I ���Y	N�X
�Y .prcskprsnull���     ctxt	N 1  ���W
�W 
tab �X  	L #  moves focus to calendar list   	M �	O	O :   m o v e s   f o c u s   t o   c a l e n d a r   l i s t	J 	P	Q	P I ���V	R�U
�V .sysodelanull��� ��� nmbr	R m  ���T�T �U  	Q 	S	T	S l ���S�R�Q�S  �R  �Q  	T 	U	V	U l ��	W	X	Y	W I ���P	Z	[
�P .prcskcodnull���     ****	Z m  ���O�O ~	[ �N	\�M
�N 
faal	\ J  ��	]	] 	^�L	^ m  ���K
�K eMdsKopt�L  �M  	X 6 0 moves focus to first calendar / option up arrow   	Y �	_	_ `   m o v e s   f o c u s   t o   f i r s t   c a l e n d a r   /   o p t i o n   u p   a r r o w	V 	`	a	` I ���J	b�I
�J .sysodelanull��� ��� nmbr	b m  ���H�H �I  	a 	c	d	c l ���G�F�E�G  �F  �E  	d 	e	f	e r  ��	g	h	g J  ���D�D  	h o      �C�C 0 calendarnames calendarNames	f 	i�B	i Y  �Y	j�A	k	l�@	j k  T	m	m 	n	o	n Q  R	p	q�?	p k  I	r	r 	s	t	s r  I	u	v	u n  E	w	x	w 1  AE�>
�> 
valL	x n  A	y	z	y 4  <A�=	{
�= 
txtf	{ m  ?@�<�< 	z n  <	|	}	| 4  7<�;	~
�; 
uiel	~ m  :;�:�: 	} n  7		�	 4  27�9	�
�9 
crow	� o  56�8�8 0 i  	� n  2	�	�	� 4  -2�7	�
�7 
outl	� m  01�6�6 	� n  -	�	�	� 4  (-�5	�
�5 
scra	� m  +,�4�4 	� n  (	�	�	� 4  #(�3	�
�3 
splg	� m  &'�2�2 	� n  #	�	�	� 4  #�1	�
�1 
splg	� m  !"�0�0 	� 4  �/	�
�/ 
cwin	� m  �.�. 	v o      �-�- $0 calendarnametext calendarNameText	t 	��,	� Z  JI	�	��+�*	� E  JQ	�	�	� o  JM�)�) $0 calendarnametext calendarNameText	� m  MP	�	� �	�	�  Y e a r	� k  TE	�	� 	�	�	� r  T\	�	�	� o  TW�(�( $0 calendarnametext calendarNameText	� n      	�	�	�  ;  Z[	� o  WZ�'�' 0 calendarnames calendarNames	� 	�	�	� l ]j	�	�	�	� I ]j�&	�	�
�& .prcskprsnull���     ctxt	� m  ]`	�	� �	�	�  f	� �%	��$
�% 
faal	� m  cf�#
�# eMdsKcmd�$  	�   search bar   	� �	�	�    s e a r c h   b a r	� 	�	�	� I kp�"	��!
�" .sysodelanull��� ��� nmbr	� m  kl� �  �!  	� 	�	�	� l qq����  �  �  	� 	�	�	� l qx	�	�	�	� I qx�	��
� .prcskprsnull���     ctxt	� 1  qt�
� 
tab �  	� #  moves focus to calendar list   	� �	�	� :   m o v e s   f o c u s   t o   c a l e n d a r   l i s t	� 	�	�	� l yy�	�	��  	�  						delay 1   	� �	�	�  	 	 	 	 	 	 d e l a y   1	� 	�	�	� l yy����  �  �  	� 	�	�	� l y�	�	�	�	� I y��	��
� .prcskcodnull���     ****	� m  y|�� }�  	�   down arrow   	� �	�	�    d o w n   a r r o w	� 	�	�	� I ���	��
� .sysodelanull��� ��� nmbr	� m  ���� �  	� 	�	�	� l ������  �  �  	� 	�	�	� O  ��	�	�	� l ��	�	�	�	� I �����

� .prcsclicnull��� ��� uiel�  �
  	�   click Edit Menu   	� �	�	�     c l i c k   E d i t   M e n u	� n  ��	�	�	� 4  ���		�
�	 
menE	� m  ��	�	� �	�	�  E d i t	� n  ��	�	�	� 4  ���	�
� 
mbri	� m  ��	�	� �	�	�  E d i t	� 4  ���	�
� 
mbar	� m  ���� 	� 	�	�	� l ���	�	��  	�  						delay 1   	� �	�	�  	 	 	 	 	 	 d e l a y   1	� 	�	�	� O  ��	�	�	� l ��	�	�	�	� I �����
� .prcsclicnull��� ��� uiel�  �  	� + % click Share Calendar under Edit Menu   	� �	�	� J   c l i c k   S h a r e   C a l e n d a r   u n d e r   E d i t   M e n u	� n  ��	�	�	� 4  ���	�
� 
uiel	� m  ��� �  	� n  ��	�	�	� 4  ����	�
�� 
menE	� m  ��	�	� �	�	�  E d i t	� n  ��	�	�	� 4  ����	�
�� 
mbri	� m  ��	�	� �	�	�  E d i t	� 4  ����	�
�� 
mbar	� m  ������ 	� 	�	�	� l ����	�	���  	�  						delay 1   	� �	�	�  	 	 	 	 	 	 d e l a y   1	� 	�	�	� O  ��	�	�	� l ��	�	�
 	� I ��������
�� .prcsclicnull��� ��� uiel��  ��  	� ( " click the share calendar checkbox   
  �

 D   c l i c k   t h e   s h a r e   c a l e n d a r   c h e c k b o x	� n  ��


 4  ����

�� 
chbx
 m  ������ 
 n  ��


 4  ����

�� 
popv
 m  ������ 
 n  ��

	
 4  ����


�� 
crow

 o  ������ 0 i  
	 n  ��


 4  ����

�� 
outl
 m  ������ 
 n  ��


 4  ����

�� 
scra
 m  ������ 
 n  ��


 4  ����

�� 
splg
 m  ������ 
 n  ��


 4  ����

�� 
splg
 m  ������ 
 4  ����

�� 
cwin
 m  ������ 	� 


 l ����

��  
  						delay 1   
 �

  	 	 	 	 	 	 d e l a y   1
 


 O  �1

 
 I +0������
�� .prcsclicnull��� ��� uiel��  ��  
  n  �(
!
"
! 4  !(��
#
�� 
butT
# m  $'
$
$ �
%
%  D o n e
" n  �!
&
'
& 4  !��
(
�� 
popv
( m   ���� 
' n  �
)
*
) 4  ��
+
�� 
crow
+ o  ���� 0 i  
* n  �
,
-
, 4  ��
.
�� 
outl
. m  ���� 
- n  �
/
0
/ 4  ��
1
�� 
scra
1 m  ���� 
0 n  �
2
3
2 4  ��
4
�� 
splg
4 m  ���� 
3 n  �
5
6
5 4  ��
7
�� 
splg
7 m  ���� 
6 4  ���
8
�� 
cwin
8 m  ���� 
 
9
:
9 l 22��
;
<��  
; 7 1					keystroke "f" using command down # seach bar   
< �
=
= b 	 	 	 	 	 k e y s t r o k e   " f "   u s i n g   c o m m a n d   d o w n   #   s e a c h   b a r
: 
>
?
> I 27��
@��
�� .sysodelanull��� ��� nmbr
@ m  23���� ��  
? 
A
B
A l 8?
C
D
E
C I 8?��
F��
�� .prcskprsnull���     ctxt
F 1  8;��
�� 
tab ��  
D #  moves focus to calendar list   
E �
G
G :   m o v e s   f o c u s   t o   c a l e n d a r   l i s t
B 
H��
H I @E��
I��
�� .sysodelanull��� ��� nmbr
I m  @A���� ��  ��  �+  �*  �,  	q R      ������
�� .ascrerr ****      � ****��  ��  �?  	o 
J��
J l SS��������  ��  ��  ��  �A 0 i  	k m  ������ 	l I ���
K��
�� .corecnte****       ****
K n �
L
M
L 2 ��
�� 
crow
M n  �
N
O
N 4  ��
P
�� 
outl
P m  ���� 
O n  �
Q
R
Q 4  ���
S
�� 
scra
S m  ���� 
R n  ��
T
U
T 4  ����
V
�� 
splg
V m  ������ 
U n  ��
W
X
W 4  ����
Y
�� 
splg
Y m  ������ 
X 4  ����
Z
�� 
cwin
Z m  ������ ��  �@  �B  	. 4  ����
[
�� 
prcs
[ m  ��
\
\ �
]
]  C a l e n d a r	, m  ��
^
^�                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  	* 
_
`
_ l \\��������  ��  ��  
` 
a
b
a l \\��
c
d��  
c N H Grabs the published calendars and writes the URLs to a numbers document   
d �
e
e �   G r a b s   t h e   p u b l i s h e d   c a l e n d a r s   a n d   w r i t e s   t h e   U R L s   t o   a   n u m b e r s   d o c u m e n t
b 
f
g
f O  \h
h
i
h I bg������
�� .miscactvnull��� ��� null��  ��  
i m  \_
j
j�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  
g 
k
l
k I ip��
m��
�� .sysottosnull���     TEXT
m m  il
n
n �
o
o X G r a b b i n g   t h e   U R L s   f o r   t h e s e   p u b l i c   c a l e n d a r s��  
l 
p
q
p l qq��������  ��  ��  
q 
r
s
r r  qw
t
u
t J  qs����  
u o      ���� 0 calendarurls calendarURLs
s 
v
w
v O  xV
x
y
x O  ~U
z
{
z k  �T
|
| 
}
~
} I ����

�
�� .prcskprsnull���     ctxt
 m  ��
�
� �
�
�  f
� ��
���
�� 
faal
� m  ����
�� eMdsKcmd��  
~ 
�
�
� I ����
���
�� .sysodelanull��� ��� nmbr
� m  ������ ��  
� 
�
�
� I ����
���
�� .prcskprsnull���     ctxt
� 1  ����
�� 
tab ��  
� 
�
�
� I ����
���
�� .sysodelanull��� ��� nmbr
� m  ������  ��  
� 
�
�
� I ����
�
�
�� .prcskcodnull���     ****
� m  ������ ~
� ��
���
�� 
faal
� J  ��
�
� 
���
� m  ����
�� eMdsKopt��  ��  
� 
�
�
� I ����
���
�� .sysodelanull��� ��� nmbr
� m  ������  ��  
� 
�
�
� l ����������  ��  ��  
� 
���
� Y  �T
���
�
���
� k  �O
�
� 
�
�
� I ����
���
�� .prcskcodnull���     ****
� m  ������ }��  
� 
�
�
� I ����
���
�� .sysodelanull��� ��� nmbr
� m  ������  ��  
� 
�
�
� I ����
�
�
�� .prcskprsnull���     ctxt
� m  ��
�
� �
�
�  i
� ��
���
�� 
faal
� m  ����
�� eMdsKcmd��  
� 
�
�
� I ���
��~
� .sysodelanull��� ��� nmbr
� m  ���}�}  �~  
� 
�
�
� O  �3
�
�
� O  �2
�
�
� k  1
�
� 
�
�
� r  
�
�
� n  
�
�
� 1  �|
�| 
valL
� 4  �{
�
�{ 
txtf
� m  	
�z�z 
� o      �y�y 0 calendarname calendarName
� 
��x
� O  1
�
�
� O  0
�
�
� r  &/
�
�
� 1  &+�w
�w 
valL
� o      �v�v 0 theurl theURL
� 4  #�u
�
�u 
txta
� m  !"�t�t 
� 4  �s
�
�s 
scra
� m  �r�r �x  
� 4  ��q
�
�q 
sheE
� m   �p�p 
� 4  ���o
�
�o 
cwin
� m  ���n�n 
� 
�
�
� I 4;�m
��l
�m .prcskprsnull���     ctxt
� o  47�k
�k 
ret �l  
� 
�
�
� I <A�j
��i
�j .sysodelanull��� ��� nmbr
� m  <=�h�h  �i  
� 
��g
� r  BO
�
�
� J  BJ
�
� 
�
�
� o  BE�f�f 0 calendarname calendarName
� 
��e
� o  EH�d�d 0 theurl theURL�e  
� n      
�
�
�  ;  MN
� o  JM�c�c 0 calendarurls calendarURLs�g  �� 0 i  
� m  ���b�b 
� I ���a
��`
�a .corecnte****       ****
� o  ���_�_ 0 calendarnames calendarNames�`  ��  ��  
{ 4  ~��^
�
�^ 
prcs
� m  ��
�
� �
�
�  C a l e n d a r
y m  x{
�
��                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  
w 
�
�
� l WW�]�\�[�]  �\  �[  
� 
�
�
� I W^�Z
��Y
�Z .sysottosnull���     TEXT
� m  WZ
�
� �
�
� p P a s t i n g   p u b l i c   c a l e n d a r   l i n k s   i n t o   a   N u m b e r s   s p r e a d s h e e t�Y  
� 
�
�
� O  _B
�
�
� k  eA
�
� 
�
�
� I ej�X�W�V
�X .miscactvnull��� ��� null�W  �V  
� 
�
�
� I kv�U�T
�
�U .corecrel****      � null�T  
� �S
��R
�S 
kocl
� m  or�Q
�Q 
docu�R  
� 
��P
� O  wA
�
�
� O  �@
�
�
� O  �?
�
�
� k  �>
�
� 
�
�
� O  ��
�
�
� k  ��
�
� 
�
�
� r  ��
�
�
� m  ��
�
� �
�
�  C a l e n d a r   N a m e
� n         1  ���O
�O 
NMCv 4  ���N
�N 
NmCl m  ���M�M 
� �L r  �� m  �� �  U R L n      	 1  ���K
�K 
NMCv	 4  ���J

�J 
NmCl
 m  ���I�I �L  
� 4  ���H
�H 
NMRw m  ���G�G 
�  l ���F�E�D�F  �E  �D   �C Y  �>�B�A k  �9  r  �� n  �� 4  ���@
�@ 
cobj m  ���?�?  n  �� 4  ���>
�> 
cobj o  ���=�= 0 i   o  ���<�< 0 calendarurls calendarURLs o      �;�; 0 calendarname calendarName  r  ��  n  ��!"! 4  ���:#
�: 
cobj# m  ���9�9 " n  ��$%$ 4  ���8&
�8 
cobj& o  ���7�7 0 i  % o  ���6�6 0 calendarurls calendarURLs  o      �5�5 0 theurl theURL '(' Z  �)*�4�3) H  ��++ l ��,�2�1, I ���0-�/
�0 .coredoexnull���     ****- 4  ���..
�. 
NMRw. l ��/�-�,/ [  ��010 o  ���+�+ 0 i  1 m  ���*�* �-  �,  �/  �2  �1  * I �	�)�(2
�) .corecrel****      � null�(  2 �'3�&
�' 
kocl3 m  �%
�% 
NMRw�&  �4  �3  ( 454 O  7676 k  688 9:9 r  ';<; o  �$�$ 0 calendarname calendarName< n      =>= 1  "&�#
�# 
NMCv> 4  "�"?
�" 
NmCl? m   !�!�! : @� @ r  (6ABA o  (+�� 0 theurl theURLB n      CDC 1  15�
� 
NMCvD 4  +1�E
� 
NmClE m  /0�� �   7 4  �F
� 
NMRwF l G��G [  HIH o  �� 0 i  I m  �� �  �  5 J�J l 88����  �  �  �  �B 0 i   m  ����  I ���K�
� .corecnte****       ****K o  ���� 0 calendarurls calendarURLs�  �A  �C  
� 4  ���L
� 
NmTbL m  ���� 
� 4  ���M
� 
NmShM m  ���� 
� 4  w}�
N
�
 
docuN m  {|�	�	 �P  
� m  _bOO�                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  
� PQP l CC����  �  �  Q R�R l CC����  �  �  �  ��   k  GSS TUT l GG�VW�  V 9 3 Outlook --> Incomplete code, needs repeat function   W �XX f   O u t l o o k   - - >   I n c o m p l e t e   c o d e ,   n e e d s   r e p e a t   f u n c t i o nU YZY l GG� [\�   [ B < Creates all the events on the first day but does not repeat   \ �]] x   C r e a t e s   a l l   t h e   e v e n t s   o n   t h e   f i r s t   d a y   b u t   d o e s   n o t   r e p e a tZ ^_^ O  G�`a` k  M�bb cdc I MR������
�� .miscactvnull��� ��� null��  ��  d efe l SS��������  ��  ��  f ghg I Sk��i��
�� .coredelonull���     obj i l Sgj����j 6 Sgklk 2 SX��
�� 
cCFol E  [fmnm 1  \`��
�� 
pnamn m  aeoo �pp  Y e a r��  ��  ��  h qrq l ll��������  ��  ��  r sts r  l�uvu 6 l�wxw 4 lr��y
�� 
cCFoy m  pq���� x E  u�z{z n  v~|}| m  z~��
�� 
emad} 2 vz��
�� 
cAct{ m  �~~ �  w i n t e cv o      ���� 0 wintecaccount wintecAccountt ���� O  ����� k  ���� ��� l ����������  ��  ��  � ��� Y  ���������� Z  ��������� H  ���� l �������� I �������
�� .coredoexnull���     ****� l �������� 4  �����
�� 
cCFo� l �������� n  ����� 4  �����
�� 
cobj� o  ������ 0 i  � o  ������ 0 
classnames 
classNames��  ��  ��  ��  ��  ��  ��  � I �������
�� .corecrel****      � null��  � ����
�� 
kocl� m  ����
�� 
cCFo� �����
�� 
prdt� K  ���� �����
�� 
pnam� n  ����� 4  �����
�� 
cobj� o  ������ 0 i  � o  ������ 0 
classnames 
classNames��  ��  ��  ��  �� 0 i  � m  ������ � I �������
�� .corecnte****       ****� o  ������ 0 
classnames 
classNames��  ��  � ��� l ����������  ��  ��  � ��� X  ������� Z  �������� I ������
�� .coredoexnull���     ****� 4  �����
�� 
cCFo� l �������� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  ��  ��  � O  Z��� I Y�����
�� .corecrel****      � null��  � ����
�� 
kocl� m  ��
�� 
cEvt� �����
�� 
prdt� K  S�� ����
�� 
subj� n  "(��� 4  #(���
�� 
cobj� m  &'���� � o  "#���� 
0 lesson  � ����
�� 
iloc� n  +1��� 4  ,1���
�� 
cobj� m  /0���� � o  +,���� 
0 lesson  � ����
�� 
offs� n  4:��� 4  5:���
�� 
cobj� m  89���� � o  45���� 
0 lesson  � ����
�� 
endT� [  =K��� l =C������ n  =C��� 4  >C���
�� 
cobj� m  AB���� � o  =>���� 
0 lesson  ��  ��  � l CJ������ ]  CJ��� 1  CF��
�� 
min � m  FI���� P��  ��  � �����
�� 
pHsR� m  NO��
�� boovfals��  ��  � 4  ���
�� 
cCFo� l 
������ n  
��� 4  ���
�� 
cobj� m  ���� � o  
���� 
0 lesson  ��  ��  ��  � k  ]��� ��� r  ]k��� n  ]g��� 2 cg��
�� 
cwor� n  ]c��� 4  ^c���
�� 
cobj� m  ab���� � o  ]^���� 
0 lesson  � o      ���� 0 problem  � ��� r  l���� b  l���� b  l���� b  l���� b  l���� b  l���� b  lx��� n  lt��� 4  ot���
�� 
cobj� m  rs���� � o  lo���� 0 problem  � m  tw�� ���   � n  x���� 4  {����
�� 
cobj� m  ~���� � o  x{���� 0 problem  � m  ���� ���   � n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 0 problem  � m  ���� ���   � n  ��� � 4  ����
�� 
cobj m  ������   o  ������ 0 problem  � o      ���� "0 neweventsummary newEventSummary� �� O  �� I ������
�� .corecrel****      � null��   ��
�� 
kocl m  ����
�� 
cEvt ��~
� 
prdt K  ��		 �}

�} 
subj
 n  �� 4  ���|
�| 
cobj m  ���{�{  o  ���z�z 
0 lesson   �y
�y 
iloc n  �� 4  ���x
�x 
cobj m  ���w�w  o  ���v�v 
0 lesson   �u
�u 
offs n  �� 4  ���t
�t 
cobj m  ���s�s  o  ���r�r 
0 lesson   �q
�q 
endT [  �� l ���p�o n  �� 4  ���n 
�n 
cobj  m  ���m�m  o  ���l�l 
0 lesson  �p  �o   l ��!�k�j! ]  ��"#" 1  ���i
�i 
min # m  ���h�h P�k  �j   �g$�f
�g 
pHsR$ m  ���e
�e boovfals�f  �~   4  ���d%
�d 
cCFo% o  ���c�c "0 neweventsummary newEventSummary��  �� 
0 lesson  � o  ���b�b 0 lessons9  � &�a& l ���`�_�^�`  �_  �^  �a  � o  ���]�] 0 wintecaccount wintecAccount��  a m  GJ''�                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  _ ()( l ���\�[�Z�\  �[  �Z  ) *�Y* O  �+,+ k  -- ./. r  010 J  �X�X  1 o      �W�W 0 	allevents 	allEvents/ 232 r  	&454 6 	"676 4 	�V8
�V 
cCFo8 m  �U�U 7 E  !9:9 n  ;<; m  �T
�T 
emad< 2 �S
�S 
cAct: m   == �>>  w i n t e c5 o      �R�R 0 wintecaccount wintecAccount3 ?@? O  '�ABA k  -�CC DED r  -EFGF 6 -AHIH 2 -2�Q
�Q 
cCFoI E  5@JKJ 1  6:�P
�P 
pnamK m  ;?LL �MM  Y e a rG o      �O�O "0 winteccalendars wintecCalendarsE N�NN X  F�O�MPO k  \�QQ RSR r  \eTUT n  \aVWV 2 ]a�L
�L 
cEvtW o  \]�K�K 0 	acalendar 	aCalendarU o      �J�J 0 	theevents 	theEventsS X�IX X  f�Y�HZY r  |�[\[ o  |}�G�G 0 anevent anEvent\ n      ]^]  ;  ��^ o  }��F�F 0 	allevents 	allEvents�H 0 anevent anEventZ o  il�E�E 0 	theevents 	theEvents�I  �M 0 	acalendar 	aCalendarP o  IL�D�D "0 winteccalendars wintecCalendars�N  B o  '*�C�C 0 wintecaccount wintecAccount@ _�B_ X  �`�Aa` k  �bb cdc n  ��efe 1  ���@
�@ 
subjf o  ���?�? 0 anevent anEventd ghg r  ��iji n  ��klk 1  ���>
�> 
offsl o  ���=�= 0 anevent anEventj o      �<�< 0 startday startDayh mnm r  ��opo c  ��qrq n  ��sts m  ���;
�; 
wkdyt o  ���:�: 0 startday startDayr m  ���9
�9 
ctxtp o      �8�8 0 startday startDayn uvu l ���7�6�5�7  �6  �5  v wxw l ���4yz�4  y L F next edit -- change start date into a variable instead of a hard date   z �{{ �   n e x t   e d i t   - -   c h a n g e   s t a r t   d a t e   i n t o   a   v a r i a b l e   i n s t e a d   o f   a   h a r d   d a t ex |}| Z  �~��3~ = ����� o  ���2�2 0 startday startDay� m  ���� ���  M o n d a y r  ���� K  ���� �1��
�1 
ReDw� K  ���� �0��/
�0 
rDMo� m  ���.
�. boovtrue�/  � �-��
�- 
ReED� K  ���� �,��
�, 
pEnT� m  ���+
�+ eEDTeEDt� �*��)
�* 
peDa� o  ���(�( 0 enddate endDate�)  � �'��
�' 
tStr� m  ���� ldt     �a� � �&��
�& 
ReOi� m  ���%�% � �$��#
�$ 
ReTy� m  ���"
�" eRPTeRwp�#  � n      ��� m   �!
�! 
rREc� o  � � �  0 anevent anEvent� ��� = ��� o  �� 0 startday startDay� m  �� ���  T u e s d a y� ��� r  H��� K  B�� ���
� 
ReDw� K  �� ���
� 
rDTu� m  �
� boovtrue�  � ���
� 
ReED� K  .�� ���
� 
pEnT� m  !$�
� eEDTeEDt� ���
� 
peDa� o  '*�� 0 enddate endDate�  � ���
� 
tStr� m  14�� ldt     �c0�� ���
� 
ReOi� m  78�� � ���
� 
ReTy� m  ;>�
� eRPTeRwp�  � n      ��� m  CG�
� 
rREc� o  BC�� 0 anevent anEvent� ��� = KR��� o  KN�� 0 startday startDay� m  NQ�� ���  W e d n e s d a y� ��� r  U���� K  U��� ���
� 
ReDw� K  X^�� �
��	
�
 
rDWe� m  [\�
� boovtrue�	  � ���
� 
ReED� K  aq�� ���
� 
pEnT� m  dg�
� eEDTeEDt� ���
� 
peDa� o  jm�� 0 enddate endDate�  � ���
� 
tStr� m  tw�� ldt     �d� � � ��
�  
ReOi� m  z{���� � �����
�� 
ReTy� m  ~���
�� eRPTeRwp��  � n      ��� m  ����
�� 
rREc� o  ������ 0 anevent anEvent� ��� = ����� o  ������ 0 startday startDay� m  ���� ���  T h u r s d a y� ��� r  ����� K  ���� ����
�� 
ReDw� K  ���� �����
�� 
rDTh� m  ����
�� boovtrue��  � ����
�� 
ReED� K  ���� ����
�� 
pEnT� m  ����
�� eEDTeEDt� �����
�� 
peDa� o  ������ 0 enddate endDate��  � ����
�� 
tStr� m  ���� ldt     �eӀ� ����
�� 
ReOi� m  ������ � �����
�� 
ReTy� m  ����
�� eRPTeRwp��  � n      ��� m  ����
�� 
rREc� o  ������ 0 anevent anEvent� ��� = ����� o  ������ 0 startday startDay� m  ���� ���  F r i d a y� ���� r  ���� K  ��� ����
�� 
ReDw� K  ���� �����
�� 
rDFr� m  ����
�� boovtrue��  � ��� 
�� 
ReED� K  �� ��
�� 
pEnT m  ����
�� eEDTeEDt ����
�� 
peDa o  ������ 0 enddate endDate��    ��
�� 
tStr m  �� ldt     �g%  ��	
�� 
ReOi m   ���� 	 ��
��
�� 
ReTy
 m  ��
�� eRPTeRwp��  � n       m  ��
�� 
rREc o  ���� 0 anevent anEvent��  �3  } �� l ��������  ��  ��  ��  �A 0 anevent anEventa o  ������ 0 	allevents 	allEvents�B  , m  ���                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  �Y  ��  ��  �  l     ��������  ��  ��    l %���� I %����
�� .sysottosnull���     TEXT m  ! �  F i n i s h e d .��  ��  ��    l     ��������  ��  ��    i   " % I      ������ 0 simple_sort   �� o      ���� 0 my_list  ��  ��   k     u  !  r     "#" J     ����  # l     $����$ o      ���� 0 
index_list  ��  ��  ! %&% r    	'(' J    ����  ( l     )����) o      ���� 0 sorted_list  ��  ��  & *+* U   
 r,-, k    m.. /0/ r    121 m    33 �44  2 l     5����5 o      ���� 0 low_item  ��  ��  0 676 Y    c8��9:��8 Z   ) ^;<����; H   ) -== E  ) ,>?> l  ) *@����@ o   ) *���� 0 
index_list  ��  ��  ? o   * +���� 0 i  < k   0 ZAA BCB r   0 8DED c   0 6FGF n   0 4HIH 4   1 4��J
�� 
cobjJ o   2 3���� 0 i  I o   0 1���� 0 my_list  G m   4 5��
�� 
ctxtE o      ���� 0 	this_item  C K��K Z   9 ZLMN��L =  9 <OPO l  9 :Q����Q o   9 :���� 0 low_item  ��  ��  P m   : ;RR �SS  M k   ? FTT UVU r   ? BWXW o   ? @���� 0 	this_item  X l     Y����Y o      ���� 0 low_item  ��  ��  V Z��Z r   C F[\[ o   C D���� 0 i  \ l     ]����] o      ���� 0 low_item_index  ��  ��  ��  N ^_^ A I L`a` o   I J���� 0 	this_item  a l  J Kb����b o   J K���� 0 low_item  ��  ��  _ c��c k   O Vdd efe r   O Rghg o   O P���� 0 	this_item  h l     i����i o      ���� 0 low_item  ��  ��  f j��j r   S Vklk o   S T���� 0 i  l l     m����m o      ���� 0 low_item_index  ��  ��  ��  ��  ��  ��  ��  ��  �� 0 i  9 m    ���� : l   $n����n n    $opo m   ! #��
�� 
nmbrp n   !qrq 2   !��
�� 
cobjr o    ���� 0 my_list  ��  ��  ��  7 sts r   d huvu l  d ew����w o   d e���� 0 low_item  ��  ��  v l     x����x n      yzy  ;   f gz o   e f���� 0 sorted_list  ��  ��  t {�{ r   i m|}| l  i j~�~�}~ o   i j�|�| 0 low_item_index  �~  �}  } l     �{�z n      ���  ;   k l� l  j k��y�x� o   j k�w�w 0 
index_list  �y  �x  �{  �z  �  - l   ��v�u� l   ��t�s� n    ��� m    �r
�r 
nmbr� n   ��� 2   �q
�q 
cobj� o    �p�p 0 my_list  �t  �s  �v  �u  + ��o� L   s u�� l  s t��n�m� o   s t�l�l 0 sorted_list  �n  �m  �o   ��� l     �k�j�i�k  �j  �i  � ��h� l     �g�f�e�g  �f  �e  �h       �d�����d  � �c�b�a
�c 
pimr�b 0 simple_sort  
�a .aevtoappnull  �   � ****� �`��` �  ���� �_ �^
�_ 
vers�^  � �]��\
�] 
cobj� ��   �[
�[ 
osax�\  � �Z��
�Z 
cobj� ��   �Y 
�Y 
scpt� �X �W
�X 
vers�W  � �V�U�T���S�V 0 simple_sort  �U �R��R �  �Q�Q 0 my_list  �T  � �P�O�N�M�L�K�J�P 0 my_list  �O 0 
index_list  �N 0 sorted_list  �M 0 low_item  �L 0 i  �K 0 	this_item  �J 0 low_item_index  � �I�H3�GR
�I 
cobj
�H 
nmbr
�G 
ctxt�S vjvE�OjvE�O g��-�,Ekh�E�O Hk��-�,Ekh �� /��/�&E�O��  �E�O�E�Y �� �E�O�E�Y hY h[OY��O��6FO��6F[OY��O�� �F��E�D���C
�F .aevtoappnull  �   � ****� k    %��  *��  0��  I��  R��  X��  ���  ���  ���  ���  ���  ��� �� �� '�� 1�� t�� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� ��� �� �� �� �� �� W�� \�� ��� ��� ���  �� ��� ��� ��� ��� =�� G�� ��� �B�B  �E  �D  � �A�@�?�>�=�<�;�A 0 i  �@ 0 j  �? 
0 lesson  �> 0 room  �= 0 theevent theEvent�< 0 	acalendar 	aCalendar�; 0 anevent anEvent�M .�: 4�9 N�8 V�7�6 e�5�4�3�2�1�0�/ �.�-�,�+�* � � � � � ��) � � �"�(%�'�&�%8>Ke�$�#�"�!� �������������������� �9����
�	��g�������� ��������?Zp�������������������� ,C������t�����������������,ESlz�����������W���������������������'��%��4������X����������������������������������� ������ �������� ������������ ����D���� ��\  ������  ����������������������	��������������	'
\	@��������������������������	�	�������	���	�	�	�����
$
n��
�
�
�����������
�
�'��o����~�������������������=L���������~�}�|�{�z�y��x�w�v�u��t���s���r���q
�: .sysottosnull���     TEXT
�9 .sysodlogaskr        TEXT�8 &0 chosencalendarapp chosenCalendarApp
�7 .misccurdldt    ��� null�6  0 thecurrentdate theCurrentDate
�5 
dtxt
�4 
dstr
�3 
spac
�2 
tstr
�1 
rslt
�0 
ttxt�/ 0 thetext theText
�. 
ldt �- 0 firstday firstDay�,  �+  
�* .sysobeepnull��� ��� long�) 0 thedate theDate
�( 
prmp
�' .gtqpchltns    @   @ ns  �&  0 delayedyesorno delayedYesOrNo
�% 
cobj�$ "0 delayedfirstday delayedFirstDay
�# 
days
�" 
year�! 0 a  
�  
mnth
� 
nmbr� 0 b  
� 
day � 0 c  � 0 e  � 0 enddate endDate� 0 thelist theList
� 
txdl� 
0 tid TID
� 
ctxt� "0 thelistasstring theListAsString� "0 endofrecurrence endOfRecurrence� 0 
classnames 
classNames
� .miscactvnull��� ��� null
� 
docu
� 
pnam� "0 listofdocuments listOfDocuments�  0 chosendocument chosenDocument
� 
NmSh� 0 listofsheets listOfSheets� 0 chosensheet chosenSheet
�
 
NmTb
�	 
NmCl�  
� 
NMCv� 0 yearonebegin yearOneBegin
� 
NMCo
� 
NMad� ,0 yearonecolumnaddress yearOneColumnAddress
� .corecnte****       ****� (0 yearonecolumncount yearOneColumnCount
� 
NMRw
�  
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
�� 
iloc
�� 
offs
�� 
endT
�� 
pHsR�� 0 	allevents 	allEvents�� "0 winteccalendars wintecCalendars�� 0 startday startDay
�� 
wkdy
� 
ReDw
�~ 
rDMo
�} 
ReED
�| 
pEnT
�{ eEDTeEDt
�z 
peDa
�y 
tStr
�x 
ReOi
�w 
ReTy
�v eRPTeRwp
�u 
rREc
�t 
rDTu
�s 
rDWe
�r 
rDTh
�q 
rDFr�C&�j O�j O�E�O�j O UhZ*j E�O����,�%��,%l O��,E` O !_ a  *a _ /E` OY hW X  *j [OY��Oa _ %a %j Oa �_ �,�%_ �,%l Oa j O WhZ*j E�Oa ���,�%��,%l O��,E` O !_ a  *a _ /E` OY hW X  *j [OY��Oa _ %a %j Oa  �_ �,�%_ �,%l Oa !j Oa "a #lva $a %l &E` 'O_ 'a (k/E` 'O_ 'a )  ea *j O WhZ*j E�Oa +���,�%��,%l O��,E` O !_ a , *a _ /E` -OY hW X  *j [OY��Y hO_ _ .k E` O_ a /,E` 0O_ a 1,a 2&E` 3O_ a 4,E` 5Oa 6E` 7O_ E` 8O_ 0_ 3_ 5mvE` 9O*a :,a ;lvE[a (k/E` <Z[a (l/*a :,FZO_ 9a =&E` >O_ <*a :,FO_ 7_ >lvE` 9O*a :,a ?lvE[a (k/E` <Z[a (l/*a :,FZO_ 9a =&E` @O_ <*a :,FOjvE` AOa B�*j CO*a D-a E,E` FO) 	a Gj UO_ Fj &E` HO_ Ha (k/E` HO*a D_ H/g) 	a Ij UO*a J-a E,E` KO_ Kj &E` LO_ La (k/E` LO*a J_ L/�*a Mk/~*a Nk/a O[a P,\Za Q@1E` RO_ Ra S,a T,E` UO*a S_ U/a N-j VE` WO_ Ra X,a T,kE` RO `_ R_ Wkh  *a S_ U/a N�/a P,a Y  Y 2*a S_ Ua Z/a N�/a P,a [%*a S_ U/a N�/a P,%_ A6FOP[OY��O*a Nl/a O[a P,\Za \@1E` RO_ Ra S,a T,E` UO*a S_ U/a N-j VE` WO_ Ra X,a T,kE` RO `_ R_ Wkh  *a S_ U/a N�/a P,a Y  Y 2*a S_ Ua Z/a N�/a P,a ]%*a S_ U/a N�/a P,%_ A6FOP[OY��O)_ Ak+ ^E` _O_ _E` AOPUOPUO) 	a `j UO*a J-a E,E` KO_ Kj &E` LO_ La (k/E` LO) 	a aj UO*a J_ L/Q*a Mk/G*a Nk/a O[a P,\Za b81E` cO_ ca S,a T,E` dO_ ca X,a T,E` eO*a S_ d/a N-j VE` fOjvE` gO c_ ek_ fkh  *a S_ dk/a N�/a P,a Y 4*a S_ dk/a N�/a P,*a S_ d/a N�/a P,a =&lv_ g6FY hOP[OY��OjvE` hO �la ikh   qa Za jkh *a S�/a N�/a P,a Y J*a S�/a Nl/a P,*a S�/a Nm/a P,*a Sk/a N�/a P,*a S�/a N�/a P,a Zv_ h6FY h[OY��[OY��UUUUOjvE` kO j_ h[a la (l Vkh �a (l/a m %�a (k/a n�a (m/�a (a Z/a Zv_ k6FY "�a (k/a o�a (m/�a (a Z/a Zv_ k6F[OY��OjvE` pO �_ k[a la (l Vkh �a (m/E` qO_ qa ra Z/a s%_ qa ra t/%E` uO_ ua rk/a 2&a v *_ ua rk/a va w%_ ua rl/%a x%a =&E` yY #_ ua rk/a z%_ ua rl/%a {%a =&E` yO�a (k/�a (l/_ y�a (a Z/a Zv_ p6F[OY�TOjvE` |O @_ p[a la (l Vkh �a (l/a }%�a (a Z/%�a (k/�a (m/mv_ |6F[OY��OjvE` ~O_ |[a la (l Vkh �a (l/a  "�a (k/_ �,a �%�a (m/%lv_ ~6FY Ѣa (l/a � (�a (k/_ _ .k �,a �%�a (m/%lv_ ~6FY ��a (l/a � (�a (k/_ _ .l �,a �%�a (m/%lv_ ~6FY k�a (l/a � (�a (k/_ _ .m �,a �%�a (m/%lv_ ~6FY 8�a (l/a � *�a (k/_ _ .a Z �,a �%�a (m/%lv_ ~6FY h[OY��OjvE` �O B_ ~[a la (l Vkh �a (l/E` �O*a _ �/E` O�a (k/_ lv_ �6F[OY��OjvE` �O �_ �[a la (l Vkh �a (k/a � �a (k/�a (l/a �mv_ �6FY J G_ g[a la (l Vkh �a (k/�a (k/ �a (k/�a (l/�a (l/mv_ �6FY h[OY��[OY��OjvE` �O_ �[a la (l Vkh �a (k/a � ��a (k/a ra Z/E` �O�a (k/a ra t/E` �O�a (k/a rk/E` �O�a (k/a rl/E` �O�a (k/a rm/E` �O_ �a �%_ �%a �%_ �%a �%_ �%a �%_ �%E` �Y c�a (k/a ra Z/E` �O�a (k/a rk/E` �O�a (k/a rl/E` �O�a (k/a rm/E` �O_ �a �%_ �%a �%_ �%a �%_ �%E` �O_ ��a (l/�a (m/mv_ �6F[OY��O�a � Ua � 2*j CO) 	a �j UO*a �-a O[a E,\Za �@1 *j �UUOa �j Oa �s*j CO >k_ Aj Vkh  *a �_ Aa (�/E/j � *a �_ Aa (�/l �Y h[OY��O) 	a �j UO_ �[a la (l Vkh *a ��a (k/E/j � [*a ��a (k/E/ H*a la �a �a ��a (k/a ��a (m/a ��a (l/a ��a (l/_ �a � a �_ @a �a Z �UY ��a (k/a r-E` �O_ �a (k/a �%_ �a (l/%a �%_ �a (m/%a �%_ �a (a Z/%E` �O*a �_ �/ H*a la �a �a ��a (k/a ��a (m/a ��a (l/a ��a (l/_ �a � a �_ @a �a Z �U[OY��UO) 	a �j UOa � *a �-a E,a O[a E,\Za �@1E` �UO)a �a �/ *j �UE` �O_ E` �O_ �_ .a t E` 8O)a �a �/ _ �a �a �a �_ �a Z �UE` �O)a �a �/ !*a �_ �a �_ 8a �_ �a �_ �a � �UE` �O \k_ �j Vkh  )a �a �/ *a �_ �a (�/a �a �a Z �UE` �O)a �a �/ *a �_ �a �_ �a Z �U[OY��O_ 'a �  �)a �a �/ *j �UE` �O)a �a �/ jva �a �a �_ �a Z �UE` �O_ E` �O_ -_ .k E` �O)a �a �/ !*a �_ �a �_ �a �_ �a �_ �a � �UE` �O ;_ �[a la (l Vkh )a �a �/ *a Τa �_ �a �fa � �U[OY��OPY hOa �j Oa � *j COmj �O*j �O*j COlj �UOa � -*a �a �/ !*a �k/ *a �k/ *j �Omj �UUUUOa � �*a �a �/ }*a �k/ s*a �k/ i*a �k/ _*a �k/ U*a �a t/ I*a �k/ ?*a �k/ 5) 	a �j UO*j �Oa tj �O) 	a �j UO*j �Okj �UUUUUUUUUOa � *j COa �j �UOa �j Oa ��*a �a �/�a �a �a �l �Okj �O_ �j �Okj �Oa �a �a �kvl �Okj �OjvE` �Onl*a �k/a �k/a �k/a �k/a �k/a �-j Vkh  6*a �k/a �k/a �k/a �k/a �k/a �/a �k/a k/a,E`O_a �__ �6FOaa �a �l �Okj �O_ �j �Oaj �Okj �O*ak/aa/a	a
/ *j �UO*ak/aa/a	a/a �a i/ *j �UO*a �k/a �k/a �k/a �k/a �k/a �/ak/ak/ *j �UO*a �k/a �k/a �k/a �k/a �k/a �/ak/a �a/ *j �UOkj �O_ �j �Okj �Y hW X  hOP[OY��UUOa � *j CUOaj OjvE`Oa � �*a �a/ �aa �a �l �Okj �O_ �j �Ojj �Oa �a �a �kvl �Ojj �O �k_ �j Vkh  aj �Ojj �Oaa �a �l �Ojj �O*a �k/ 8*ak/ .*a k/a,E`O*a �k/ *ak/ *a,E`UUUUO_j �Ojj �O__lv_6F[OY�|UUOaj Oa B �*j CO*a la Dl �O*a Dk/ �*a Jk/ �*a Mk/ �*a Xk/ a*a Nk/a P,FOa*a Nl/a P,FUO �k_j Vkh  _a (�/a (k/E`O_a (�/a (l/E`O*a X�k/j � *a la Xl �Y hO*a X�k/ _*a Nk/a P,FO_*a Nl/a P,FUOP[OY��UUUUOPY�a�*j CO*a-a O[a E,\Za@1j �O*ak/a O[a -a!,\Za"@1E`#O_#k Kk_ Aj Vkh  *a_ Aa (�/E/j � "*a laa �a E_ Aa (�/la Z �Y h[OY��O_ �[a la (l Vkh *a�a (k/E/j � Y*a�a (k/E/ F*a la$a �a%�a (k/a&�a (m/a'�a (l/a(�a (l/_ �a � a)fa �a Z �UY ��a (k/a r-E` �O_ �a (k/a*%_ �a (l/%a+%_ �a (m/%a,%_ �a (a Z/%E` �O*a_ �/ F*a la$a �a%�a (k/a&�a (m/a'�a (l/a(�a (l/_ �a � a)fa �a Z �U[OY��OPUUOajvE`-O*ak/a O[a -a!,\Za.@1E`#O_# a*a-a O[a E,\Za/@1E`0O E_0[a la (l Vkh �a$-E` �O  _ �[a la (l Vkh �_-6F[OY��[OY��UO�_-[a la (l Vkh �a%,EO�a',E`1O_1a2,a =&E`1O_1a3  ;a4a5ela6a7a8a9_ 8a Za:a;a<ka=a>a ��a?,FY_1a@  ;a4aAela6a7a8a9_ 8a Za:aBa<ka=a>a ��a?,FY �_1aC  ;a4aDela6a7a8a9_ 8a Za:aEa<ka=a>a ��a?,FY �_1aF  ;a4aGela6a7a8a9_ 8a Za:aHa<ka=a>a ��a?,FY F_1aI  ;a4aJela6a7a8a9_ 8a Za:aKa<ka=a>a ��a?,FY hOP[OY��UOaLj ascr  ��ޭ