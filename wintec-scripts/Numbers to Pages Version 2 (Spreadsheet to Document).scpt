FasdUAS 1.101.10   ��   ��    k             l     ��  ��    d ^ This script grabs records from Numbers and puts them into tables in Pages, one table per page     � 	 	 �   T h i s   s c r i p t   g r a b s   r e c o r d s   f r o m   N u m b e r s   a n d   p u t s   t h e m   i n t o   t a b l e s   i n   P a g e s ,   o n e   t a b l e   p e r   p a g e   
  
 l     ��������  ��  ��        l     ��  ��    5 / Prompt user to choose excel file and grab data     �   ^   P r o m p t   u s e r   t o   c h o o s e   e x c e l   f i l e   a n d   g r a b   d a t a      l     ��  ��    . (say "Please choose the left margin size"     �   P s a y   " P l e a s e   c h o o s e   t h e   l e f t   m a r g i n   s i z e "      l     ����  r         I    ��  
�� .sysodlogaskr        TEXT  m        �   � C h o o s e   t h e   w i d t h   o f   t h e   m a r g i n   o n   t h e   l e f t .   ( d e f a u l t :   2 5 ,   m i n i m u m :   1 )  ��   
�� 
dtxt  m    ����    �� ! "
�� 
disp ! m    ��
�� stic    " �� # $
�� 
btns # J    
 % %  & ' & m     ( ( � ) )  C a n c e l '  *�� * m     + + � , ,  C o n t i n u e��   $ �� -��
�� 
dflt - m     . . � / /  C o n t i n u e��    o      ���� 0 
marginleft 
marginLeft��  ��     0 1 0 l     �� 2 3��   2 - 'say "Please choose the top margin size"    3 � 4 4 N s a y   " P l e a s e   c h o o s e   t h e   t o p   m a r g i n   s i z e " 1  5 6 5 l   + 7���� 7 r    + 8 9 8 I   '�� : ;
�� .sysodlogaskr        TEXT : m     < < � = = � C h o o s e   t h e   w i d t h   o f   t h e   m a r g i n   o n   t h e   t o p .   ( d e f a u l t :   7 0 , m i n i m u m :   1 ) ; �� > ?
�� 
dtxt > m    ���� F ? �� @ A
�� 
disp @ m    ��
�� stic    A �� B C
�� 
btns B J     D D  E F E m     G G � H H  C a n c e l F  I�� I m     J J � K K  C o n t i n u e��   C �� L��
�� 
dflt L m     # M M � N N  C o n t i n u e��   9 o      ���� 0 	margintop 	marginTop��  ��   6  O P O l     �� Q R��   Q / )say "Please choose the right margin size"    R � S S R s a y   " P l e a s e   c h o o s e   t h e   r i g h t   m a r g i n   s i z e " P  T U T l  , H V���� V r   , H W X W I  , D�� Y Z
�� .sysodlogaskr        TEXT Y m   , / [ [ � \ \ � C h o o s e   t h e   w i d t h   o f   t h e   m a r g i n   o n   t h e   r i g h t .   ( d e f a u l t :   2 5 ,   m i n i m u m :   1 ) Z �� ] ^
�� 
dtxt ] m   0 1����  ^ �� _ `
�� 
disp _ m   2 3��
�� stic    ` �� a b
�� 
btns a J   4 < c c  d e d m   4 7 f f � g g  C a n c e l e  h�� h m   7 : i i � j j  C o n t i n u e��   b �� k��
�� 
dflt k m   = @ l l � m m  C o n t i n u e��   X o      ���� 0 marginright marginRight��  ��   U  n o n l     ��������  ��  ��   o  p q p l     �� r s��   r 5 /set marginLeftRightMatchChoice to {true, false}    s � t t ^ s e t   m a r g i n L e f t R i g h t M a t c h C h o i c e   t o   { t r u e ,   f a l s e } q  u v u l     �� w x��   w G Asay "Do you want the left and right margin sizes to be the same?"    x � y y � s a y   " D o   y o u   w a n t   t h e   l e f t   a n d   r i g h t   m a r g i n   s i z e s   t o   b e   t h e   s a m e ? " v  z { z l     �� | }��   | � �set marginLeftRightMatch to choose from list marginLeftRightMatchChoice with prompt "Do you want the margin for the left side and right side to match?" default items {false}    } � ~ ~Z s e t   m a r g i n L e f t R i g h t M a t c h   t o   c h o o s e   f r o m   l i s t   m a r g i n L e f t R i g h t M a t c h C h o i c e   w i t h   p r o m p t   " D o   y o u   w a n t   t h e   m a r g i n   f o r   t h e   l e f t   s i d e   a n d   r i g h t   s i d e   t o   m a t c h ? "   d e f a u l t   i t e m s   { f a l s e } {   �  l     �� � ���   �   promput user for images    � � � � 0   p r o m p u t   u s e r   f o r   i m a g e s �  � � � l     �� � ���   � 3 -say "Please tell me where the wintec logo is"    � � � � Z s a y   " P l e a s e   t e l l   m e   w h e r e   t h e   w i n t e c   l o g o   i s " �  � � � l  I P ����� � I  I P�� ���
�� .sysodlogaskr        TEXT � l  I L ����� � m   I L � � � � � @ S e l e c t   t h e   W i n t e c   l o g o   t o   i m p o r t��  ��  ��  ��  ��   �  � � � l  Q ` ����� � r   Q ` � � � l  Q \ ����� � I  Q \���� �
�� .sysostdfalis    ��� null��   � �� ���
�� 
prmp � m   U X � � � � � B S e l e c t   t h e   W i n t e c   L o g o   t o   i m p o r t :��  ��  ��   � o      ���� 0 
logowintec 
logoWintec��  ��   �  � � � l     �� � ���   � 3 -say "Please tell me where the jinhua logo is"    � � � � Z s a y   " P l e a s e   t e l l   m e   w h e r e   t h e   j i n h u a   l o g o   i s " �  � � � l  a h ����� � I  a h�� ���
�� .sysodlogaskr        TEXT � l  a d ����� � m   a d � � � � � X S e l e c t   t h e   J i n h u a   P o l y t e c h n i c   l o g o   t o   i m p o r t��  ��  ��  ��  ��   �  � � � l  i x ����� � r   i x � � � l  i t ����� � I  i t���� �
�� .sysostdfalis    ��� null��   � �� ���
�� 
prmp � m   m p � � � � � B S e l e c t   t h e   J i n h u a   L o g o   t o   i m p o r t :��  ��  ��   � o      ���� 0 
logojinhua 
logoJinhua��  ��   �  � � � l     ��������  ��  ��   �  � � � l  y � ����� � r   y � � � � n   y ~ � � � 1   z ~��
�� 
ttxt � o   y z���� 0 
marginleft 
marginLeft � o      ���� 0 
leftmargin 
leftMargin��  ��   �  � � � l  � � ����� � r   � � � � � n   � � � � � 1   � ���
�� 
ttxt � o   � ����� 0 	margintop 	marginTop � o      ���� 0 	topmargin 	topMargin��  ��   �  � � � l  � � ����� � r   � � � � � n   � � � � � 1   � ���
�� 
ttxt � o   � ����� 0 marginright marginRight � o      ���� 0 rightmargin rightMargin��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � F @set logo to (choose file with prompt "Please select logo file:")    � � � � � s e t   l o g o   t o   ( c h o o s e   f i l e   w i t h   p r o m p t   " P l e a s e   s e l e c t   l o g o   f i l e : " ) �  � � � l  � � ����� � r   � � � � � J   � �����   � o      ���� 0 rowdata rowData��  ��   �  � � � l  �� ����� � O   �� � � � k   �� � �  � � � l  � ��� � ���   � > 8	tell me to say "Please tell me where the excel file is"    � � � � p 	 t e l l   m e   t o   s a y   " P l e a s e   t e l l   m e   w h e r e   t h e   e x c e l   f i l e   i s " �  � � � I  � ��� ���
�� .sysodlogaskr        TEXT � m   � � � � � � � H P l e a s e   c h o o s e   t h e   i n p u t   f i l e   ( E x c e l )��   �  � � � I  � ��� ���
�� .aevtodocnull  �    alis � l  � � ����� � I  � ����� �
�� .sysostdfalis    ��� null��   � �� ���
�� 
prmp � m   � � � � � � � : P l e a s e   c h o o s e   t h e   i n p u t   f i l e :��  ��  ��  ��   �  ��� � O   �� � � � O   �� � � � O   �� � � � k   �� � �  � � � r   � � � � � \   � � � � � l  � � ����� � I  � ��� ���
�� .corecnte****       **** � 2  � ��
� 
NMRw��  ��  ��   � m   � ��~�~  � o      �}�} 0 recordcount recordCount �  ��| � Y   �� ��{ � ��z � O   �� �  � r  � J  �  n   1  �y
�y 
NMCv 4  �x
�x 
NmCl m  
�w�w  	
	 n   1  �v
�v 
NMCv 4  �u
�u 
NmCl m  �t�t 
  n  ( 1  $(�s
�s 
NMCv 4  $�r
�r 
NmCl m  "#�q�q   n  (4 1  04�p
�p 
NMCv 4  (0�o
�o 
NmCl m  ,/�n�n 	  n  4@ 1  <@�m
�m 
NMCv 4  4<�l
�l 
NmCl m  8;�k�k 
  n  @L  1  HL�j
�j 
NMCv  4  @H�i!
�i 
NmCl! m  DG�h�h  "#" n  LX$%$ 1  TX�g
�g 
NMCv% 4  LT�f&
�f 
NmCl& m  PS�e�e # '(' n  Xd)*) 1  `d�d
�d 
NMCv* 4  X`�c+
�c 
NmCl+ m  \_�b�b ( ,-, n  dp./. 1  lp�a
�a 
NMCv/ 4  dl�`0
�` 
NmCl0 m  hk�_�_ - 121 n  p|343 1  x|�^
�^ 
NMCv4 4  px�]5
�] 
NmCl5 m  tw�\�\ 2 6�[6 n  |�787 1  ���Z
�Z 
NMCv8 4  |��Y9
�Y 
NmCl9 m  ���X�X �[   n      :;:  ;  ��; o  ���W�W 0 rowdata rowData  4   ��V<
�V 
NMRw< o  �U�U 0 i  �{ 0 i   � m   � ��T�T  � I  � ��S=�R
�S .corecnte****       ****= 2  � ��Q
�Q 
NMRw�R  �z  �|   � 4   � ��P>
�P 
NmTb> m   � ��O�O  � 4   � ��N?
�N 
NmSh? m   � ��M�M  � 4   � ��L@
�L 
docu@ m   � ��K�K ��   � m   � �AA�                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  ��  ��   � BCB l     �J�I�H�J  �I  �H  C DED l     �G�F�E�G  �F  �E  E FGF l     �D�C�B�D  �C  �B  G HIH l     �AJK�A  J I CCreate new pages document and create the number of pages and tables   K �LL � C r e a t e   n e w   p a g e s   d o c u m e n t   a n d   c r e a t e   t h e   n u m b e r   o f   p a g e s   a n d   t a b l e sI MNM l ��O�@�?O r  ��PQP m  ���>�>cQ o      �=�= 0 maxpagewidth maxPageWidth�@  �?  N RSR l �
T�<�;T O  �
UVU k  �	WW XYX I ���:�9�8
�: .miscactvnull��� ��� null�9  �8  Y Z[Z I ���7�6\
�7 .corecrel****      � null�6  \ �5]�4
�5 
kocl] m  ���3
�3 
docu�4  [ ^�2^ O  �	_`_ k  �aa bcb Y  ��d�1ef�0d I ���/�.g
�/ .corecrel****      � null�.  g �-h�,
�- 
koclh m  ���+
�+ 
cPag�,  �1 0 i  e m  ���*�* f \  ��iji o  ���)�) 0 recordcount recordCountj m  ���(�( �0  c klk l ���'�&�%�'  �&  �%  l mnm Y  �lo�$pq�#o O  �grsr k  �ftt uvu l ���"wx�"  w   import the image   x �yy "   i m p o r t   t h e   i m a g ev z{z I �.�!� |
�! .corecrel****      � null�   | �}~
� 
kocl} m  �
� 
imag~ ��
� 
prdt K  	(�� ���
� 
file� o  �� 0 
logowintec 
logoWintec� ���
� 
sith� m  �� (� ���
� 
sipo� J  $�� ��� o  �� 0 
leftmargin 
leftMargin� ��� \  "��� o  �� 0 	topmargin 	topMargin� m  !�� 2�  �  �  { ��� I /f���
� .corecrel****      � null�  � ���
� 
kocl� m  36�
� 
imag� ���
� 
prdt� K  9`�� �
��
�
 
file� o  <?�	�	 0 
logojinhua 
logoJinhua� ���
� 
sith� m  BE�� (� ���
� 
sipo� J  H\�� ��� \  HS��� \  HO��� o  HK�� 0 maxpagewidth maxPageWidth� m  KN�� (� o  OR�� 0 rightmargin rightMargin� ��� \  SZ��� o  SV� �  0 	topmargin 	topMargin� m  VY���� 2�  �  �  �  s 4  �����
�� 
cPag� o  ������ 0 i  �$ 0 i  p m  ������ q I �������
�� .corecnte****       ****� 2 ����
�� 
cPag��  �#  n ��� l mm��������  ��  ��  � ��� Y  m�������� O  ��� k  � �� ��� I ������
�� .corecrel****      � null� m  ����
�� 
NmTb� �����
�� 
prdt� K  ���� ����
�� 
NmTc� m  ������ � �����
�� 
NmTr� l �������� I �������
�� .corecnte****       ****� n  ����� 2 ����
�� 
cobj� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 0 rowdata rowData��  ��  ��  ��  ��  � ���� O  � ��� k  ���� ��� r  ����� J  ���� ��� o  ������ 0 
leftmargin 
leftMargin� ���� [  ����� o  ������ 0 	topmargin 	topMargin� l �������� ]  ����� m  ������� l �������� \  ����� o  ������ 0 i  � m  ������ ��  ��  ��  ��  ��  � 1  ����
�� 
sipo� ��� r  ����� m  ������ _� o      ���� $0 firstcolumnwidth firstColumnWidth� ��� r  ����� o  ������ $0 firstcolumnwidth firstColumnWidth� n      ��� 1  ����
�� 
sitw� 4 �����
�� 
NMCo� m  ������ � ���� r  ����� l �������� \  ����� \  ����� \  ����� o  ������ 0 maxpagewidth maxPageWidth� o  ������ 0 
leftmargin 
leftMargin� o  ������ $0 firstcolumnwidth firstColumnWidth� o  ������ 0 rightmargin rightMargin��  ��  � n      ��� 1  ����
�� 
sitw� 4 �����
�� 
NMCo� m  ������ ��  � 4  �����
�� 
NmTb� m  ������ ��  � 4  ����
�� 
cPag� o  ������ 0 i  �� 0 i  � m  pq���� � I qz�����
�� .corecnte****       ****� 2 qv��
�� 
cPag��  ��  � ��� l ��������  ��  ��  � ��� l ������  � ! merge the first row / Title   � ��� 6 m e r g e   t h e   f i r s t   r o w   /   T i t l e� ��� l  ������  � � }
		repeat with i from 1 to count of tables
			tell table i
				tell row 1
					merge
				end tell
			end tell
		end repeat
		   � ��� � 
 	 	 r e p e a t   w i t h   i   f r o m   1   t o   c o u n t   o f   t a b l e s 
 	 	 	 t e l l   t a b l e   i 
 	 	 	 	 t e l l   r o w   1 
 	 	 	 	 	 m e r g e 
 	 	 	 	 e n d   t e l l 
 	 	 	 e n d   t e l l 
 	 	 e n d   r e p e a t 
 	 	� ���� l ��������  ��  ��  ��  ` 4  �����
�� 
docu� m  ������ �2  V m  ����|                                                                                  page  alis      Macintosh HD                   BD ����	Pages.app                                                      ����            ����  
 cu             Applications  /:Applications:Pages.app/    	 P a g e s . a p p    M a c i n t o s h   H D  Applications/Pages.app  / ��  �<  �;  S ��� l     ��������  ��  ��  � ��� l     ��� ��  �   Add a Title to each page     � 2   A d d   a   T i t l e   t o   e a c h   p a g e�  l     ��������  ��  ��    l ���� r   \  	
	 \   o  ���� 0 maxpagewidth maxPageWidth o  ���� 0 
leftmargin 
leftMargin
 o  ���� 0 rightmargin rightMargin o      ���� 0 
tablewidth 
tableWidth��  ��    l     ��������  ��  ��    l ����� O  � O  !� Y  *����� O  <� k  E�  I Ev��
�� .corecrel****      � null m  EH��
�� 
shtx �� ��
�� 
prdt  K  Kr!! ��"#
�� 
sipo" J  Nb$$ %&% [  NY'(' o  NQ���� 0 
leftmargin 
leftMargin( ]  QX)*) o  QT���� 0 
tablewidth 
tableWidth* m  TW++ ?�(�\)& ,��, \  Y`-.- o  Y\���� 0 	topmargin 	topMargin. m  \_���� (��  # ��/0
�� 
sith/ m  eh���� -0 ��1��
�� 
sitw1 m  kn�������  ��   232 O  w�454 k  ��66 787 r  ��9:9 m  ��;; �<< �                     J i n h u a   P o l y t e c h n i c   W i n t e c   I n t e r n a t i o n a l   C o l l e g e 
           W i n t e c   E n g l i s h   T e a c h e r s '   O n l i n e   C l a s s r o o m   A c t i v i t y   I d e a s: 1  ����
�� 
pDTx8 =��= r  ��>?> m  ������ ? n      @A@ 1  ����
�� 
ptszA 1  ����
�� 
pDTx��  5 4  w}��B
�� 
shtxB m  {|���� 3 CDC l ����������  ��  ��  D EFE l  ����GH��  G
				make text item with properties {position:{leftMargin + tableWidth * 0.11, topMargin - 20}, height:25, width:400}
				tell text item 2
					set object text to "Wintec English Teachers' Online Classroom Activity Ideas"
					set size of object text to 16
				end tell
				   H �II$ 
 	 	 	 	 m a k e   t e x t   i t e m   w i t h   p r o p e r t i e s   { p o s i t i o n : { l e f t M a r g i n   +   t a b l e W i d t h   *   0 . 1 1 ,   t o p M a r g i n   -   2 0 } ,   h e i g h t : 2 5 ,   w i d t h : 4 0 0 } 
 	 	 	 	 t e l l   t e x t   i t e m   2 
 	 	 	 	 	 s e t   o b j e c t   t e x t   t o   " W i n t e c   E n g l i s h   T e a c h e r s '   O n l i n e   C l a s s r o o m   A c t i v i t y   I d e a s " 
 	 	 	 	 	 s e t   s i z e   o f   o b j e c t   t e x t   t o   1 6 
 	 	 	 	 e n d   t e l l 
 	 	 	 	F JKJ l ����������  ��  ��  K L��L l ����������  ��  ��  ��   4  <B�M
� 
cPagM o  @A�~�~ 0 i  �� 0 i   m  -.�}�}  I .7�|N�{
�| .corecnte****       ****N 2 .3�z
�z 
cPag�{  ��   4  !'�yO
�y 
docuO m  %&�x�x  m  PP|                                                                                  page  alis      Macintosh HD                   BD ����	Pages.app                                                      ����            ����  
 cu             Applications  /:Applications:Pages.app/    	 P a g e s . a p p    M a c i n t o s h   H D  Applications/Pages.app  / ��  ��  ��   QRQ l     �w�v�u�w  �v  �u  R STS l     �tUV�t  U   Set the Field labels   V �WW *   S e t   t h e   F i e l d   l a b e l sT XYX l �6Z�s�rZ O  �6[\[ O  �5]^] k  �4__ `a` r  ��bcb J  ��dd efe m  ��gg �hh  A c t i v i t y   T i t l ef iji m  ��kk �ll 
 L e v e lj mnm m  ��oo �pp " C u t t i n g   E d g e   U n i tn qrq m  ��ss �tt  L a n g u a g e   F o c u sr uvu m  ��ww �xx " A c t i v i t y   D u r a t i o nv yzy m  ��{{ �||  O v e r v i e wz }~} m  �� ���  I n s t r u c t i o n s~ ��� m  ���� ���  A d a p t a t i o n s� ��� m  ���� ���  S u b m i t t e d   b y� ��� m  ���� ��� " C o n t r i b u t o r   E m a i l� ��q� m  ���� ��� $ M a t e r i a l s   R e q u i r e d�q  c o      �p�p 0 fieldlabels fieldLabelsa ��o� Y  �4��n���m� O  �/��� O  �.��� Y   -��l���k� r  (��� n  ��� 4  �j�
�j 
cobj� o  �i�i 0 j  � o  �h�h 0 fieldlabels fieldLabels� n      ��� 1  #'�g
�g 
NMCv� n  #��� 4  #�f�
�f 
NmCl� m  !"�e�e � 4  �d�
�d 
NMRw� o  �c�c 0 j  �l 0 j  � m  �b�b � I �a��`
�a .corecnte****       ****� o  �_�_ 0 fieldlabels fieldLabels�`  �k  � 4  ���^�
�^ 
NmTb� m  ���]�] � 4  ���\�
�\ 
cPag� o  ���[�[ 0 i  �n 0 i  � m  ���Z�Z � I ���Y��X
�Y .corecnte****       ****� 2 ���W
�W 
cPag�X  �m  �o  ^ 4  ���V�
�V 
docu� m  ���U�U \ m  ����|                                                                                  page  alis      Macintosh HD                   BD ����	Pages.app                                                      ����            ����  
 cu             Applications  /:Applications:Pages.app/    	 P a g e s . a p p    M a c i n t o s h   H D  Applications/Pages.app  / ��  �s  �r  Y ��� l     �T�S�R�T  �S  �R  � ��� l     �Q�P�O�Q  �P  �O  � ��� l     �N���N  � &   Fill in the data for each table   � ��� @   F i l l   i n   t h e   d a t a   f o r   e a c h   t a b l e� ��� l 7���M�L� O  7���� O  =���� Y  F���K���J� O  X���� O  a���� Y  j���I���H� O  |���� k  ���� ��� r  ����� n  ����� 4  ���G�
�G 
cobj� o  ���F�F 0 j  � n  ����� 4  ���E�
�E 
cobj� l ����D�C� [  ����� o  ���B�B 0 i  � m  ���A�A �D  �C  � o  ���@�@ 0 rowdata rowData� n      ��� 1  ���?
�? 
NMCv� 4  ���>�
�> 
NmCl� m  ���=�= � ��<� r  ����� m  ���;
�; tAHTalft� 1  ���:
�: 
texA�<  � 4  |��9�
�9 
NMRw� o  ���8�8 0 j  �I 0 j  � m  mn�7�7 � l nw��6�5� l nw��4�3� I nw�2��1
�2 .corecnte****       ****� 2 ns�0
�0 
NMRw�1  �4  �3  �6  �5  �H  � 4  ag�/�
�/ 
NmTb� m  ef�.�. � 4  X^�-�
�- 
cPag� o  \]�,�, 0 i  �K 0 i  � m  IJ�+�+ � I JS�*��)
�* .corecnte****       ****� 2 JO�(
�( 
cPag�)  �J  � 4  =C�'�
�' 
docu� m  AB�&�& � m  7:��|                                                                                  page  alis      Macintosh HD                   BD ����	Pages.app                                                      ����            ����  
 cu             Applications  /:Applications:Pages.app/    	 P a g e s . a p p    M a c i n t o s h   H D  Applications/Pages.app  / ��  �M  �L  � ��%� l     �$�#�"�$  �#  �"  �%       �!���!  � � 
�  .aevtoappnull  �   � ****� �������
� .aevtoappnull  �   � ****� k    ���  ��  5��  T��  ���  ���  ���  ���  ���  ���  ���  ���  ��� M�� R�� �� �� X�� ���  �  �  � ��� 0 i  � 0 j  � h ����� ( +� .��� <� G J M� [ f i l� �� ���
 � ��	�����A � ����� �����������������������������������������������������������������������������+����;������gkosw{����������
� 
dtxt� 
� 
disp
� stic   
� 
btns
� 
dflt� 
� .sysodlogaskr        TEXT� 0 
marginleft 
marginLeft� F� 0 	margintop 	marginTop� 0 marginright marginRight
� 
prmp
� .sysostdfalis    ��� null�
 0 
logowintec 
logoWintec�	 0 
logojinhua 
logoJinhua
� 
ttxt� 0 
leftmargin 
leftMargin� 0 	topmargin 	topMargin� 0 rightmargin rightMargin� 0 rowdata rowData
� .aevtodocnull  �    alis
� 
docu
� 
NmSh
�  
NmTb
�� 
NMRw
�� .corecnte****       ****�� 0 recordcount recordCount
�� 
NmCl�� 
�� 
NMCv�� �� 	�� 
�� �� �� �� �� �� ��c�� 0 maxpagewidth maxPageWidth
�� .miscactvnull��� ��� null
�� 
kocl
�� .corecrel****      � null
�� 
cPag
�� 
imag
�� 
prdt
�� 
file
�� 
sith�� (
�� 
sipo�� 2
�� 
NmTc
�� 
NmTr
�� 
cobj���� _�� $0 firstcolumnwidth firstColumnWidth
�� 
NMCo
�� 
sitw�� 0 
tablewidth 
tableWidth
�� 
shtx�� -���
�� 
pDTx�� 
�� 
ptsz�� 0 fieldlabels fieldLabels
�� tAHTalft
�� 
texA����������lv��� E�O�������a lv�a � E` Oa �����a a lv�a � E` Oa j O*a a l E` Oa j O*a a l E` O�a  ,E` !O_ a  ,E` "O_ a  ,E` #OjvE` $Oa % �a &j O*a a 'l j (O*a )k/ �*a *k/ �*a +k/ �*a ,-j -lE` .O �k*a ,-j -kh  *a ,�/ �*a /a 0/a 1,*a /a 2/a 1,*a /�/a 1,*a /a 3/a 1,*a /a 4/a 1,*a /a 5/a 1,*a /a 6/a 1,*a /a 7/a 1,*a /a 8/a 1,*a /a 9/a 1,*a /a :/a 1,a 5v_ $6FU[OY�eUUUUOa ;E` <Oa =a*j >O*a ?a )l @O*a )k/E k_ .kkh  *a ?a Al @[OY��O �k*a A-j -kh  *a A�/ i*a ?a Ba Ca D_ a Ea Fa G_ !_ "a Hlva 0a 9 @O*a ?a Ba Ca D_ a Ea Fa G_ <a F_ #_ "a Hlva 0a 9 @U[OY��O �k*a A-j -kh  *a A�/ za +a Ca Ila J_ $a Kk/a K-j -a 9l @O*a +k/ J_ !_ "a L�k lv*a G,FOa ME` NO_ N*a Ok/a P,FO_ <_ !_ N_ #*a Ol/a P,FUU[OY�xOPUUO_ <_ !_ #E` QOa = �*a )k/ x uk*a A-j -kh  *a A�/ Wa Ra Ca G_ !_ Qa S _ "a Flva Ea Ta Pa Ua 0l @O*a Rk/ a V*a W,FOa X*a W,a Y,FUOPU[OY��UUOa = �*a )k/ �a Za [a \a ]a ^a _a `a aa ba ca da 5vE` eO Wk*a A-j -kh  *a A�/ 9*a +k/ / ,k_ ej -kh _ ea K�/*a ,�/a /k/a 1,F[OY��UU[OY��UUOa = |*a )k/ r ok*a A-j -kh  *a A�/ Q*a +k/ G Dk*a ,-j -kh *a ,�/ &_ $a K�k/a K�/*a /l/a 1,FOa f*a g,FU[OY��UU[OY��UU ascr  ��ޭ