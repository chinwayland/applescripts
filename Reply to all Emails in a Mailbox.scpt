FasdUAS 1.101.10   ��   ��    k             l     ��  ��    k e Loop through each message in currently selected mailbox to grab sender's info and put into variables     � 	 	 �   L o o p   t h r o u g h   e a c h   m e s s a g e   i n   c u r r e n t l y   s e l e c t e d   m a i l b o x   t o   g r a b   s e n d e r ' s   i n f o   a n d   p u t   i n t o   v a r i a b l e s   
  
 l    � ����  O     �    k    �       r        J    ����    o      ����  0 sendercontents senderContents      r   	     J   	 ����    o      ���� 0 senderinfos senderInfos      r        J    ����    o      ���� "0 senderaddresses senderAddresses      r        J    ����    o      ����  0 sendersubjects senderSubjects     !   r     " # " J    ����   # o      ���� 60 senderattachmentfilenames senderAttachmentFileNames !  $ % $ r    * & ' & n    ( ( ) ( 1   & (��
�� 
pnam ) n    & * + * m   $ &��
�� 
mbxp + n    $ , - , 4   ! $�� .
�� 
cobj . m   " #����  - l   ! /���� / e    ! 0 0 1    !��
�� 
slct��  ��   ' o      ���� 0 mailboxname mailboxName %  1 2 1 r   + : 3 4 3 n   + 8 5 6 5 1   6 8��
�� 
pnam 6 n   + 6 7 8 7 m   4 6��
�� 
mact 8 n   + 4 9 : 9 m   2 4��
�� 
mbxp : n   + 2 ; < ; 4   / 2�� =
�� 
cobj = m   0 1����  < l  + / >���� > e   + / ? ? 1   + /��
�� 
slct��  ��   4 o      ���� 0 accountname accountName 2  @ A @ l  ; ;��������  ��  ��   A  B�� B O   ; � C D C O   B � E F E O   I � G H G Y   O � I�� J K L I O   ] � M N M k   d � O O  P Q P r   d j R S R 1   d g��
�� 
ctnt S n       T U T  ;   h i U o   g h����  0 sendercontents senderContents Q  V W V r   k s X Y X 1   k p��
�� 
sndr Y n       Z [ Z  ;   q r [ o   p q���� 0 senderinfos senderInfos W  \ ] \ r   t � ^ _ ^ I  t }�� `��
�� .emaleauanull���     ctxt ` c   t y a b a o   t u���� 0 senderinfos senderInfos b m   u x��
�� 
TEXT��   _ n       c d c  ;   ~  d o   } ~���� "0 senderaddresses senderAddresses ]  e�� e l   � ��� f g��   f						set end of senderSubjects to subject						if mail attachments exists then							set end of senderAttachmentFileNames to name of item 1 of mail attachments						else							set end of senderAttachmentFileNames to "No Attachments"						end if						    g � h h  	 	 	 	 	 	 s e t   e n d   o f   s e n d e r S u b j e c t s   t o   s u b j e c t  	 	 	 	 	 	 i f   m a i l   a t t a c h m e n t s   e x i s t s   t h e n  	 	 	 	 	 	 	 s e t   e n d   o f   s e n d e r A t t a c h m e n t F i l e N a m e s   t o   n a m e   o f   i t e m   1   o f   m a i l   a t t a c h m e n t s  	 	 	 	 	 	 e l s e  	 	 	 	 	 	 	 s e t   e n d   o f   s e n d e r A t t a c h m e n t F i l e N a m e s   t o   " N o   A t t a c h m e n t s "  	 	 	 	 	 	 e n d   i f  	 	 	 	 	 	��   N 4   ] a�� i
�� 
cobj i o   _ `���� 0 messagenumber messageNumber�� 0 messagenumber messageNumber J l  R W j���� j I  R W������
�� .corecnte****       ****��  ��  ��  ��   K m   W X����  L m   X Y������ H 2  I L��
�� 
mssg F 4   B F�� k
�� 
mbxp k o   D E���� 0 mailboxname mailboxName D 4   ; ?�� l
�� 
mact l o   = >���� 0 accountname accountName��    m      m m�                                                                                  emal  alis    (  Macintosh HD                   BD ����Mail.app                                                       ����            ����  
 cu             Applications  /:System:Applications:Mail.app/     M a i l . a p p    M a c i n t o s h   H D  System/Applications/Mail.app  / ��  ��  ��     n o n l     ��������  ��  ��   o  p q p l     �� r s��   r R L Get English first names from Numbers document and put into list senderNames    s � t t �   G e t   E n g l i s h   f i r s t   n a m e s   f r o m   N u m b e r s   d o c u m e n t   a n d   p u t   i n t o   l i s t   s e n d e r N a m e s q  u v u l  � � w���� w r   � � x y x J   � �����   y o      ���� 0 cellhit cellHit��  ��   v  z { z l  � � |���� | r   � � } ~ } J   � �����   ~ o      ���� 0 sendernames senderNames��  ��   {   �  l  �` ����� � X   �` ��� � � O   �[ � � � O   �Z � � � k   �Y � �  � � � Y   � ��� � ��� � O   � � � � O   � � � � Z   � � ����� � >  � � � � � l  � � ����� � 6  � � � � � 2   � ���
�� 
NmCl � =  � � � � � 1   � ���
�� 
NMCv � o   � ����� 0 senderaddress senderAddress��  ��   � J   � �����   � k   � � �  � � � r   � � � � l  � ����� � 6  � � � � 2   � ���
�� 
NmCl � =  � � � � 1   � ���
�� 
NMCv � o   ���� 0 senderaddress senderAddress��  ��   � o      ���� 0 cellhit cellHit �  ��� � r  	 � � � o  	
���� 0 sheetnumber sheetNumber � o      ���� 0 
sheetindex 
sheetIndex��  ��  ��   � 4   � ��� �
�� 
NmTb � m   � �����  � 4   � ��� �
�� 
NmSh � o   � ����� 0 sheetnumber sheetNumber�� 0 sheetnumber sheetNumber � m   � �����  � m   � ����� ��   �  � � � r  , � � � n  ( � � � 1  $(��
�� 
NMad � n  $ � � � m   $��
�� 
NMRw � n    � � � 4   �� �
�� 
cobj � m  ����  � o  ���� 0 cellhit cellHit � o      ���� 0 
rowaddress 
rowAddress �  ��� � O  -Y � � � O  8X � � � r  AW � � � n  AR � � � 1  NR��
�� 
NMCv � n  AN � � � 4 IN�� �
�� 
NmCl � m  LM����  � 4  AI�� �
�� 
NMRw � o  EH���� 0 
rowaddress 
rowAddress � n       � � �  ;  UV � o  RU���� 0 sendernames senderNames � 4  8>�� �
�� 
NmTb � m  <=����  � 4  -5�� �
�� 
NmSh � o  14���� 0 
sheetindex 
sheetIndex��   � 4   � ��� �
�� 
docu � m   � �����  � m   � � � ��                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  �� 0 senderaddress senderAddress � o   � ����� "0 senderaddresses senderAddresses��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � 4 .send a message to each sender from above steps    � � � � \ s e n d   a   m e s s a g e   t o   e a c h   s e n d e r   f r o m   a b o v e   s t e p s �  � � � l a  ����� � Y  a  ��� � �� � O  o� � � � k  s� � �  � � � r  s � � � n  s{ � � � 1  y{�~
�~ 
ctnt � 4  sy�} �
�} 
situ � m  wx�|�|  � o      �{�{ $0 signaturecontent signatureContent �  � � � r  �� � � � I ���z�y �
�z .corecrel****      � null�y   � �x ��w
�x 
kocl � m  ���v
�v 
bcke�w   � o      �u�u (0 newoutgoingmessage newOutGoingMessage �  ��t � O  �� � � � k  �� � �  � � � r  �� � � � m  ���s
�s boovtrue � 1  ���r
�r 
pvis �  � � � I ���q�p �
�q .corecrel****      � null�p   � �o � �
�o 
kocl � m  ���n
�n 
trcp � �m � �
�m 
insh � n  �� � � �  ;  �� � 2 ���l
�l 
trcp � �k ��j
�k 
prdt � K  �� � � �i ��h
�i 
radd � n  �� � � � 4  ���g �
�g 
cobj � o  ���f�f 0 messagenumber messageNumber � o  ���e�e "0 senderaddresses senderAddresses�h  �j   �  � � � r  �� � � � b  ��   m  �� �  R e :   n  �� 4  ���d
�d 
cobj o  ���c�c 0 messagenumber messageNumber o  ���b�b  0 sendersubjects senderSubjects � 1  ���a
�a 
subj �  r  ��	
	 b  �� b  �� b  �� b  �� b  �� b  �� m  �� �  T h a n k   y o u   n  �� 4  ���`
�` 
cobj o  ���_�_ 0 messagenumber messageNumber o  ���^�^ 0 sendernames senderNames m  �� �  !   
 o  ���]�] $0 signaturecontent signatureContent m  �� �   	 
 	 Y o u   w r o t e :   
 	 n  �� !  4  ���\"
�\ 
cobj" o  ���[�[ 0 messagenumber messageNumber! o  ���Z�Z  0 sendercontents senderContents m  ��## �$$  

 1  ���Y
�Y 
ctnt %&% l ���X'(�X  ' 7 1& item messageNumber of senderAttachmentFileNames   ( �)) b &   i t e m   m e s s a g e N u m b e r   o f   s e n d e r A t t a c h m e n t F i l e N a m e s& *�W* I ���V�U�T
�V .emsgsendnull���     bcke�U  �T  �W   � o  ���S�S (0 newoutgoingmessage newOutGoingMessage�t   � m  op++�                                                                                  emal  alis    (  Macintosh HD                   BD ����Mail.app                                                       ����            ����  
 cu             Applications  /:System:Applications:Mail.app/     M a i l . a p p    M a c i n t o s h   H D  System/Applications/Mail.app  / ��  �� 0 messagenumber messageNumber � m  de�R�R  � I ej�Q,�P
�Q .corecnte****       ****, o  ef�O�O "0 senderaddresses senderAddresses�P  �  ��  ��   � -�N- l     �M�L�K�M  �L  �K  �N       
�J./0123456�J  . �I�H�G�F�E�D�C�B
�I .aevtoappnull  �   � ****�H  0 sendercontents senderContents�G 0 senderinfos senderInfos�F "0 senderaddresses senderAddresses�E  0 sendersubjects senderSubjects�D 60 senderattachmentfilenames senderAttachmentFileNames�C 0 mailboxname mailboxName�B 0 accountname accountName/ �A7�@�?89�>
�A .aevtoappnull  �   � ****7 k     ::  
;;  u<<  z==  >>  ��=�=  �@  �?  8 �<�;�:�< 0 messagenumber messageNumber�; 0 senderaddress senderAddress�: 0 sheetnumber sheetNumber9 4 m�9�8�7�6�5�4�3�2�1�0�/�.�-�,�+�*�)�(�'�&�% ��$�#�"�!� ?�����������������#��9  0 sendercontents senderContents�8 0 senderinfos senderInfos�7 "0 senderaddresses senderAddresses�6  0 sendersubjects senderSubjects�5 60 senderattachmentfilenames senderAttachmentFileNames
�4 
slct
�3 
cobj
�2 
mbxp
�1 
pnam�0 0 mailboxname mailboxName
�/ 
mact�. 0 accountname accountName
�- 
mssg
�, .corecnte****       ****
�+ 
ctnt
�* 
sndr
�) 
TEXT
�( .emaleauanull���     ctxt�' 0 cellhit cellHit�& 0 sendernames senderNames
�% 
kocl
�$ 
docu�# 
�" 
NmSh
�! 
NmTb
�  
NmCl?  
� 
NMCv� 0 
sheetindex 
sheetIndex
� 
NMRw
� 
NMad� 0 
rowaddress 
rowAddress
� 
situ� $0 signaturecontent signatureContent
� 
bcke
� .corecrel****      � null� (0 newoutgoingmessage newOutGoingMessage
� 
pvis
� 
trcp
� 
insh
� 
prdt
� 
radd� 
� 
subj
� .emsgsendnull���     bcke�>� �jvE�OjvE�OjvE�OjvE�OjvE�O*�,E�k/�,�,E�O*�,E�k/�,�,�,E�O*��/ J*��/ B*�- ; 8*j kih  *�/  *�,�6FO*a ,�6FO�a &j �6FOPU[OY��UUUUOjvE` OjvE` O ��[a �l kh a  �*a k/ � \ka kh *a �/ D*a k/ :*a -a [a ,\Z�81jv !*a -a [a ,\Z�81E` O�E` Y hUU[OY��O_ �k/a ,a  ,E` !O*a _ / "*a k/ *a _ !/a k/a ,_ 6FUUUU[OY�LO �k�j kh  � �*a "k/�,E` #O*a a $l %E` &O_ & ee*a ',FO*a a (a )*a (-6a *a +��/la , %Oa -��/%*a .,FOa /_ �/%a 0%_ #%a 1%��/%a 2%*�,FO*j 3UU[OY�n0 �@� @   ABCDEFGHIJKLMNOPQR���
�	��������� ��A �SS  B �TT � ( 
N�Qf 1 6 3��{�S�gev�N��DN� 
�� 
 N 2 - C l o v e r - W e e k 1 7 - l e s s o n 4 . m p 4   ( 7 3 . 8 3 M ,   2 0 2 0^t 6g 2 7e�   1 0 : 0 6  R0g ) 
W(~�����   |  N�} 
 
 
 � 
C �UU � E n g l i s h � n a m e : L i n 
 i d : 2 0 1 9 1 0 7 0 1 5 8 0 0 5 0 
 c l a s s : N 2 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 I M G _ 6 7 2 9 . M O V   ( 1 1 5 . 5 5 M ,   2 0 2 0^t 0 7g 1 2e�   1 0 : 0 7  R0g ) 
��QeN�}�u�b 
D �VV� T h e   b e s t   t h i n g   a b o u t   l e a r n i n g   E n g l i s h   o n l i n e   i s   t h a t   I   c a n   r e p e a t   w h a t   t h e   t e a c h e r   s a y s ,   b e c a u s e   i t   c a n   i m p r o v e   m y   g r a m m a r . 
 I   l i k e   t h e   o n l i n e   c o u r s e ,   b e c a u s e   I   c a n   s e e   t h e   f a m i l i a r   f a c e s   a n d   h u m o r o u s   t e a c h e r s   o f   m y   c l a s s m a t e s ,   I   l i k e   t h e   a t m o s p h e r e   o f   t h e   c l a s s r o o m   v e r y   m u c h 
 T h e   b i g g e s t   c h a l l e n g e   I   e n c o u n t e r e d   i n   l a n g u a g e   l e a r n i n g   o n l i n e   w a s   m y s e l f ,   b e c a u s e   I   w a s   a f r a i d   t h a t   I   w o u l d   n o t   b e   a b l e   t o   g e t   u p   f o r   c l a s s ,   w h i c h   w o u l d   l e a d   t o   t h e   d e d u c t i o n   o f   p o i n t s 
 W h a t   I   w a s   m o s t   e x c i t e d   a b o u t   t h i s   s u m m e r   w a s   t h a t   I   c o u l d   h a v e   s e v e r a l   s u m m e r   h o l i d a y s   a n d   g o   t o   m a n y   p l a c e s 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 W e e k 1 7 - L e s s o n 4 - K e v i n . m p 4   ( 1 2 . 4 7 M ,  e��Pg ) 
��QeN�}�u�b 
E �WW � 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 6 6 c 4 c 7 c 5 d 0 3 b e a d c 3 f 7 a c a 0 0 b d 6 b b a 9 b . m p 4   ( 9 . 3 9 M ,  e��Pg ) 
��QeN�}�u�b 
F �XX  
G �YY � 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 W I N _ 2 0 2 0 0 6 1 2 _ 1 0 _ 3 7 _ 2 0 _ P r o . m p 4   ( 7 1 . 1 7 M ,   2 0 2 0^t 0 7g 1 2e�   1 0 : 3 9  R0g ) 
��QeN�}�u�b 
H �ZZ � 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 v i d e o _ 2 0 2 0 0 6 1 2 _ 1 0 2 9 4 0 . m p 4   ( 3 9 . 1 4 M ,  e��Pg ) 
��QeN�}�u�b 
I �[[ j c l a s s : A 2 
 E n g l i s h   n a m e : s k y 
 I D   n u m b e r : 2 0 1 9 1 0 7 0 1 5 6 0 0 7 6 
 
J �\\ � 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 W I N _ 2 0 2 0 0 6 1 2 _ 1 0 _ 5 0 _ 0 4 _ P r o . m p 4   ( 7 0 . 4 6 M ,   2 0 2 0^t 0 7g 1 2e�   1 0 : 5 2  R0g ) 
��QeN�}�u�b 
K �]] � 
 
 
 
 
 
 
��  N� Q Q��{�S�gev���Y'�DN�   
��   
 V I D _ 2 0 2 0 0 6 1 2 _ 1 0 5 0 5 2 . m p 4   ( 1 0 3 . 9 9 M ,   2 0 2 0^t 7g 1 2e�   1 0 : 5 6  R0g ) 
��QeN�}�u�b   
L �^^ � 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 W I N _ 2 0 2 0 0 6 1 2 _ 1 0 _ 5 5 _ 2 5 _ P r o . m p 4   ( 5 9 . 6 8 M ,   2 0 2 0^t 0 7g 1 2e�   1 1 : 0 1  R0g ) 
��QeN�}�u�b 
M �__ � 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 W I N _ 2 0 2 0 0 6 1 2 _ 1 1 _ 0 5 _ 4 8 _ P r o . m p 4   ( 1 1 5 . 0 9 M ,   2 0 2 0^t 0 7g 1 2e�   1 1 : 1 0  R0g ) 
��QeN�}�u�b 
N �`` � j a c k e y l o v e   Z h o u W e n D o n g  ]�NO`QqN�   O n e D r i v e  e�N�0�剁g�w���SUQ�N�bv���c�0 
�� 
 W I N _ 2 0 2 0 0 6 1 2 _ 1 1 _ 2 6 _ 1 8 _ P r o   1 . m p 4 
 
O �aa � F a k e r   E n g l i s h   w o r k 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 V I D _ 2 0 2 0 0 6 1 2 _ 1 2 1 3 1 8 . m p 4   ( 6 1 . 9 6 M ,   2 0 2 0^t 0 7g 1 2e�   1 2 : 1 4  R0g ) 
��QeN�}�u�b 
P �bb � E n g l i s h   n a m e : P u p u 
 C h i n e s e   n a m e : C h e n g   x u 
 I D   n u m b e r : 2 0 1 9 1 0 7 0 1 5 6 0 0 8 9 
Q �cc � 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 w e e k 1 7   l e s s o n 4 . m p 4   ( 5 5 . 5 2 M ,   2 0 2 0^t 0 7g 1 2e�   1 3 : 0 6  R0g ) 
��QeN�}�u�b 
R �dd � 
��N� Q Q��{�S�gev���Y'�DN� 
�� 
 I M G _ 1 2 9 3 . M P 4   ( 9 4 . 1 9 M ,   2 0 2 0^t 0 7g 1 2e�   1 3 : 3 4  R0g ) 
��QeN�}�u�b 
�  �  �
  �	  �  �  �  �  �  �  �  �  �   ��  1 ��e�� e   fghijklmnopqrstuvw����������������������������f �xx ,�Ĕk   < 1 2 4 1 8 9 2 8 1 5 @ q q . c o m >g �yy <Q~� 1 9 1  �H��`!   < a 1 0 0 4 1 0 0 8 6 @ 1 6 3 . c o m >h �zz .Oލ5g�   < 1 0 1 6 1 9 9 6 1 1 @ q q . c o m >i �{{ 8`<yb /�� @ D : o   < 1 2 6 2 8 5 6 1 3 6 @ q q . c o m >j �|| < 1 5 6 4 7 0 7 4 9 2   < 1 5 6 4 7 0 7 4 9 2 @ q q . c o m >k �}} ,mw~�   < 1 7 7 1 5 5 2 9 0 3 @ q q . c o m >l �~~ < 1 0 6 3 8 6 1 4 8 5   < 1 0 6 3 8 6 1 4 8 5 @ q q . c o m >m � . l i e   < 1 0 7 8 0 4 2 8 3 7 @ q q . c o m >n ��� 2e�W#p��=ܗ�   < 3 2 0 3 7 9 4 9 4 @ q q . c o m >o ��� ( L   < 5 1 6 9 2 8 4 4 0 @ q q . c o m >p ��� D N 2   X u   R u n s h e n g   < 1 8 2 9 8 7 4 2 2 5 @ q q . c o m >q ��� 2Nr�O`�ZF `   < 4 1 1 8 1 5 8 1 4 @ q q . c o m >r ��� .SM"�SP   < 2 8 2 2 8 5 4 3 8 2 @ q q . c o m >s ��� | j a c k e y l o v e   Z h o u W e n D o n g   < j a c k e y l o v e 2 0 1 9 1 0 7 0 1 5 8 0 0 4 4 @ o u t l o o k . c o m >t ��� : F a k e r  ��g�m�   < 1 4 3 4 9 5 1 2 5 5 @ q q . c o m >u ��� 8N S�e���v�\pu�B   < 2 5 1 4 5 2 2 3 8 9 @ q q . c o m >v ��� ,:l'   < 2 5 1 0 4 4 6 9 6 7 @ q q . c o m >w ��� 8 8 2 4 2 5 0 9 0 2   < 8 2 4 2 5 0 9 0 2 @ q q . c o m >��  ��  ��  ��  ��  ��  ��  ��  ��  ��  ��  ��  ��  ��  2 ����� �   ����������������������������������������������� ��� " 1 2 4 1 8 9 2 8 1 5 @ q q . c o m� ��� $ a 1 0 0 4 1 0 0 8 6 @ 1 6 3 . c o m� ��� " 1 0 1 6 1 9 9 6 1 1 @ q q . c o m� ��� " 1 2 6 2 8 5 6 1 3 6 @ q q . c o m� ��� " 1 5 6 4 7 0 7 4 9 2 @ q q . c o m� ��� " 1 7 7 1 5 5 2 9 0 3 @ q q . c o m� ��� " 1 0 6 3 8 6 1 4 8 5 @ q q . c o m� ��� " 1 0 7 8 0 4 2 8 3 7 @ q q . c o m� ���   3 2 0 3 7 9 4 9 4 @ q q . c o m� ���   5 1 6 9 2 8 4 4 0 @ q q . c o m� ��� " 1 8 2 9 8 7 4 2 2 5 @ q q . c o m� ���   4 1 1 8 1 5 8 1 4 @ q q . c o m� ��� " 2 8 2 2 8 5 4 3 8 2 @ q q . c o m� ��� J j a c k e y l o v e 2 0 1 9 1 0 7 0 1 5 8 0 0 4 4 @ o u t l o o k . c o m� ��� " 1 4 3 4 9 5 1 2 5 5 @ q q . c o m� ��� " 2 5 1 4 5 2 2 3 8 9 @ q q . c o m� ��� " 2 5 1 0 4 4 6 9 6 7 @ q q . c o m� ���   8 2 4 2 5 0 9 0 2 @ q q . c o m��  ��  ��  ��  ��  ��  ��  ��  ��  ��  ��  ��  ��  ��  3 ������  ��  4 ������  ��  5 ���   W e e k   1 7   L e s s o n   46 ���  W i n t e c ascr  ��ޭ