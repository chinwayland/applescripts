FasdUAS 1.101.10   ��   ��    k             l        	  x     �� 
 ��   
 1      ��
�� 
ascr  �� ��
�� 
minv  m         �    2 . 4��       Yosemite (10.10) or later    	 �   4   Y o s e m i t e   ( 1 0 . 1 0 )   o r   l a t e r      x    �� ����    2  	 ��
�� 
osax��        l          x    "��  ��    4  
�� 
�� 
scpt  m  	   �    C a l e n d a r L i b   E C  �� ��
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
tstr o  �a�a 0 thedate theDate�g  �f  �h  �l  �k   �  l     �`�_�^�`  �_  �^   	 l      �]
�]  

-- Ask for delayed start
say "Is the first day of the semester delayed?"
set delayedYesOrNo to choose from list {"Yes", "No"} with prompt "Is the first day of the semester delayed?"
set delayedYesOrNo to item 1 of delayedYesOrNo

if delayedYesOrNo is "Yes" then
	say "If the first day is delayed and not on a Monday, when is the delayed first day?"
	repeat
		set theCurrentDate to current date
		display dialog "If the first day is delayed and not on a Monday, when is the delayed first day?" default answer (date string of theCurrentDate & space & time string of theCurrentDate)
		set theText to text returned of result
		try
			if theText is not "" then
				set delayedFirstDay to date theText -- a date object
				exit repeat
			end if
		on error
			beep
		end try
	end repeat
end if
    �( 
 - -   A s k   f o r   d e l a y e d   s t a r t 
 s a y   " I s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r   d e l a y e d ? " 
 s e t   d e l a y e d Y e s O r N o   t o   c h o o s e   f r o m   l i s t   { " Y e s " ,   " N o " }   w i t h   p r o m p t   " I s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r   d e l a y e d ? " 
 s e t   d e l a y e d Y e s O r N o   t o   i t e m   1   o f   d e l a y e d Y e s O r N o 
 
 i f   d e l a y e d Y e s O r N o   i s   " Y e s "   t h e n 
 	 s a y   " I f   t h e   f i r s t   d a y   i s   d e l a y e d   a n d   n o t   o n   a   M o n d a y ,   w h e n   i s   t h e   d e l a y e d   f i r s t   d a y ? " 
 	 r e p e a t 
 	 	 s e t   t h e C u r r e n t D a t e   t o   c u r r e n t   d a t e 
 	 	 d i s p l a y   d i a l o g   " I f   t h e   f i r s t   d a y   i s   d e l a y e d   a n d   n o t   o n   a   M o n d a y ,   w h e n   i s   t h e   d e l a y e d   f i r s t   d a y ? "   d e f a u l t   a n s w e r   ( d a t e   s t r i n g   o f   t h e C u r r e n t D a t e   &   s p a c e   &   t i m e   s t r i n g   o f   t h e C u r r e n t D a t e ) 
 	 	 s e t   t h e T e x t   t o   t e x t   r e t u r n e d   o f   r e s u l t 
 	 	 t r y 
 	 	 	 i f   t h e T e x t   i s   n o t   " "   t h e n 
 	 	 	 	 s e t   d e l a y e d F i r s t D a y   t o   d a t e   t h e T e x t   - -   a   d a t e   o b j e c t 
 	 	 	 	 e x i t   r e p e a t 
 	 	 	 e n d   i f 
 	 	 o n   e r r o r 
 	 	 	 b e e p 
 	 	 e n d   t r y 
 	 e n d   r e p e a t 
 e n d   i f 
	  l     �\�[�Z�\  �[  �Z    l '�Y�X r  ' [  # o  �W�W 0 thedate theDate l "�V�U ]  " 1   �T
�T 
days m   !�S�S �V  �U   o      �R�R 0 thedate theDate�Y  �X    l     �Q�P�O�Q  �P  �O    l (3�N�M r  (3 l (/ �L�K  n  (/!"! 1  +/�J
�J 
year" o  (+�I�I 0 thedate theDate�L  �K   o      �H�H 0 a  �N  �M   #$# l 4C%�G�F% r  4C&'& c  4?()( l 4;*�E�D* n  4;+,+ m  7;�C
�C 
mnth, o  47�B�B 0 thedate theDate�E  �D  ) m  ;>�A
�A 
nmbr' o      �@�@ 0 b  �G  �F  $ -.- l DO/�?�>/ r  DO010 n  DK232 1  GK�=
�= 
day 3 o  DG�<�< 0 thedate theDate1 o      �;�; 0 c  �?  �>  . 454 l     �:67�:  6  set d to "T000000Z"   7 �88 & s e t   d   t o   " T 0 0 0 0 0 0 Z "5 9:9 l PW;�9�8; r  PW<=< m  PS>> �?? H F R E Q = W E E K L Y ; I N T E R V A L = 1 ; B Y D A Y = ; U N T I L == o      �7�7 0 e  �9  �8  : @A@ l     �6�5�4�6  �5  �4  A BCB l X_DEFD r  X_GHG o  X[�3�3 0 thedate theDateH o      �2�2 0 enddate endDateE R L this variable represents the last day of the semester for Microsoft Outlook   F �II �   t h i s   v a r i a b l e   r e p r e s e n t s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r   f o r   M i c r o s o f t   O u t l o o kC JKJ l     �1�0�/�1  �0  �/  K LML l `oN�.�-N r  `oOPO J  `kQQ RSR o  `c�,�, 0 a  S TUT o  cf�+�+ 0 b  U V�*V o  fi�)�) 0 c  �*  P o      �(�( 0 thelist theList�.  �-  M WXW l     �'�&�%�'  �&  �%  X YZY l p�[�$�#[ r  p�\]\ J  pz^^ _`_ 1  pu�"
�" 
txdl` a�!a m  uxbb �cc  0�!  ] J      dd efe o      � �  
0 tid TIDf g�g 1  ���
� 
txdl�  �$  �#  Z hih l ��j��j r  ��klk c  ��mnm o  ���� 0 thelist theListn m  ���
� 
ctxtl o      �� "0 thelistasstring theListAsString�  �  i opo l ��q��q r  ��rsr o  ���� 
0 tid TIDs 1  ���
� 
txdl�  �  p tut l     ����  �  �  u vwv l ��x��x r  ��yzy J  ��{{ |}| o  ���� 0 e  } ~�~ o  ���� "0 thelistasstring theListAsString�  z o      �� 0 thelist theList�  �  w � l     ��
�	�  �
  �	  � ��� l     ����  � Y S the variable endOfRecurrence represents the last day of classes for Apple Calendar   � ��� �   t h e   v a r i a b l e   e n d O f R e c u r r e n c e   r e p r e s e n t s   t h e   l a s t   d a y   o f   c l a s s e s   f o r   A p p l e   C a l e n d a r� ��� l ������ r  ����� J  ���� ��� 1  ���
� 
txdl� ��� m  ���� ���  �  � J      �� ��� o      �� 
0 tid TID� ��� 1  ���
� 
txdl�  �  �  � ��� l ���� ��� r  ����� c  ����� o  ������ 0 thelist theList� m  ����
�� 
ctxt� o      ���� "0 endofrecurrence endOfRecurrence�   ��  � ��� l �������� r  ����� o  ������ 
0 tid TID� 1  ����
�� 
txdl��  ��  � ��� l     ��������  ��  ��  � ��� l �������� r  ����� J  ������  � o      ���� 0 
classnames 
classNames��  ��  � ��� l �������� O  ����� k  ���� ��� I �������
�� .miscactvnull��� ��� null��  ��  � ��� r  ��� n  ��� 1  	��
�� 
pnam� 2 	��
�� 
docu� o      ���� "0 listofdocuments listOfDocuments� ��� O ��� I �����
�� .sysottosnull���     TEXT� m  �� ��� � C h o o s e   w h i c h   N u m b e r s   d o c u m e n t   h a s   t h e   t i m e t a b l e   a n d   c l a s s   n a m e s .��  �  f  � ��� r  *��� I &�����
�� .gtqpchltns    @   @ ns  � o  "���� "0 listofdocuments listOfDocuments��  � o      ����  0 chosendocument chosenDocument� ��� r  +7��� n  +3��� 4  .3���
�� 
cobj� m  12���� � o  +.����  0 chosendocument chosenDocument� o      ����  0 chosendocument chosenDocument� ��� l 88��������  ��  ��  � ��� l 88��������  ��  ��  � ���� O  8���� k  C��� ��� l  CC������  �/)
		tell me to say "Choose the sheet named \"Classes\"."
		set listOfSheets to name of sheets
		set chosenSheet to choose from list listOfSheets
		set chosenSheet to item 1 of chosenSheet
		tell sheet chosenSheet
			tell table 1
				# get year one class names
				set yearOneBegin to first cell whose value of it contains "Class"
				set yearOneColumnAddress to address of column of yearOneBegin
				set yearOneColumnCount to count of cells of column yearOneColumnAddress
				set yearOneBegin to (address of row of yearOneBegin) + 1
				
				repeat with i from yearOneBegin to yearOneColumnCount
					if value of cell i of column yearOneColumnAddress is missing value then
						exit repeat
					else
						#						set end of classNames to "Year 1 " & (value of cell i of column yearOneColumnAddress & " " & value of cell i of column (yearOneColumnAddress + 4))
						set end of classNames to value of cell i of column (yearOneColumnAddress + 4) & " Year 1 " & (value of cell i of column yearOneColumnAddress)
					end if
					#					value of cell i of column yearOneColumnAddress
				end repeat
				
				# get year two class names
				set yearOneBegin to second cell whose value of it contains "Class"
				set yearOneColumnAddress to address of column of yearOneBegin
				set yearOneColumnCount to count of cells of column yearOneColumnAddress
				set yearOneBegin to (address of row of yearOneBegin) + 1
				
				repeat with i from yearOneBegin to yearOneColumnCount
					if value of cell i of column yearOneColumnAddress is missing value then
						exit repeat
					else
						#						set end of classNames to "Year 2 " & (value of cell i of column yearOneColumnAddress & " " & value of cell i of column (yearOneColumnAddress + 4))
						set end of classNames to value of cell i of column (yearOneColumnAddress + 4) & " Year 2 " & (value of cell i of column yearOneColumnAddress)
					end if
					#					value of cell i of column yearOneColumnAddress
				end repeat
				
				set sortedList to my simple_sort(the classNames)
				set classNames to sortedList
				
			end tell
			
		end tell
		   � ���R 
 	 	 t e l l   m e   t o   s a y   " C h o o s e   t h e   s h e e t   n a m e d   \ " C l a s s e s \ " . " 
 	 	 s e t   l i s t O f S h e e t s   t o   n a m e   o f   s h e e t s 
 	 	 s e t   c h o s e n S h e e t   t o   c h o o s e   f r o m   l i s t   l i s t O f S h e e t s 
 	 	 s e t   c h o s e n S h e e t   t o   i t e m   1   o f   c h o s e n S h e e t 
 	 	 t e l l   s h e e t   c h o s e n S h e e t 
 	 	 	 t e l l   t a b l e   1 
 	 	 	 	 #   g e t   y e a r   o n e   c l a s s   n a m e s 
 	 	 	 	 s e t   y e a r O n e B e g i n   t o   f i r s t   c e l l   w h o s e   v a l u e   o f   i t   c o n t a i n s   " C l a s s " 
 	 	 	 	 s e t   y e a r O n e C o l u m n A d d r e s s   t o   a d d r e s s   o f   c o l u m n   o f   y e a r O n e B e g i n 
 	 	 	 	 s e t   y e a r O n e C o l u m n C o u n t   t o   c o u n t   o f   c e l l s   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s 
 	 	 	 	 s e t   y e a r O n e B e g i n   t o   ( a d d r e s s   o f   r o w   o f   y e a r O n e B e g i n )   +   1 
 	 	 	 	 
 	 	 	 	 r e p e a t   w i t h   i   f r o m   y e a r O n e B e g i n   t o   y e a r O n e C o l u m n C o u n t 
 	 	 	 	 	 i f   v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s   i s   m i s s i n g   v a l u e   t h e n 
 	 	 	 	 	 	 e x i t   r e p e a t 
 	 	 	 	 	 e l s e 
 	 	 	 	 	 	 # 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e s   t o   " Y e a r   1   "   &   ( v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s   &   "   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   ( y e a r O n e C o l u m n A d d r e s s   +   4 ) ) 
 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e s   t o   v a l u e   o f   c e l l   i   o f   c o l u m n   ( y e a r O n e C o l u m n A d d r e s s   +   4 )   &   "   Y e a r   1   "   &   ( v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s ) 
 	 	 	 	 	 e n d   i f 
 	 	 	 	 	 # 	 	 	 	 	 v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s 
 	 	 	 	 e n d   r e p e a t 
 	 	 	 	 
 	 	 	 	 #   g e t   y e a r   t w o   c l a s s   n a m e s 
 	 	 	 	 s e t   y e a r O n e B e g i n   t o   s e c o n d   c e l l   w h o s e   v a l u e   o f   i t   c o n t a i n s   " C l a s s " 
 	 	 	 	 s e t   y e a r O n e C o l u m n A d d r e s s   t o   a d d r e s s   o f   c o l u m n   o f   y e a r O n e B e g i n 
 	 	 	 	 s e t   y e a r O n e C o l u m n C o u n t   t o   c o u n t   o f   c e l l s   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s 
 	 	 	 	 s e t   y e a r O n e B e g i n   t o   ( a d d r e s s   o f   r o w   o f   y e a r O n e B e g i n )   +   1 
 	 	 	 	 
 	 	 	 	 r e p e a t   w i t h   i   f r o m   y e a r O n e B e g i n   t o   y e a r O n e C o l u m n C o u n t 
 	 	 	 	 	 i f   v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s   i s   m i s s i n g   v a l u e   t h e n 
 	 	 	 	 	 	 e x i t   r e p e a t 
 	 	 	 	 	 e l s e 
 	 	 	 	 	 	 # 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e s   t o   " Y e a r   2   "   &   ( v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s   &   "   "   &   v a l u e   o f   c e l l   i   o f   c o l u m n   ( y e a r O n e C o l u m n A d d r e s s   +   4 ) ) 
 	 	 	 	 	 	 s e t   e n d   o f   c l a s s N a m e s   t o   v a l u e   o f   c e l l   i   o f   c o l u m n   ( y e a r O n e C o l u m n A d d r e s s   +   4 )   &   "   Y e a r   2   "   &   ( v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s ) 
 	 	 	 	 	 e n d   i f 
 	 	 	 	 	 # 	 	 	 	 	 v a l u e   o f   c e l l   i   o f   c o l u m n   y e a r O n e C o l u m n A d d r e s s 
 	 	 	 	 e n d   r e p e a t 
 	 	 	 	 
 	 	 	 	 s e t   s o r t e d L i s t   t o   m y   s i m p l e _ s o r t ( t h e   c l a s s N a m e s ) 
 	 	 	 	 s e t   c l a s s N a m e s   t o   s o r t e d L i s t 
 	 	 	 	 
 	 	 	 e n d   t e l l 
 	 	 	 
 	 	 e n d   t e l l 
 	 	� ��� l CC��������  ��  ��  � ��� l CC��������  ��  ��  � ��� O CO��� I GN�����
�� .sysottosnull���     TEXT� m  GJ�� ��� H C h o o s e   t h e   s h e e t   w i t h   t h e   t i m e t a b l e .��  �  f  CD� ��� r  P]��� n  PY��� 1  UY��
�� 
pnam� 2 PU��
�� 
NmSh� o      ���� 0 listofsheets listOfSheets� ��� r  ^i��� I ^e�����
�� .gtqpchltns    @   @ ns  � o  ^a���� 0 listofsheets listOfSheets��  � o      ���� 0 chosensheet chosenSheet� ��� r  jv��� n  jr��� 4  mr���
�� 
cobj� m  pq���� � o  jm���� 0 chosensheet chosenSheet� o      ���� 0 chosensheet chosenSheet� ��� l ww��������  ��  ��  � ��� l ww��������  ��  ��  � ��� O w���� I {������
�� .sysottosnull���     TEXT� m  {~�� ��� ( G e t t i n g   c l a s s   n a m e s .��  �  f  wx� ��� l ����������  ��  ��  �    O  �g O  �f Y  �e���� k  �`		 

 U  �� k  ��  r  �� n  �� 1  ����
�� 
NMCv n  �� 4  ����
�� 
NmCl m  ������  4  ����
�� 
NMCo o  ������ 0 columnindex columnIndex o      ���� 0 	yearvalue 	yearValue �� Z  �� E  �� o  ������ 0 	yearvalue 	yearValue m  ��   �!! Y'N  r  ��"#" m  ��$$ �%%  Y e a r   1# o      ���� 0 	yearvalue 	yearValue &'& E  ��()( o  ������ 0 	yearvalue 	yearValue) m  ��** �++ Y'N�' ,��, r  ��-.- m  ��// �00  Y e a r   2. o      ���� 0 	yearvalue 	yearValue��    S  ����   m  ������  121 l ����������  ��  ��  2 343 Y  �^5��67��5 k  �Y88 9:9 Z  �W;<����; > �=>= n  �?@? 1  ��
�� 
NMCv@ n  �ABA 4  ��C
�� 
NmClC o  ���� 0 rowindex rowIndexB 4  ���D
�� 
NMCoD o   ���� 0 columnindex columnIndex> m  ��
�� 
msng< k  SEE FGF Z  QHI����H H  .JJ E -KLK o  ���� 0 
classnames 
classNamesL b  ,MNM b  OPO o  ���� 0 	yearvalue 	yearValueP m  QQ �RR   N n  +STS 1  '+��
�� 
NMCvT n  'UVU 4  "'��W
�� 
NmClW o  %&���� 0 rowindex rowIndexV 4  "��X
�� 
NMCoX o   !���� 0 columnindex columnIndexI r  1MYZY b  1H[\[ b  18]^] o  14���� 0 	yearvalue 	yearValue^ m  47__ �``   \ n  8Gaba 1  CG��
�� 
NMCvb n  8Ccdc 4  >C��e
�� 
NmCle o  AB���� 0 rowindex rowIndexd 4  8>��f
�� 
NMCof o  <=���� 0 columnindex columnIndexZ n      ghg  ;  KLh o  HK���� 0 
classnames 
classNames��  ��  G i��i l RR��������  ��  ��  ��  ��  ��  : j��j l XX��������  ��  ��  ��  �� 0 rowindex rowIndex6 m  ������ 7 m  ������ %��  4 k��k l __�������  ��  �  ��  �� 0 columnindex columnIndex m  ���~�~  m  ���}�} ��   4  ���|l
�| 
NmTbl m  ���{�{  4  ���zm
�z 
NmShm o  ���y�y 0 chosensheet chosenSheet non O htpqp I ls�xr�w
�x .sysottosnull���     TEXTr m  loss �tt V G r a b b i n g   d a t a   f r o m   t h e   n u m b e r s   s p r e a d s h e e t .�w  q  f  hio u�vu O  u�vwv O  ��xyx k  ��zz {|{ l ���u}~�u  } E ? Find table of room numbers and teachers and put it into a list   ~ � ~   F i n d   t a b l e   o f   r o o m   n u m b e r s   a n d   t e a c h e r s   a n d   p u t   i t   i n t o   a   l i s t| ��� r  ����� 6 ����� 4 ���t�
�t 
NmCl� m  ���s�s � = ����� 1  ���r
�r 
NMCv� m  ���� ���  R o o m� o      �q�q 0 	roomindex 	roomIndex� ��� r  ����� n  ����� 1  ���p
�p 
NMad� n  ����� m  ���o
�o 
NMCo� o  ���n�n 0 	roomindex 	roomIndex� o      �m�m 0 columnnumber columnNumber� ��� r  ����� n  ����� 1  ���l
�l 
NMad� n  ����� m  ���k
�k 
NMRw� o  ���j�j 0 	roomindex 	roomIndex� o      �i�i 0 	rownumber 	rowNumber� ��� r  ����� I ���h��g
�h .corecnte****       ****� n ����� 2 ���f
�f 
NmCl� 4  ���e�
�e 
NMCo� o  ���d�d 0 columnnumber columnNumber�g  � o      �c�c 0 rowcount rowCount� ��� r  ����� J  ���b�b  � o      �a�a 0 roomnumbers roomNumbers� ��� Y  �C��`���_� k  �>�� ��� Z  �<���^�]� > ���� n  ���� 1  ��\
�\ 
NMCv� n  ����� 4  ���[�
�[ 
NmCl� o  ���Z�Z 0 i  � 4  ���Y�
�Y 
NMCo� l ����X�W� \  ����� o  ���V�V 0 columnnumber columnNumber� m  ���U�U �X  �W  � m  �T
�T 
msng� r  	8��� J  	3�� ��� l 	��S�R� n  	��� 1  �Q
�Q 
NMCv� n  	��� 4  �P�
�P 
NmCl� o  �O�O 0 i  � 4  	�N�
�N 
NMCo� l ��M�L� \  ��� o  �K�K 0 columnnumber columnNumber� m  �J�J �M  �L  �S  �R  � ��I� c  1��� n  -��� 1  )-�H
�H 
NMCv� n  )��� 4  $)�G�
�G 
NmCl� o  '(�F�F 0 i  � 4  $�E�
�E 
NMCo� o   #�D�D 0 columnnumber columnNumber� m  -0�C
�C 
ctxt�I  � n      ���  ;  67� o  36�B�B 0 roomnumbers roomNumbers�^  �]  � ��A� l ==�@�?�>�@  �?  �>  �A  �` 0 i  � [  ����� o  ���=�= 0 	rownumber 	rowNumber� m  ���<�< � o  ���;�; 0 rowcount rowCount�_  � ��� l  DD�:���:  ���
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
				   � ���	j 
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
 	 	 	 	� ��� l DD�9���9  � !  Grab data from spreadsheet   � ��� 6   G r a b   d a t a   f r o m   s p r e a d s h e e t� ��� r  DJ��� J  DF�8�8  � o      �7�7 0 lessons  � ��6� Y  K���5���4� Y  W���3���2� Z  e����1�0� > ex��� n  et��� 1  pt�/
�/ 
NMCv� n  ep��� 4  kp�.�
�. 
NmCl� o  no�-�- 0 j  � 4  ek�,�
�, 
NMCo� o  ij�+�+ 0 i  � m  tw�*
�* 
msng� r  {���� J  {��� ��� n  {���� 1  ���)
�) 
NMCv� n  {�� � 4  ���(
�( 
NmCl m  ���'�'   4  {��&
�& 
NMCo o  ��%�% 0 i  �  n  �� 1  ���$
�$ 
NMCv n  �� 4  ���#	
�# 
NmCl	 m  ���"�"  4  ���!

�! 
NMCo
 o  ��� �  0 i    n  �� 1  ���
� 
NMCv n  �� 4  ���
� 
NmCl o  ���� 0 j   4  ���
� 
NMCo m  ����  � n  �� 1  ���
� 
NMCv n  �� 4  ���
� 
NmCl o  ���� 0 j   4  ���
� 
NMCo o  ���� 0 i  �  � n        ;  �� o  ���� 0 lessons  �1  �0  �3 0 j  � m  Z]�� � m  ]`�� "�2  �5 0 i  � m  NO�� � m  OR�� �4  �6  y 4  ���
� 
NmTb m  ���� w 4  u}�
� 
NmSh o  y|�� 0 chosensheet chosenSheet�v  � 4  8@�
� 
docu o  <?�
�
  0 chosendocument chosenDocument��  � m  ���                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  ��  ��  �  !  l     �	���	  �  �  ! "#" l     �$%�  $ 7 1 convert year 1 or year 2 from chinese to english   % �&& b   c o n v e r t   y e a r   1   o r   y e a r   2   f r o m   c h i n e s e   t o   e n g l i s h# '(' l ��)��) r  ��*+* J  ����  + o      �� 0 lessons3  �  �  ( ,-, l �E.�� . X  �E/��0/ Z  �@12��31 E  ��454 n  ��676 4  ����8
�� 
cobj8 m  ������ 7 o  ������ 
0 lesson  5 m  ��99 �:: Y'N 2 r  �;<; J  �== >?> n  �@A@ 4  ���B
�� 
cobjB m  ���� A o  ������ 
0 lesson  ? CDC m  EE �FF  Y e a r   1D GHG n  IJI 4  ��K
�� 
cobjK m  
���� J o  ���� 
0 lesson  H L��L n  MNM 4  ��O
�� 
cobjO m  ���� N o  ���� 
0 lesson  ��  < n      PQP  ;  Q o  ���� 0 lessons3  ��  3 r   @RSR J   ;TT UVU n   &WXW 4  !&��Y
�� 
cobjY m  $%���� X o   !���� 
0 lesson  V Z[Z m  &)\\ �]]  Y e a r   2[ ^_^ n  )/`a` 4  */��b
�� 
cobjb m  -.���� a o  )*���� 
0 lesson  _ c��c n  /7ded 4  07��f
�� 
cobjf m  36���� e o  /0���� 
0 lesson  ��  S n      ghg  ;  >?h o  ;>���� 0 lessons3  �� 
0 lesson  0 o  ������ 0 lessons  �  �   - iji l     ��������  ��  ��  j klk l     ��mn��  m   add time to each event   n �oo .   a d d   t i m e   t o   e a c h   e v e n tl pqp l FLr����r r  FLsts J  FH����  t o      ���� 0 lessons4  ��  ��  q uvu l Mw����w X  Mx��yx k  c	zz {|{ r  cm}~} n  ci� 4  di���
�� 
cobj� m  gh���� � o  cd���� 
0 lesson  ~ o      ���� 0 	inputtime 	inputTime| ��� r  n���� b  n���� b  n|��� n  nx��� 4  qx���
�� 
cwor� m  tw���� � o  nq���� 0 	inputtime 	inputTime� m  x{�� ���  :� n  |���� 4  ����
�� 
cwor� m  ������ � o  |���� 0 	inputtime 	inputTime� o      ���� 0 thetime theTime� ��� Z  �������� ?  ����� c  ����� n  ����� 4  �����
�� 
cwor� m  ������ � l �������� o  ������ 0 thetime theTime��  ��  � m  ����
�� 
nmbr� m  ������ � r  ����� c  ����� b  ����� b  ����� b  ����� \  ����� l �������� n  ����� 4  �����
�� 
cwor� m  ������ � o  ������ 0 thetime theTime��  ��  � m  ������ � m  ���� ���  :� n  ����� 4  �����
�� 
cwor� m  ������ � o  ������ 0 thetime theTime� m  ���� ���  P M� m  ����
�� 
ctxt� o      ���� 0 thetime2 theTime2��  � r  ����� c  ����� b  ����� b  ����� b  ����� n  ����� 4  �����
�� 
cwor� m  ������ � o  ������ 0 thetime theTime� m  ���� ���  :� n  ����� 4  �����
�� 
cwor� m  ������ � o  ������ 0 thetime theTime� m  ���� ���  A M� m  ����
�� 
ctxt� o      ���� 0 thetime2 theTime2� ���� r  �	��� J  ��� ��� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ��� n  ����� 4  �����
�� 
cobj� m  ������ � o  ������ 
0 lesson  � ��� o  ������ 0 thetime2 theTime2� ���� n  � ��� 4  � ���
�� 
cobj� m  ������ � o  ������ 
0 lesson  ��  � n      ���  ;  � o  ���� 0 lessons4  ��  �� 
0 lesson  y o  PS���� 0 lessons3  ��  ��  v ��� l     ��������  ��  ��  � ��� l     ������  � , & Combine Year, class, and teacher name   � ��� L   C o m b i n e   Y e a r ,   c l a s s ,   a n d   t e a c h e r   n a m e� ��� l ������ r  ��� J  ����  � o      ���� 0 lessons5  ��  ��  � ��� l W������ X  W����� r  ,R��� J  ,M�� ��� b  ,?��� b  ,6��� n  ,2��� 4  -2���
�� 
cobj� m  01���� � o  ,-���� 
0 lesson  � m  25�� ���   � n  6>� � 4  7>��
�� 
cobj m  :=����   o  67���� 
0 lesson  �  n  ?E 4  @E��
�� 
cobj m  CD����  o  ?@���� 
0 lesson   �� n  EK	 4  FK��

�� 
cobj
 m  IJ���� 	 o  EF���� 
0 lesson  ��  � n        ;  PQ o  MP���� 0 lessons5  �� 
0 lesson  � o  ���� 0 lessons4  ��  ��  �  l     ��������  ��  ��    l     ����   ( " add date to each item in the list    � D   a d d   d a t e   t o   e a c h   i t e m   i n   t h e   l i s t  l X^���� r  X^ J  XZ����   o      ���� 0 lessons6  ��  ��    l _v���� X  _v�� Z  uq � E  u!"! n  u{#$# 4  v{�~%
�~ 
cobj% m  yz�}�} $ o  uv�|�| 
0 lesson  " m  {~&& �''  M o n d a y r  ��()( J  ��** +,+ n  ��-.- 4  ���{/
�{ 
cobj/ m  ���z�z . o  ���y�y 
0 lesson  , 0�x0 b  ��121 b  ��343 n  ��565 1  ���w
�w 
dstr6 o  ���v�v 0 firstday firstDay4 m  ��77 �88    a t  2 n  ��9:9 4  ���u;
�u 
cobj; m  ���t�t : o  ���s�s 
0 lesson  �x  ) n      <=<  ;  ��= o  ���r�r 0 lessons6    >?> E  ��@A@ n  ��BCB 4  ���qD
�q 
cobjD m  ���p�p C o  ���o�o 
0 lesson  A m  ��EE �FF  T u e s d a y? GHG r  ��IJI J  ��KK LML n  ��NON 4  ���nP
�n 
cobjP m  ���m�m O o  ���l�l 
0 lesson  M Q�kQ b  ��RSR b  ��TUT n  ��VWV 1  ���j
�j 
dstrW l ��X�i�hX [  ��YZY o  ���g�g 0 firstday firstDayZ l ��[�f�e[ ]  ��\]\ 1  ���d
�d 
days] m  ���c�c �f  �e  �i  �h  U m  ��^^ �__    a t  S n  ��`a` 4  ���bb
�b 
cobjb m  ���a�a a o  ���`�` 
0 lesson  �k  J n      cdc  ;  ��d o  ���_�_ 0 lessons6  H efe E  ��ghg n  ��iji 4  ���^k
�^ 
cobjk m  ���]�] j o  ���\�\ 
0 lesson  h m  ��ll �mm  W e d n e s d a yf non r  �pqp J  � rr sts n  ��uvu 4  ���[w
�[ 
cobjw m  ���Z�Z v o  ���Y�Y 
0 lesson  t x�Xx b  ��yzy b  ��{|{ n  ��}~} 1  ���W
�W 
dstr~ l ���V�U [  ����� o  ���T�T 0 firstday firstDay� l ����S�R� ]  ����� 1  ���Q
�Q 
days� m  ���P�P �S  �R  �V  �U  | m  ���� ���    a t  z n  ����� 4  ���O�
�O 
cobj� m  ���N�N � o  ���M�M 
0 lesson  �X  q n      ���  ;  � o   �L�L 0 lessons6  o ��� E  ��� n  ��� 4  	�K�
�K 
cobj� m  �J�J � o  	�I�I 
0 lesson  � m  �� ���  T h u r s d a y� ��� r  8��� J  3�� ��� n  ��� 4  �H�
�H 
cobj� m  �G�G � o  �F�F 
0 lesson  � ��E� b  1��� b  *��� n  &��� 1  $&�D
�D 
dstr� l $��C�B� [  $��� o  �A�A 0 firstday firstDay� l #��@�?� ]  #��� 1  !�>
�> 
days� m  !"�=�= �@  �?  �C  �B  � m  &)�� ���    a t  � n  *0��� 4  +0�<�
�< 
cobj� m  ./�;�; � o  *+�:�: 
0 lesson  �E  � n      ���  ;  67� o  36�9�9 0 lessons6  � ��� E  ;E��� n  ;A��� 4  <A�8�
�8 
cobj� m  ?@�7�7 � o  ;<�6�6 
0 lesson  � m  AD�� ���  F r i d a y� ��5� r  Hm��� J  Hh�� ��� n  HN��� 4  IN�4�
�4 
cobj� m  LM�3�3 � o  HI�2�2 
0 lesson  � ��1� b  Nf��� b  N_��� n  N[��� 1  Y[�0
�0 
dstr� l NY��/�.� [  NY��� o  NQ�-�- 0 firstday firstDay� l QX��,�+� ]  QX��� 1  QT�*
�* 
days� m  TW�)�) �,  �+  �/  �.  � m  [^�� ���    a t  � n  _e��� 4  `e�(�
�( 
cobj� m  cd�'�' � o  _`�&�& 
0 lesson  �1  � n      ���  ;  kl� o  hk�%�% 0 lessons6  �5  �  �� 
0 lesson   o  be�$�$ 0 lessons5  ��  ��   ��� l     �#�"�!�#  �"  �!  � ��� l     � ���   � ) # convert each item to a date object   � ��� F   c o n v e r t   e a c h   i t e m   t o   a   d a t e   o b j e c t� ��� l w}���� r  w}��� J  wy��  � o      �� 0 lessons7  �  �  � ��� l ~����� X  ~����� k  ���� ��� r  ����� n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  � o      �� 0 
datestring 
dateString� ��� r  ����� 4  ����
� 
ldt � o  ���� 0 
datestring 
dateString� o      �� 0 thedate theDate� ��� r  ����� J  ���� ��� n  ����� 4  ����
� 
cobj� m  ���� � o  ���� 
0 lesson  �  �  o  ���� 0 thedate theDate�  � n        ;  �� o  ���� 0 lessons7  �  � 
0 lesson  � o  ���
�
 0 lessons6  �  �  �  l     �	���	  �  �    l     ��   #  Add room numbers to the list    �		 :   A d d   r o o m   n u m b e r s   t o   t h e   l i s t 

 l ���� r  �� J  ����   o      �� 0 lessons8  �  �    l �R��  X  �R�� Z  �M�� E  �� n  �� 4  ����
�� 
cobj m  ������  o  ������ 
0 lesson   m  �� �  / r  � J  ��   !"! n  ��#$# 4  ����%
�� 
cobj% m  ������ $ o  ������ 
0 lesson  " &'& n  ��()( 4  ����*
�� 
cobj* m  ������ ) o  ������ 
0 lesson  ' +��+ m  ��,, �--  M u l t i p l e��   n      ./.  ;   / o  � ���� 0 lessons8  ��   X  M0��10 Z  H23����2 E  (454 n  !676 4  !��8
�� 
cobj8 m   ���� 7 o  ���� 
0 lesson  5 n  !'9:9 4  "'��;
�� 
cobj; m  %&���� : o  !"���� 0 room  3 r  +D<=< J  +?>> ?@? n  +1ABA 4  ,1��C
�� 
cobjC m  /0���� B o  +,���� 
0 lesson  @ DED n  17FGF 4  27��H
�� 
cobjH m  56���� G o  12���� 
0 lesson  E I��I n  7=JKJ 4  8=��L
�� 
cobjL m  ;<���� K o  78���� 0 room  ��  = n      MNM  ;  BCN o  ?B���� 0 lessons8  ��  ��  �� 0 room  1 o  ���� 0 roomnumbers roomNumbers�� 
0 lesson   o  ������ 0 lessons7  �  �    OPO l     ��������  ��  ��  P QRQ l      ��ST��  S � �
set classNames to {}
repeat with lesson in lessons8
	set end of classNames to item 1 of lesson
end repeat

set classNames2 to addr(classNames)
   T �UU  
 s e t   c l a s s N a m e s   t o   { } 
 r e p e a t   w i t h   l e s s o n   i n   l e s s o n s 8 
 	 s e t   e n d   o f   c l a s s N a m e s   t o   i t e m   1   o f   l e s s o n 
 e n d   r e p e a t 
 
 s e t   c l a s s N a m e s 2   t o   a d d r ( c l a s s N a m e s ) 
R VWV l     ��������  ��  ��  W XYX l      ��Z[��  ZC=
set lessons9 to {}
#set classNamesTeacherFirst to {}
repeat with lesson in lessons8
	if item 1 of lesson contains "/" then
		set teacherFirst to word 4 of item 1 of lesson
		set teacherSecond to word 5 of item 1 of lesson
		set yearWord to word 1 of item 1 of lesson
		set yearNumber to word 2 of item 1 of lesson
		set className to word 3 of item 1 of lesson
		set eventName to teacherFirst & "/" & teacherSecond & " " & yearWord & " " & yearNumber & " " & className
	else
		set teacher to word 4 of item 1 of lesson
		set yearWord to word 1 of item 1 of lesson
		set yearNumber to word 2 of item 1 of lesson
		set className to word 3 of item 1 of lesson
		set eventName to teacher & " " & yearWord & " " & yearNumber & " " & className
	end if
	set end of lessons9 to {eventName, item 2 of lesson, item 3 of lesson}
end repeat
   [ �\\z 
 s e t   l e s s o n s 9   t o   { } 
 # s e t   c l a s s N a m e s T e a c h e r F i r s t   t o   { } 
 r e p e a t   w i t h   l e s s o n   i n   l e s s o n s 8 
 	 i f   i t e m   1   o f   l e s s o n   c o n t a i n s   " / "   t h e n 
 	 	 s e t   t e a c h e r F i r s t   t o   w o r d   4   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   t e a c h e r S e c o n d   t o   w o r d   5   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   y e a r W o r d   t o   w o r d   1   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   y e a r N u m b e r   t o   w o r d   2   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   c l a s s N a m e   t o   w o r d   3   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   e v e n t N a m e   t o   t e a c h e r F i r s t   &   " / "   &   t e a c h e r S e c o n d   &   "   "   &   y e a r W o r d   &   "   "   &   y e a r N u m b e r   &   "   "   &   c l a s s N a m e 
 	 e l s e 
 	 	 s e t   t e a c h e r   t o   w o r d   4   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   y e a r W o r d   t o   w o r d   1   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   y e a r N u m b e r   t o   w o r d   2   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   c l a s s N a m e   t o   w o r d   3   o f   i t e m   1   o f   l e s s o n 
 	 	 s e t   e v e n t N a m e   t o   t e a c h e r   &   "   "   &   y e a r W o r d   &   "   "   &   y e a r N u m b e r   &   "   "   &   c l a s s N a m e 
 	 e n d   i f 
 	 s e t   e n d   o f   l e s s o n s 9   t o   { e v e n t N a m e ,   i t e m   2   o f   l e s s o n ,   i t e m   3   o f   l e s s o n } 
 e n d   r e p e a t 
Y ]^] l     ��������  ��  ��  ^ _`_ l     ��������  ��  ��  ` aba l     ��cd��  c . ( Creates the calendars in the chosen app   d �ee P   C r e a t e s   t h e   c a l e n d a r s   i n   t h e   c h o s e n   a p pb fgf l S�h����h Z  S�ij��ki = SXlml o  ST���� &0 chosencalendarapp chosenCalendarAppm m  TWnn �oo  A p p l e   C a l e n d a rj k  [�pp qrq l [[��st��  s &   Deletes previous work calendars   t �uu @   D e l e t e s   p r e v i o u s   w o r k   c a l e n d a r sr vwv O  [�xyx k  a�zz {|{ I af������
�� .miscactvnull��� ��� null��  ��  | }~} O gs� I kr�����
�� .sysottosnull���     TEXT� m  kn�� ��� n D e l e t i n g   p r e v i o u s   C a l e n d a r s   w i t h   t h e   w o r d   " Y e a r "   i n   i t .��  �  f  gh~ ��� l tt��������  ��  ��  � ���� O  t���� I ��������
�� .coredelonull���     obj ��  ��  � l t������� 6 t���� 2 ty��
�� 
wres� E  |���� 1  }���
�� 
pnam� m  ���� ���  Y e a r��  ��  ��  y m  [^���                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  w ��� l ����������  ��  ��  � ��� l ����������  ��  ��  � ��� l ��������  � 3 - add calendars and events to the Calendar app   � ��� Z   a d d   c a l e n d a r s   a n d   e v e n t s   t o   t h e   C a l e n d a r   a p p� ��� I �������
�� .sysottosnull���     TEXT� m  ���� ��� $ C r e a t i n g   C a l e n d a r s��  � ��� O  �
��� k  �
�� ��� I ��������
�� .miscactvnull��� ��� null��  ��  � ��� Y  ���������� Z  ��������� H  ���� l �������� I �������
�� .coredoexnull���     ****� l �������� 4  �����
�� 
wres� l �������� n  ����� 4  �����
�� 
cobj� o  ������ 0 i  � o  ������ 0 
classnames 
classNames��  ��  ��  ��  ��  ��  ��  � I �������
�� .wrbtaec2null��� ��� null��  � �����
�� 
wtnm� n  ����� 4  �����
�� 
cobj� o  ������ 0 i  � o  ������ 0 
classnames 
classNames��  ��  ��  �� 0 i  � m  ������ � I �������
�� .corecnte****       ****� o  ������ 0 
classnames 
classNames��  ��  � ��� O ����� I �������
�� .sysottosnull���     TEXT� m  ���� ���  C r e a t i n g   E v e n t s��  �  f  ��� ���� X  �
����� Z  	

������ I 	
	�����
�� .coredoexnull���     ****� 4  	
	���
�� 
wres� l 		������ n  		��� 4  		���
�� 
cobj� m  		���� � o  		���� 
0 lesson  ��  ��  ��  � O  		s��� I 	,	r�����
�� .corecrel****      � null��  � ����
�� 
kocl� m  	0	3��
�� 
wrev� ����
�� 
prdt� K  	6	l�� �~��
�~ 
wr11� n  	9	?��� 4  	:	?�}�
�} 
cobj� m  	=	>�|�| � o  	9	:�{�{ 
0 lesson  � �z��
�z 
wr14� n  	B	H��� 4  	C	H�y�
�y 
cobj� m  	F	G�x�x � o  	B	C�w�w 
0 lesson  � �v��
�v 
wr1s� n  	K	Q��� 4  	L	Q�u�
�u 
cobj� m  	O	P�t�t � o  	K	L�s�s 
0 lesson  � �r��
�r 
wr5s� [  	T	b��� l 	T	Z��q�p� n  	T	Z��� 4  	U	Z�o�
�o 
cobj� m  	X	Y�n�n � o  	T	U�m�m 
0 lesson  �q  �p  � l 	Z	a��l�k� ]  	Z	a��� 1  	Z	]�j
�j 
min � m  	]	`�i�i P�l  �k  � �h��g
�h 
wr15� o  	e	h�f�f "0 endofrecurrence endOfRecurrence�g  �  � 4  		)�e�
�e 
wres� l 	!	(��d�c� n  	!	(��� 4  	"	'�b�
�b 
cobj� m  	%	&�a�a � o  	!	"�`�` 
0 lesson  �d  �c  ��  � k  	v
�� ��� r  	v	���� n  	v	���� 2 	|	��_
�_ 
cwor� n  	v	|��� 4  	w	|�^�
�^ 
cobj� m  	z	{�]�] � o  	v	w�\�\ 
0 lesson  � o      �[�[ 0 problem  �    r  	�	� b  	�	� b  	�	� b  	�	�	 b  	�	�

 b  	�	� b  	�	� n  	�	� 4  	�	��Z
�Z 
cobj m  	�	��Y�Y  o  	�	��X�X 0 problem   m  	�	� �    n  	�	� 4  	�	��W
�W 
cobj m  	�	��V�V  o  	�	��U�U 0 problem   m  	�	� �   	 n  	�	� 4  	�	��T
�T 
cobj m  	�	��S�S  o  	�	��R�R 0 problem   m  	�	� �    n  	�	�  4  	�	��Q!
�Q 
cobj! m  	�	��P�P   o  	�	��O�O 0 problem   o      �N�N "0 neweventsummary newEventSummary "�M" O  	�
#$# I 	�
�L�K%
�L .corecrel****      � null�K  % �J&'
�J 
kocl& m  	�	��I
�I 
wrev' �H(�G
�H 
prdt( K  	�
)) �F*+
�F 
wr11* n  	�	�,-, 4  	�	��E.
�E 
cobj. m  	�	��D�D - o  	�	��C�C 
0 lesson  + �B/0
�B 
wr14/ n  	�	�121 4  	�	��A3
�A 
cobj3 m  	�	��@�@ 2 o  	�	��?�? 
0 lesson  0 �>45
�> 
wr1s4 n  	�	�676 4  	�	��=8
�= 
cobj8 m  	�	��<�< 7 o  	�	��;�; 
0 lesson  5 �:9:
�: 
wr5s9 [  	�	�;<; l 	�	�=�9�8= n  	�	�>?> 4  	�	��7@
�7 
cobj@ m  	�	��6�6 ? o  	�	��5�5 
0 lesson  �9  �8  < l 	�	�A�4�3A ]  	�	�BCB 1  	�	��2
�2 
min C m  	�	��1�1 P�4  �3  : �0D�/
�0 
wr15D o  	�
�.�. "0 endofrecurrence endOfRecurrence�/  �G  $ 4  	�	��-E
�- 
wresE o  	�	��,�, "0 neweventsummary newEventSummary�M  �� 
0 lesson  � o  ���+�+ 0 lessons8  ��  � m  ��FF�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  � GHG l 

�*�)�(�*  �)  �(  H IJI l 

�'�&�%�'  �&  �%  J KLK l 

�$MN�$  M   Sets the time zone   N �OO &   S e t s   t h e   t i m e   z o n eL PQP O 

 RSR I 

�#T�"
�# .sysottosnull���     TEXTT m  

UU �VV ^ C h a n g i n g   t h e   t i m e   z o n e   t o   t h e   C h i n e s e   t i m e   z o n e�"  S  f  

Q WXW l 
!
!�!� ��!  �   �  X YZY O  
!
D[\[ r  
'
C]^] 6 
'
?_`_ n  
'
0aba 1  
,
0�
� 
pnamb 2 
'
,�
� 
wres` E  
3
>cdc 1  
4
8�
� 
pnamd m  
9
=ee �ff  Y e a r^ o      �� 0 thecalendars theCalendars\ m  
!
$gg�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  Z hih l 
E
E����  �  �  i jkj r  
E
Zlml I 
E
Vn��n z� 
� .!Cls!fstnull��� ��� null�  �  m o      �� 0 thestore theStorek opo l 
[
[����  �  �  p qrq r  
[
bsts o  
[
^�� 0 firstday firstDayt o      �� 0 	startdate 	startDater uvu r  
c
rwxw [  
c
nyzy o  
c
f�� 0 	startdate 	startDatez ]  
f
m{|{ 1  
f
i�
� 
days| m  
i
l�� x o      �� 0 enddate endDatev }~} l 
s
s�
�	��
  �	  �  ~ � r  
s
���� I 
s
����� z� 
� .!CLs!fcAnull���     ****� o  
~
��� 0 thecalendars theCalendars� ���
� 
!CtY� m  
�
��� z� 
� !Tct!TtL� ���
� 
!Cst� o  
�
��� 0 thestore theStore�  � o      � �  0 thecalendars theCalendars� ��� l 
�
���������  ��  ��  � ��� r  
�
���� I 
�
������ z�� 
�� .!CLs!feSnull��� ��� null��  � ����
�� 
!Csd� o  
�
����� 0 	startdate 	startDate� ����
�� 
!Ced� o  
�
����� 0 enddate endDate� ����
�� 
!Csc� o  
�
����� 0 thecalendars theCalendars� �����
�� 
!Cst� o  
�
����� 0 thestore theStore��  � o      ���� 0 	theevents 	theEvents� ��� l 
�
���������  ��  ��  � ��� Y  
�&�������� k  
�!�� ��� r  
���� I 
�
������ z�� 
�� .!CLs!mo2null��� ��� null��  � ����
�� 
!Cev� n  
�
���� 4  
�
����
�� 
cobj� o  
�
����� 0 i  � o  
�
����� 0 	theevents 	theEvents� �����
�� 
!Ctz� m  
�
��� ���  A s i a / S h a n g h a i��  � o      ���� 0 thenewevent theNewEvent� ���� I !����� z�� 
�� .!CLs!updnull��� ��� null��  � ����
�� 
!Cev� o  ���� 0 thenewevent theNewEvent� �����
�� 
!Cst� o  ���� 0 thestore theStore��  ��  �� 0 i  � m  
�
����� � I 
�
������
�� .corecnte****       ****� o  
�
����� 0 	theevents 	theEvents��  ��  � ��� l ''��������  ��  ��  � ��� l  ''������  �%
	-- remove days because of delayed start
	if delayedYesOrNo is "Yes" then
		set theStore to fetch store
		
		set theCalendars to fetch calendars {} cal type list cal cloud event store theStore
		set startingDate to firstDay
		set endingDate to delayedFirstDay - (days * 1)
		set theEvents to fetch events starting date startingDate ending date endingDate searching cals theCalendars event store theStore
		
		repeat with theEvent in theEvents
			remove event event theEvent event store theStore without future events
		end repeat
		
	end if
	   � ���> 
 	 - -   r e m o v e   d a y s   b e c a u s e   o f   d e l a y e d   s t a r t 
 	 i f   d e l a y e d Y e s O r N o   i s   " Y e s "   t h e n 
 	 	 s e t   t h e S t o r e   t o   f e t c h   s t o r e 
 	 	 
 	 	 s e t   t h e C a l e n d a r s   t o   f e t c h   c a l e n d a r s   { }   c a l   t y p e   l i s t   c a l   c l o u d   e v e n t   s t o r e   t h e S t o r e 
 	 	 s e t   s t a r t i n g D a t e   t o   f i r s t D a y 
 	 	 s e t   e n d i n g D a t e   t o   d e l a y e d F i r s t D a y   -   ( d a y s   *   1 ) 
 	 	 s e t   t h e E v e n t s   t o   f e t c h   e v e n t s   s t a r t i n g   d a t e   s t a r t i n g D a t e   e n d i n g   d a t e   e n d i n g D a t e   s e a r c h i n g   c a l s   t h e C a l e n d a r s   e v e n t   s t o r e   t h e S t o r e 
 	 	 
 	 	 r e p e a t   w i t h   t h e E v e n t   i n   t h e E v e n t s 
 	 	 	 r e m o v e   e v e n t   e v e n t   t h e E v e n t   e v e n t   s t o r e   t h e S t o r e   w i t h o u t   f u t u r e   e v e n t s 
 	 	 e n d   r e p e a t 
 	 	 
 	 e n d   i f 
 	� ��� l ''��������  ��  ��  � ��� l ''������  � k e This turns off and on the iCloud Calendar in preferences so the calendars gets uploaded to the cloud   � ��� �   T h i s   t u r n s   o f f   a n d   o n   t h e   i C l o u d   C a l e n d a r   i n   p r e f e r e n c e s   s o   t h e   c a l e n d a r s   g e t s   u p l o a d e d   t o   t h e   c l o u d� ��� I '.�����
�� .sysottosnull���     TEXT� m  '*�� ��� J M o v i n g   C a l e n d a r s   f r o m   l o c a l   t o   i C l o u d��  � ��� O  /S��� k  5R�� ��� I 5:������
�� .miscactvnull��� ��� null��  ��  � ��� I ;@�����
�� .sysodelanull��� ��� nmbr� m  ;<���� ��  � ��� I AF������
�� .aevtquitnull��� ��� null��  ��  � ��� I GL������
�� .miscactvnull��� ��� null��  ��  � ���� I MR�����
�� .sysodelanull��� ��� nmbr� m  MN���� ��  ��  � m  /2���                                                                                  sprf  alis    `  Macintosh HD                   BD ����System Preferences.app                                         ����            ����  
 cu             Applications  -/:System:Applications:System Preferences.app/   .  S y s t e m   P r e f e r e n c e s . a p p    M a c i n t o s h   H D  *System/Applications/System Preferences.app  / ��  � ��� l TT��������  ��  ��  � ��� O  T���� O  Z���� O  e���� O  n���� k  w��� ��� I w|������
�� .prcsclicnull��� ��� uiel��  ��  � ���� I }������
�� .sysodelanull��� ��� nmbr� m  }~���� ��  ��  � 4  nt���
�� 
butT� m  rs���� � 4  ek���
�� 
cwin� m  ij���� � 4  Zb���
�� 
prcs� m  ^a�� ��� $ S y s t e m   P r e f e r e n c e s� m  TW���                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � ��� l ����������  ��  ��  � ��� O  ���� O  ���� O  ���� O  ���� O  ���� O  �� � O  � O  � O  � k  � 	 O ��

 I ������
�� .sysottosnull���     TEXT m  �� � 8 U n c h e c k i n g   i C l o u d   C a l e n d a r s .��    f  ��	  l �� I ��������
�� .prcsclicnull��� ��� uiel��  ��   #  turn off icloud for calendar    � :   t u r n   o f f   i c l o u d   f o r   c a l e n d a r  I ������
�� .sysodelanull��� ��� nmbr m  ������ ��    O �  I ������
�� .sysottosnull���     TEXT m  �� � 4 C h e c k i n g   i C l o u d   C a l e n d a r s .��    f  ��   l !"#! I ������
�� .prcsclicnull��� ��� uiel��  ��  " "  turn on icloud for calendar   # �$$ 8   t u r n   o n   i c l o u d   f o r   c a l e n d a r  %��% I ��&��
�� .sysodelanull��� ��� nmbr& m  ���� ��  ��   4  ����'
�� 
uiel' m  ������  4  ����(
�� 
uiel( m  ������  4  ����)
�� 
crow) m  ������   4  ����*
�� 
tabB* m  ������ � 4  ����+
�� 
scra+ m  ������ � 4  ����,
�� 
sgrp, m  ������ � 4  ����-
�� 
cwin- m  ������ � 4  ����.
�� 
prcs. m  ��// �00 $ S y s t e m   P r e f e r e n c e s� m  ��11�                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � 232 l  ��45��  4 � �
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
	   5 �66� 
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
 	3 787 O  *9:9 k  );; <=< I !������
�� .miscactvnull��� ��� null��  ��  = >��> I ")��?��
�� .sysodelanull��� ��� nmbr? m  "%���� ���  ��  : m  @@�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  8 ABA l ++��������  ��  ��  B CDC l ++�EF�  E ( " Publishes the Calendars as Public   F �GG D   P u b l i s h e s   t h e   C a l e n d a r s   a s   P u b l i cD HIH I +2�~J�}
�~ .sysottosnull���     TEXTJ m  +.KK �LL \ P u b l i s h i n g   t h e   c a l e n d a r s   b y   m a k i n g   t h e m   p u b l i c�}  I MNM O  3�OPO O  9�QRQ k  D�SS TUT l DD�|�{�z�|  �{  �z  U VWV l  DD�yXY�y  X��
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
				   Y �ZZ| 
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
 	 	 	 	W [\[ l DD�x�w�v�x  �w  �v  \ ]^] l DQ_`a_ I DQ�ubc
�u .prcskprsnull���     ctxtb m  DGdd �ee  fc �tf�s
�t 
faalf m  JM�r
�r eMdsKcmd�s  `   search bar   a �gg    s e a r c h   b a r^ hih I RW�qj�p
�q .sysodelanull��� ��� nmbrj m  RS�o�o �p  i klk l XX�n�m�l�n  �m  �l  l mnm l X_opqo I X_�kr�j
�k .prcskprsnull���     ctxtr 1  X[�i
�i 
tab �j  p #  moves focus to calendar list   q �ss :   m o v e s   f o c u s   t o   c a l e n d a r   l i s tn tut I `e�hv�g
�h .sysodelanull��� ��� nmbrv m  `a�f�f �g  u wxw l ff�e�d�c�e  �d  �c  x yzy l fu{|}{ I fu�b~
�b .prcskcodnull���     ****~ m  fi�a�a ~ �`��_
�` 
faal� J  lq�� ��^� m  lo�]
�] eMdsKopt�^  �_  | 6 0 moves focus to first calendar / option up arrow   } ��� `   m o v e s   f o c u s   t o   f i r s t   c a l e n d a r   /   o p t i o n   u p   a r r o wz ��� I v{�\��[
�\ .sysodelanull��� ��� nmbr� m  vw�Z�Z �[  � ��� l ||�Y�X�W�Y  �X  �W  � ��� r  |���� J  |~�V�V  � o      �U�U 0 calendarnames calendarNames� ��T� Y  ����S���R� k  ���� ��� Q  �����Q� k  ���� ��� r  ����� n  ����� 1  ���P
�P 
valL� n  ����� 4  ���O�
�O 
txtf� m  ���N�N � n  ����� 4  ���M�
�M 
uiel� m  ���L�L � n  ����� 4  ���K�
�K 
crow� o  ���J�J 0 i  � n  ����� 4  ���I�
�I 
outl� m  ���H�H � n  ����� 4  ���G�
�G 
scra� m  ���F�F � n  ����� 4  ���E�
�E 
splg� m  ���D�D � n  ����� 4  ���C�
�C 
splg� m  ���B�B � 4  ���A�
�A 
cwin� m  ���@�@ � o      �?�? $0 calendarnametext calendarNameText� ��>� Z  �����=�<� E  ����� o  ���;�; $0 calendarnametext calendarNameText� m  ���� ���  Y e a r� k  ���� ��� r  ����� o  ���:�: $0 calendarnametext calendarNameText� n      ���  ;  ��� o  ���9�9 0 calendarnames calendarNames� ��� l ����� I ��8��
�8 .prcskprsnull���     ctxt� m  ���� ���  f� �7��6
�7 
faal� m  ���5
�5 eMdsKcmd�6  �   search bar   � ���    s e a r c h   b a r� ��� I 	�4��3
�4 .sysodelanull��� ��� nmbr� m  �2�2 �3  � ��� l 

�1�0�/�1  �0  �/  � ��� l 
���� I 
�.��-
�. .prcskprsnull���     ctxt� 1  
�,
�, 
tab �-  � #  moves focus to calendar list   � ��� :   m o v e s   f o c u s   t o   c a l e n d a r   l i s t� ��� l �+���+  �  						delay 1   � ���  	 	 	 	 	 	 d e l a y   1� ��� l �*�)�(�*  �)  �(  � ��� l ���� I �'��&
�' .prcskcodnull���     ****� m  �%�% }�&  �   down arrow   � ���    d o w n   a r r o w� ��� I �$��#
�$ .sysodelanull��� ��� nmbr� m  �"�" �#  � ��� l   �!� ��!  �   �  � ��� O   =��� l 7<���� I 7<���
� .prcsclicnull��� ��� uiel�  �  �   click Edit Menu   � ���     c l i c k   E d i t   M e n u� n   4��� 4  -4��
� 
menE� m  03�� ���  E d i t� n   -��� 4  &-��
� 
mbri� m  ),�� ���  E d i t� 4   &��
� 
mbar� m  $%�� � ��� l >>� �     						delay 1    �  	 	 	 	 	 	 d e l a y   1�  O  >b l \a	 I \a���
� .prcsclicnull��� ��� uiel�  �   + % click Share Calendar under Edit Menu   	 �

 J   c l i c k   S h a r e   C a l e n d a r   u n d e r   E d i t   M e n u n  >Y 4  RY�
� 
uiel m  UX��  n  >R 4  KR�
� 
menE m  NQ �  E d i t n  >K 4  DK�
� 
mbri m  GJ �  E d i t 4  >D�
� 
mbar m  BC��   l cc��    						delay 1    �  	 	 	 	 	 	 d e l a y   1  O  c� !  l ��"#$" I �����

� .prcsclicnull��� ��� uiel�  �
  # ( " click the share calendar checkbox   $ �%% D   c l i c k   t h e   s h a r e   c a l e n d a r   c h e c k b o x! n  c�&'& 4  ���	(
�	 
chbx( m  ���� ' n  c�)*) 4  ���+
� 
popv+ m  ���� * n  c�,-, 4  }��.
� 
crow. o  ���� 0 i  - n  c}/0/ 4  x}�1
� 
outl1 m  {|�� 0 n  cx232 4  sx�4
� 
scra4 m  vw� �  3 n  cs565 4  ns��7
�� 
splg7 m  qr���� 6 n  cn898 4  in��:
�� 
splg: m  lm���� 9 4  ci��;
�� 
cwin; m  gh����  <=< l ����>?��  >  						delay 1   ? �@@  	 	 	 	 	 	 d e l a y   1= ABA O  ��CDC I ��������
�� .prcsclicnull��� ��� uiel��  ��  D n  ��EFE 4  ����G
�� 
butTG m  ��HH �II  D o n eF n  ��JKJ 4  ����L
�� 
popvL m  ������ K n  ��MNM 4  ����O
�� 
crowO o  ������ 0 i  N n  ��PQP 4  ����R
�� 
outlR m  ������ Q n  ��STS 4  ����U
�� 
scraU m  ������ T n  ��VWV 4  ����X
�� 
splgX m  ������ W n  ��YZY 4  ����[
�� 
splg[ m  ������ Z 4  ����\
�� 
cwin\ m  ������ B ]^] l ����_`��  _ 7 1					keystroke "f" using command down # seach bar   ` �aa b 	 	 	 	 	 k e y s t r o k e   " f "   u s i n g   c o m m a n d   d o w n   #   s e a c h   b a r^ bcb I ����d��
�� .sysodelanull��� ��� nmbrd m  ������ ��  c efe l ��ghig I ����j��
�� .prcskprsnull���     ctxtj 1  ����
�� 
tab ��  h #  moves focus to calendar list   i �kk :   m o v e s   f o c u s   t o   c a l e n d a r   l i s tf l��l I ����m��
�� .sysodelanull��� ��� nmbrm m  ������ ��  ��  �=  �<  �>  � R      ������
�� .ascrerr ****      � ****��  ��  �Q  � n��n l ����������  ��  ��  ��  �S 0 i  � m  ������ � I ����o��
�� .corecnte****       ****o n ��pqp 2 ����
�� 
crowq n  ��rsr 4  ����t
�� 
outlt m  ������ s n  ��uvu 4  ����w
�� 
scraw m  ������ v n  ��xyx 4  ����z
�� 
splgz m  ������ y n  ��{|{ 4  ����}
�� 
splg} m  ������ | 4  ����~
�� 
cwin~ m  ������ ��  �R  �T  R 4  9A��
�� 
prcs m  =@�� ���  C a l e n d a rP m  36���                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  N ��� l ����������  ��  ��  � ��� l ��������  � N H Grabs the published calendars and writes the URLs to a numbers document   � ��� �   G r a b s   t h e   p u b l i s h e d   c a l e n d a r s   a n d   w r i t e s   t h e   U R L s   t o   a   n u m b e r s   d o c u m e n t� ��� O  ���� I � ������
�� .miscactvnull��� ��� null��  ��  � m  �����                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��  � ��� I 	�����
�� .sysottosnull���     TEXT� m  �� ��� X G r a b b i n g   t h e   U R L s   f o r   t h e s e   p u b l i c   c a l e n d a r s��  � ��� l 

��������  ��  ��  � ��� r  
��� J  
����  � o      ���� 0 calendarurls calendarURLs� ��� O  ���� O  ���� k  "��� ��� I "/����
�� .prcskprsnull���     ctxt� m  "%�� ���  f� �����
�� 
faal� m  (+��
�� eMdsKcmd��  � ��� I 05�����
�� .sysodelanull��� ��� nmbr� m  01���� ��  � ��� I 6=�����
�� .prcskprsnull���     ctxt� 1  69��
�� 
tab ��  � ��� I >C�����
�� .sysodelanull��� ��� nmbr� m  >?����  ��  � ��� I DS����
�� .prcskcodnull���     ****� m  DG���� ~� �����
�� 
faal� J  JO�� ���� m  JM��
�� eMdsKopt��  ��  � ��� I TY�����
�� .sysodelanull��� ��� nmbr� m  TU����  ��  � ��� l ZZ��������  ��  ��  � ���� Y  Z��������� k  j��� ��� I jq�����
�� .prcskcodnull���     ****� m  jm���� }��  � ��� I rw�����
�� .sysodelanull��� ��� nmbr� m  rs����  ��  � ��� I x�����
�� .prcskprsnull���     ctxt� m  x{�� ���  i� �����
�� 
faal� m  ~���
�� eMdsKcmd��  � ��� I �������
�� .sysodelanull��� ��� nmbr� m  ������  ��  � ��� O  ����� O  ����� k  ���� ��� r  ����� n  ����� 1  ����
�� 
valL� 4  �����
�� 
txtf� m  ������ � o      ���� 0 calendarname calendarName� ���� O  ����� O  ����� r  ����� 1  ����
�� 
valL� o      ���� 0 theurl theURL� 4  �����
�� 
txta� m  ������ � 4  �����
�� 
scra� m  ������ ��  � 4  �����
�� 
sheE� m  ������ � 4  �����
�� 
cwin� m  ������ � ��� I �����~
� .prcskprsnull���     ctxt� o  ���}
�} 
ret �~  � ��� I ���|��{
�| .sysodelanull��� ��� nmbr� m  ���z�z  �{  � ��y� r  ����� J  ���� ��� o  ���x�x 0 calendarname calendarName� ��w� o  ���v�v 0 theurl theURL�w  � n      ���  ;  ��� o  ���u�u 0 calendarurls calendarURLs�y  �� 0 i  � m  ]^�t�t � I ^e�s��r
�s .corecnte****       ****� o  ^a�q�q 0 calendarnames calendarNames�r  ��  ��  � 4  �p�
�p 
prcs� m  �� ���  C a l e n d a r� m  ���                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  � �	 � l ���o�n�m�o  �n  �m  	  			 I ���l	�k
�l .sysottosnull���     TEXT	 m  ��		 �		 p P a s t i n g   p u b l i c   c a l e n d a r   l i n k s   i n t o   a   N u m b e r s   s p r e a d s h e e t�k  	 			 O  ��				 k  ��	
	
 			 I ��j�i�h
�j .miscactvnull��� ��� null�i  �h  	 			 I �g�f	
�g .corecrel****      � null�f  	 �e	�d
�e 
kocl	 m  �c
�c 
docu�d  	 	�b	 O  �			 O  �			 O  "�			 k  +�		 			 O  +R			 k  4Q		 			 r  4B	 	!	  m  47	"	" �	#	#  C a l e n d a r   N a m e	! n      	$	%	$ 1  =A�a
�a 
NMCv	% 4  7=�`	&
�` 
NmCl	& m  ;<�_�_ 	 	'�^	' r  CQ	(	)	( m  CF	*	* �	+	+  U R L	) n      	,	-	, 1  LP�]
�] 
NMCv	- 4  FL�\	.
�\ 
NmCl	. m  JK�[�[ �^  	 4  +1�Z	/
�Z 
NMRw	/ m  /0�Y�Y 	 	0	1	0 l SS�X�W�V�X  �W  �V  	1 	2�U	2 Y  S�	3�T	4	5�S	3 k  c�	6	6 	7	8	7 r  ct	9	:	9 n  cp	;	<	; 4  kp�R	=
�R 
cobj	= m  no�Q�Q 	< n  ck	>	?	> 4  fk�P	@
�P 
cobj	@ o  ij�O�O 0 i  	? o  cf�N�N 0 calendarurls calendarURLs	: o      �M�M 0 calendarname calendarName	8 	A	B	A r  u�	C	D	C n  u�	E	F	E 4  }��L	G
�L 
cobj	G m  ���K�K 	F n  u}	H	I	H 4  x}�J	J
�J 
cobj	J o  {|�I�I 0 i  	I o  ux�H�H 0 calendarurls calendarURLs	D o      �G�G 0 theurl theURL	B 	K	L	K Z  ��	M	N�F�E	M H  ��	O	O l ��	P�D�C	P I ���B	Q�A
�B .coredoexnull���     ****	Q 4  ���@	R
�@ 
NMRw	R l ��	S�?�>	S [  ��	T	U	T o  ���=�= 0 i  	U m  ���<�< �?  �>  �A  �D  �C  	N I ���;�:	V
�; .corecrel****      � null�:  	V �9	W�8
�9 
kocl	W m  ���7
�7 
NMRw�8  �F  �E  	L 	X	Y	X O  ��	Z	[	Z k  ��	\	\ 	]	^	] r  ��	_	`	_ o  ���6�6 0 calendarname calendarName	` n      	a	b	a 1  ���5
�5 
NMCv	b 4  ���4	c
�4 
NmCl	c m  ���3�3 	^ 	d�2	d r  ��	e	f	e o  ���1�1 0 theurl theURL	f n      	g	h	g 1  ���0
�0 
NMCv	h 4  ���/	i
�/ 
NmCl	i m  ���.�. �2  	[ 4  ���-	j
�- 
NMRw	j l ��	k�,�+	k [  ��	l	m	l o  ���*�* 0 i  	m m  ���)�) �,  �+  	Y 	n�(	n l ���'�&�%�'  �&  �%  �(  �T 0 i  	4 m  VW�$�$ 	5 I W^�#	o�"
�# .corecnte****       ****	o o  WZ�!�! 0 calendarurls calendarURLs�"  �S  �U  	 4  "(� 	p
�  
NmTb	p m  &'�� 	 4  �	q
� 
NmSh	q m  �� 	 4  �	r
� 
docu	r m  �� �b  		 m  ��	s	s�                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  	 	t	u	t l ������  �  �  	u 	v�	v l ������  �  �  �  ��  k k  ��	w	w 	x	y	x l ���	z	{�  	z 9 3 Outlook --> Incomplete code, needs repeat function   	{ �	|	| f   O u t l o o k   - - >   I n c o m p l e t e   c o d e ,   n e e d s   r e p e a t   f u n c t i o n	y 	}	~	} l ���		��  	 B < Creates all the events on the first day but does not repeat   	� �	�	� x   C r e a t e s   a l l   t h e   e v e n t s   o n   t h e   f i r s t   d a y   b u t   d o e s   n o t   r e p e a t	~ 	�	�	� O  ��	�	�	� k  ��	�	� 	�	�	� I �����
� .miscactvnull��� ��� null�  �  	� 	�	�	� l ������  �  �  	� 	�	�	� I ��	��

� .coredelonull���     obj 	� l � 	��	�	� 6 � 	�	�	� 2 ���
� 
cCFo	� E  ��	�	�	� 1  ���
� 
pnam	� m  ��	�	� �	�	�  Y e a r�	  �  �
  	� 	�	�	� l ����  �  �  	� 	�	�	� r  "	�	�	� 6 	�	�	� 4 �	�
� 
cCFo	� m  	
�� 	� E  	�	�	� n  	�	�	� m  � 
�  
emad	� 2 ��
�� 
cAct	� m  	�	� �	�	�  w i n t e c	� o      ���� 0 wintecaccount wintecAccount	� 	���	� O  #�	�	�	� k  )�	�	� 	�	�	� l ))��������  ��  ��  	� 	�	�	� Y  )u	���	�	���	� Z  9p	�	�����	� H  9L	�	� l 9K	�����	� I 9K��	���
�� .coredoexnull���     ****	� l 9G	�����	� 4  9G��	�
�� 
cCFo	� l =F	�����	� n  =F	�	�	� 4  @E��	�
�� 
cobj	� o  CD���� 0 i  	� o  =@���� 0 
classnames 
classNames��  ��  ��  ��  ��  ��  ��  	� I Ol����	�
�� .corecrel****      � null��  	� ��	�	�
�� 
kocl	� m  SV��
�� 
cCFo	� ��	���
�� 
prdt	� K  Yf	�	� ��	���
�� 
pnam	� n  \d	�	�	� 4  _d��	�
�� 
cobj	� o  bc���� 0 i  	� o  \_���� 0 
classnames 
classNames��  ��  ��  ��  �� 0 i  	� m  ,-���� 	� I -4��	���
�� .corecnte****       ****	� o  -0���� 0 
classnames 
classNames��  ��  	� 	�	�	� l vv��������  ��  ��  	� 	�	�	� X  v�	���	�	� Z  ��	�	���	�	� I ����	���
�� .coredoexnull���     ****	� 4  ����	�
�� 
cCFo	� l ��	�����	� n  ��	�	�	� 4  ����	�
�� 
cobj	� m  ������ 	� o  ������ 
0 lesson  ��  ��  ��  	� O  ��	�	�	� I ������	�
�� .corecrel****      � null��  	� ��	�	�
�� 
kocl	� m  ����
�� 
cEvt	� ��	���
�� 
prdt	� K  ��	�	� ��	�	�
�� 
subj	� n  ��	�	�	� 4  ����	�
�� 
cobj	� m  ������ 	� o  ������ 
0 lesson  	� ��	�	�
�� 
iloc	� n  ��	�	�	� 4  ����	�
�� 
cobj	� m  ������ 	� o  ������ 
0 lesson  	� ��	�	�
�� 
offs	� n  ��	�	�	� 4  ����	�
�� 
cobj	� m  ������ 	� o  ������ 
0 lesson  	� ��	�	�
�� 
endT	� [  ��	�	�	� l ��	�����	� n  ��	�	�	� 4  ����	�
�� 
cobj	� m  ������ 	� o  ������ 
0 lesson  ��  ��  	� l ��	�����	� ]  ��	�	�	� 1  ����
�� 
min 	� m  ������ P��  ��  	� ��	���
�� 
pHsR	� m  ����
�� boovfals��  ��  	� 4  ����	�
�� 
cCFo	� l ��	�����	� n  ��	�	�	� 4  ����	�
�� 
cobj	� m  ������ 	� o  ������ 
0 lesson  ��  ��  ��  	� k  ��	�	� 	�	�	� r  �	�	�	� n  � 	�
 	� 2 � ��
�� 
cwor
  n  ��


 4  ����

�� 
cobj
 m  ������ 
 o  ������ 
0 lesson  	� o      ���� 0 problem  	� 


 r  :


 b  6

	
 b  +




 b  '


 b  


 b  


 b  


 n  


 4  ��

�� 
cobj
 m  ���� 
 o  ���� 0 problem  
 m  

 �

   
 n  


 4  ��

�� 
cobj
 m  ���� 
 o  ���� 0 problem  
 m  

 �

   
 n  &


 4  !&��
 
�� 
cobj
  m  $%���� 
 o  !���� 0 problem  
 m  '*
!
! �
"
"   
	 n  +5
#
$
# 4  .5��
%
�� 
cobj
% m  14���� 
$ o  +.���� 0 problem  
 o      ���� "0 neweventsummary newEventSummary
 
&��
& O  ;�
'
(
' I F�����
)
�� .corecrel****      � null��  
) ��
*
+
�� 
kocl
* m  JM��
�� 
cEvt
+ ��
,��
�� 
prdt
, K  P�
-
- ��
.
/
�� 
subj
. n  SY
0
1
0 4  TY��
2
�� 
cobj
2 m  WX���� 
1 o  ST���� 
0 lesson  
/ ��
3
4
�� 
iloc
3 n  \b
5
6
5 4  ]b��
7
�� 
cobj
7 m  `a���� 
6 o  \]���� 
0 lesson  
4 ��
8
9
�� 
offs
8 n  ek
:
;
: 4  fk��
<
�� 
cobj
< m  ij���� 
; o  ef���� 
0 lesson  
9 ��
=
>
�� 
endT
= [  n|
?
@
? l nt
A����
A n  nt
B
C
B 4  ot��
D
�� 
cobj
D m  rs�� 
C o  no�~�~ 
0 lesson  ��  ��  
@ l t{
E�}�|
E ]  t{
F
G
F 1  tw�{
�{ 
min 
G m  wz�z�z P�}  �|  
> �y
H�x
�y 
pHsR
H m  ��w
�w boovfals�x  ��  
( 4  ;C�v
I
�v 
cCFo
I o  ?B�u�u "0 neweventsummary newEventSummary��  �� 
0 lesson  	� o  y|�t�t 0 lessons8  	� 
J�s
J l ���r�q�p�r  �q  �p  �s  	� o  #&�o�o 0 wintecaccount wintecAccount��  	� m  ��
K
K�                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  	� 
L
M
L l ���n�m�l�n  �m  �l  
M 
N�k
N O  ��
O
P
O k  ��
Q
Q 
R
S
R r  ��
T
U
T J  ���j�j  
U o      �i�i 0 	allevents 	allEvents
S 
V
W
V r  ��
X
Y
X 6 ��
Z
[
Z 4 ���h
\
�h 
cCFo
\ m  ���g�g 
[ E  ��
]
^
] n  ��
_
`
_ m  ���f
�f 
emad
` 2 ���e
�e 
cAct
^ m  ��
a
a �
b
b  w i n t e c
Y o      �d�d 0 wintecaccount wintecAccount
W 
c
d
c O  �&
e
f
e k  �%
g
g 
h
i
h r  ��
j
k
j 6 ��
l
m
l 2 ���c
�c 
cCFo
m E  ��
n
o
n 1  ���b
�b 
pnam
o m  ��
p
p �
q
q  Y e a r
k o      �a�a "0 winteccalendars wintecCalendars
i 
r�`
r X  �%
s�_
t
s k  � 
u
u 
v
w
v r  ��
x
y
x n  ��
z
{
z 2 ���^
�^ 
cEvt
{ o  ���]�] 0 	acalendar 	aCalendar
y o      �\�\ 0 	theevents 	theEvents
w 
|�[
| X  � 
}�Z
~
} r  

�
 o  �Y�Y 0 anevent anEvent
� n      
�
�
�  ;  
� o  �X�X 0 	allevents 	allEvents�Z 0 anevent anEvent
~ o  �W�W 0 	theevents 	theEvents�[  �_ 0 	acalendar 	aCalendar
t o  ���V�V "0 winteccalendars wintecCalendars�`  
f o  ���U�U 0 wintecaccount wintecAccount
d 
��T
� X  '�
��S
�
� k  =�
�
� 
�
�
� n  =C
�
�
� 1  >B�R
�R 
subj
� o  =>�Q�Q 0 anevent anEvent
� 
�
�
� r  DM
�
�
� n  DI
�
�
� 1  EI�P
�P 
offs
� o  DE�O�O 0 anevent anEvent
� o      �N�N 0 startday startDay
� 
�
�
� r  N]
�
�
� c  NY
�
�
� n  NU
�
�
� m  QU�M
�M 
wkdy
� o  NQ�L�L 0 startday startDay
� m  UX�K
�K 
ctxt
� o      �J�J 0 startday startDay
� 
�
�
� l ^^�I�H�G�I  �H  �G  
� 
�
�
� l ^^�F
�
��F  
� L F next edit -- change start date into a variable instead of a hard date   
� �
�
� �   n e x t   e d i t   - -   c h a n g e   s t a r t   d a t e   i n t o   a   v a r i a b l e   i n s t e a d   o f   a   h a r d   d a t e
� 
�
�
� Z  ^�
�
�
��E
� = ^e
�
�
� o  ^a�D�D 0 startday startDay
� m  ad
�
� �
�
�  M o n d a y
� r  h�
�
�
� K  h�
�
� �C
�
�
�C 
ReDw
� K  kq
�
� �B
��A
�B 
rDMo
� m  no�@
�@ boovtrue�A  
� �?
�
�
�? 
ReED
� K  t�
�
� �>
�
�
�> 
pEnT
� m  wz�=
�= eEDTeEDt
� �<
��;
�< 
peDa
� o  }��:�: 0 enddate endDate�;  
� �9
�
�
�9 
tStr
� m  ��
�
� ldt     �a� 
� �8
�
�
�8 
ReOi
� m  ���7�7 
� �6
��5
�6 
ReTy
� m  ���4
�4 eRPTeRwp�5  
� n      
�
�
� m  ���3
�3 
rREc
� o  ���2�2 0 anevent anEvent
� 
�
�
� = ��
�
�
� o  ���1�1 0 startday startDay
� m  ��
�
� �
�
�  T u e s d a y
� 
�
�
� r  ��
�
�
� K  ��
�
� �0
�
�
�0 
ReDw
� K  ��
�
� �/
��.
�/ 
rDTu
� m  ���-
�- boovtrue�.  
� �,
�
�
�, 
ReED
� K  ��
�
� �+
�
�
�+ 
pEnT
� m  ���*
�* eEDTeEDt
� �)
��(
�) 
peDa
� o  ���'�' 0 enddate endDate�(  
� �&
�
�
�& 
tStr
� m  ��
�
� ldt     �c0�
� �%
�
�
�% 
ReOi
� m  ���$�$ 
� �#
��"
�# 
ReTy
� m  ���!
�! eRPTeRwp�"  
� n      
�
�
� m  ��� 
�  
rREc
� o  ���� 0 anevent anEvent
� 
�
�
� = ��
�
�
� o  ���� 0 startday startDay
� m  ��
�
� �
�
�  W e d n e s d a y
� 
�
�
� r  �$
�
�
� K  �
�
� �
�
�
� 
ReDw
� K  ��
�
� �
��
� 
rDWe
� m  ���
� boovtrue�  
� �
�
�
� 
ReED
� K  �

�
� �
�
�
� 
pEnT
� m  � �
� eEDTeEDt
� �
��
� 
peDa
� o  �� 0 enddate endDate�  
� �
�
�
� 
tStr
� m  
�
� ldt     �d� 
� �
�
�
� 
ReOi
� m  �� 
� �
��
� 
ReTy
� m  �
� eRPTeRwp�  
� n      
�
�
� m  #�
� 
rREc
� o  �� 0 anevent anEvent
� 
�
�
� = '.
�
�
� o  '*�� 0 startday startDay
� m  *-
�
� �
�
�  T h u r s d a y
� 
�
�
� r  1g   K  1a �

�
 
ReDw K  4: �	�
�	 
rDTh m  78�
� boovtrue�   �
� 
ReED K  =M		 �

� 
pEnT
 m  @C�
� eEDTeEDt ��
� 
peDa o  FI�� 0 enddate endDate�   � 
�  
tStr m  PS ldt     �eӀ ��
�� 
ReOi m  VW����  ����
�� 
ReTy m  Z]��
�� eRPTeRwp��   n       m  bf��
�� 
rREc o  ab���� 0 anevent anEvent
�  = jq o  jm���� 0 startday startDay m  mp �  F r i d a y �� r  t� K  t� �� 
�� 
ReDw K  w}!! ��"��
�� 
rDFr" m  z{��
�� boovtrue��    ��#$
�� 
ReED# K  ��%% ��&'
�� 
pEnT& m  ����
�� eEDTeEDt' ��(��
�� 
peDa( o  ������ 0 enddate endDate��  $ ��)*
�� 
tStr) m  ��++ ldt     �g% * ��,-
�� 
ReOi, m  ������ - ��.��
�� 
ReTy. m  ����
�� eRPTeRwp��   n      /0/ m  ����
�� 
rREc0 o  ������ 0 anevent anEvent��  �E  
� 1��1 l ����������  ��  ��  ��  �S 0 anevent anEvent
� o  *-���� 0 	allevents 	allEvents�T  
P m  ��22�                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  �k  ��  ��  g 343 l     ��������  ��  ��  4 565 l ��7����7 I ����8��
�� .sysottosnull���     TEXT8 m  ��99 �::  F i n i s h e d .��  ��  ��  6 ;<; l     ��������  ��  ��  < =>= i   " %?@? I      ��A���� 0 simple_sort  A B��B o      ���� 0 my_list  ��  ��  @ k     uCC DED r     FGF J     ����  G l     H����H o      ���� 0 
index_list  ��  ��  E IJI r    	KLK J    ����  L l     M����M o      ���� 0 sorted_list  ��  ��  J NON U   
 rPQP k    mRR STS r    UVU m    WW �XX  V l     Y����Y o      ���� 0 low_item  ��  ��  T Z[Z Y    c\��]^��\ Z   ) ^_`����_ H   ) -aa E  ) ,bcb l  ) *d����d o   ) *���� 0 
index_list  ��  ��  c o   * +���� 0 i  ` k   0 Zee fgf r   0 8hih c   0 6jkj n   0 4lml 4   1 4��n
�� 
cobjn o   2 3���� 0 i  m o   0 1���� 0 my_list  k m   4 5��
�� 
ctxti o      ���� 0 	this_item  g o��o Z   9 Zpqr��p =  9 <sts l  9 :u����u o   9 :���� 0 low_item  ��  ��  t m   : ;vv �ww  q k   ? Fxx yzy r   ? B{|{ o   ? @���� 0 	this_item  | l     }����} o      ���� 0 low_item  ��  ��  z ~��~ r   C F� o   C D���� 0 i  � l     ������ o      ���� 0 low_item_index  ��  ��  ��  r ��� A I L��� o   I J���� 0 	this_item  � l  J K������ o   J K���� 0 low_item  ��  ��  � ���� k   O V�� ��� r   O R��� o   O P���� 0 	this_item  � l     ������ o      ���� 0 low_item  ��  ��  � ���� r   S V��� o   S T���� 0 i  � l     ������ o      ���� 0 low_item_index  ��  ��  ��  ��  ��  ��  ��  ��  �� 0 i  ] m    ���� ^ l   $������ n    $��� m   ! #��
�� 
nmbr� n   !��� 2   !��
�� 
cobj� o    ���� 0 my_list  ��  ��  ��  [ ��� r   d h��� l  d e������ o   d e���� 0 low_item  ��  ��  � l     ������ n      ���  ;   f g� o   e f���� 0 sorted_list  ��  ��  � ���� r   i m��� l  i j������ o   i j���� 0 low_item_index  ��  ��  � l     ������ n      ���  ;   k l� l  j k������ o   j k���� 0 
index_list  ��  ��  ��  ��  ��  Q l   ������ l   ������ n    ��� m    ��
�� 
nmbr� n   ��� 2   ��
�� 
cobj� o    ���� 0 my_list  ��  ��  ��  ��  O ���� L   s u�� l  s t����� o   s t�~�~ 0 sorted_list  ��  �  ��  > ��� l     �}�|�{�}  �|  �{  � ��z� l     �y�x�w�y  �x  �w  �z       �v�����v  � �u�t�s
�u 
pimr�t 0 simple_sort  
�s .aevtoappnull  �   � ****� �r��r �  ���� �q �p
�q 
vers�p  � �o��n
�o 
cobj� ��   �m
�m 
osax�n  � �l��
�l 
cobj� ��   �k 
�k 
scpt� �j �i
�j 
vers�i  � �h@�g�f���e�h 0 simple_sort  �g �d��d �  �c�c 0 my_list  �f  � �b�a�`�_�^�]�\�b 0 my_list  �a 0 
index_list  �` 0 sorted_list  �_ 0 low_item  �^ 0 i  �] 0 	this_item  �\ 0 low_item_index  � �[�ZW�Yv
�[ 
cobj
�Z 
nmbr
�Y 
ctxt�e vjvE�OjvE�O g��-�,Ekh�E�O Hk��-�,Ekh �� /��/�&E�O��  �E�O�E�Y �� �E�O�E�Y hY h[OY��O��6FO��6F[OY��O�� �X��W�V���U
�X .aevtoappnull  �   � ****� k    ���  *��  0��  I��  R��  X��  ���  ���  ���  ���  ���  ��� �� �� #�� -�� 9�� B�� L�� Y�� h�� o�� v�� ��� ��� ��� ��� ��� '�� ,�� p�� u�� ��� ��� �� �� ��� ��� 
�� �� f�� 5�T�T  �W  �V  � �S�R�Q�P�O�N�M�L�S 0 columnindex columnIndex�R 0 rowindex rowIndex�Q 0 i  �P 0 j  �O 
0 lesson  �N 0 room  �M 0 	acalendar 	aCalendar�L 0 anevent anEvent�& .�K 4�J N�I V�H�G e�F�E�D�C�B�A�@ �?�>�=�<�; � � � � � ��: � � ��9�8�7�6�5�4�3�2>�1�0�/�.b�-�,�+�*��)�(�'�&�%�$��#�"��!� �������� $*/���Q_s��������������
9E\�	��������������&7E^l������ ����,n�����������������������������������������Ue���� ������ �������� ������������ ��������� ������������������/������������K�d������������������������������������������H����������������		"	*
K��	�����	���������������


!��
a
p������
���������������
���������
���
�
���
�
�����+9
�K .sysottosnull���     TEXT
�J .sysodlogaskr        TEXT�I &0 chosencalendarapp chosenCalendarApp
�H .misccurdldt    ��� null�G  0 thecurrentdate theCurrentDate
�F 
dtxt
�E 
dstr
�D 
spac
�C 
tstr
�B 
rslt
�A 
ttxt�@ 0 thetext theText
�? 
ldt �> 0 firstday firstDay�=  �<  
�; .sysobeepnull��� ��� long�: 0 thedate theDate
�9 
days
�8 
year�7 0 a  
�6 
mnth
�5 
nmbr�4 0 b  
�3 
day �2 0 c  �1 0 e  �0 0 enddate endDate�/ 0 thelist theList
�. 
txdl
�- 
cobj�, 
0 tid TID
�+ 
ctxt�* "0 thelistasstring theListAsString�) "0 endofrecurrence endOfRecurrence�( 0 
classnames 
classNames
�' .miscactvnull��� ��� null
�& 
docu
�% 
pnam�$ "0 listofdocuments listOfDocuments
�# .gtqpchltns    @   @ ns  �"  0 chosendocument chosenDocument
�! 
NmSh�  0 listofsheets listOfSheets� 0 chosensheet chosenSheet
� 
NmTb� 
� 
NMCo
� 
NmCl
� 
NMCv� 0 	yearvalue 	yearValue� � %
� 
msng�  � 0 	roomindex 	roomIndex
� 
NMad� 0 columnnumber columnNumber
� 
NMRw� 0 	rownumber 	rowNumber
� .corecnte****       ****� 0 rowcount rowCount� 0 roomnumbers roomNumbers� 0 lessons  � "� 0 lessons3  
�
 
kocl�	 0 lessons4  � 0 	inputtime 	inputTime
� 
cwor� � 0 thetime theTime� � 0 thetime2 theTime2� 0 lessons5  � 0 lessons6  �  0 lessons7  �� 0 
datestring 
dateString�� 0 lessons8  
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
uiel�� �
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
�� 
ReDw
�� 
rDMo
�� 
ReED
�� 
pEnT
�� eEDTeEDt
�� 
peDa
�� 
tStr
�� 
ReOi
�� 
ReTy
�� eRPTeRwp
�� 
rREc
�� 
rDTu
�� 
rDWe
�� 
rDTh
�� 
rDFr�U��j O�j O�E�O�j O UhZ*j E�O����,�%��,%l O��,E` O !_ a  *a _ /E` OY hW X  *j [OY��Oa _ %a %j Oa �_ �,�%_ �,%l Oa j O WhZ*j E�Oa ���,�%��,%l O��,E` O !_ a  *a _ /E` OY hW X  *j [OY��Oa _ %a %j Oa  �_ �,�%_ �,%l O_ _ !k E` O_ a ",E` #O_ a $,a %&E` &O_ a ',E` (Oa )E` *O_ E` +O_ #_ &_ (mvE` ,O*a -,a .lvE[a /k/E` 0Z[a /l/*a -,FZO_ ,a 1&E` 2O_ 0*a -,FO_ *_ 2lvE` ,O*a -,a 3lvE[a /k/E` 0Z[a /l/*a -,FZO_ ,a 1&E` 4O_ 0*a -,FOjvE` 5Oa 6�*j 7O*a 8-a 9,E` :O) 	a ;j UO_ :j <E` =O_ =a /k/E` =O*a 8_ =/�) 	a >j UO*a ?-a 9,E` @O_ @j <E` AO_ Aa /k/E` AO) 	a Bj UO*a ?_ A/ �*a Ck/ � �la Dkh   Hkkh*a E�/a Fm/a G,E` HO_ Ha I a JE` HY _ Ha K a LE` HY [OY��O oa Ma Nkh *a E�/a F�/a G,a O F_ 5_ Ha P%*a E�/a F�/a G,% !_ Ha Q%*a E�/a F�/a G,%_ 56FY hOPY hOP[OY��OP[OY�>UUO) 	a Rj UO*a ?_ A/Q*a Ck/G*a Fk/a S[a G,\Za T81E` UO_ Ua E,a V,E` WO_ Ua X,a V,E` YO*a E_ W/a F-j ZE` [OjvE` \O c_ Yk_ [kh *a E_ Wk/a F�/a G,a O 4*a E_ Wk/a F�/a G,*a E_ W/a F�/a G,a 1&lv_ \6FY hOP[OY��OjvE` ]O �la Dkh  qa Ma ^kh *a E�/a F�/a G,a O J*a E�/a Fl/a G,*a E�/a Fm/a G,*a Ek/a F�/a G,*a E�/a F�/a G,a Mv_ ]6FY h[OY��[OY��UUUUOjvE` _O j_ ][a `a /l Zkh �a /l/a a %�a /k/a b�a /m/�a /a M/a Mv_ _6FY "�a /k/a c�a /m/�a /a M/a Mv_ _6F[OY��OjvE` dO �_ _[a `a /l Zkh �a /m/E` eO_ ea fa M/a g%_ ea fa h/%E` iO_ ia fk/a %&a j *_ ia fk/a ja k%_ ia fl/%a l%a 1&E` mY #_ ia fk/a n%_ ia fl/%a o%a 1&E` mO�a /k/�a /l/_ m�a /a M/a Mv_ d6F[OY�TOjvE` pO @_ d[a `a /l Zkh �a /l/a q%�a /a M/%�a /k/�a /m/mv_ p6F[OY��OjvE` rO_ p[a `a /l Zkh �a /l/a s "�a /k/_ �,a t%�a /m/%lv_ r6FY Ѥa /l/a u (�a /k/_ _ !k �,a v%�a /m/%lv_ r6FY ��a /l/a w (�a /k/_ _ !l �,a x%�a /m/%lv_ r6FY k�a /l/a y (�a /k/_ _ !m �,a z%�a /m/%lv_ r6FY 8�a /l/a { *�a /k/_ _ !a M �,a |%�a /m/%lv_ r6FY h[OY��OjvE` }O B_ r[a `a /l Zkh �a /l/E` ~O*a _ ~/E` O�a /k/_ lv_ }6F[OY��OjvE` O �_ }[a `a /l Zkh �a /k/a � �a /k/�a /l/a �mv_ 6FY J G_ \[a `a /l Zkh �a /k/�a /k/ �a /k/�a /l/�a /l/mv_ 6FY h[OY��[OY��O�a � �a � 2*j 7O) 	a �j UO*a �-a S[a 9,\Za �@1 *j �UUOa �j Oa �s*j 7O >k_ 5j Zkh *a �_ 5a /�/E/j � *a �_ 5a /�/l �Y h[OY��O) 	a �j UO_ [a `a /l Zkh *a ��a /k/E/j � [*a ��a /k/E/ H*a `a �a �a ��a /k/a ��a /m/a ��a /l/a ��a /l/_ �a � a �_ 4a �a M �UY ��a /k/a f-E` �O_ �a /k/a �%_ �a /l/%a �%_ �a /m/%a �%_ �a /a M/%E` �O*a �_ �/ H*a `a �a �a ��a /k/a ��a /m/a ��a /l/a ��a /l/_ �a � a �_ 4a �a M �U[OY��UO) 	a �j UOa � *a �-a 9,a S[a 9,\Za �@1E` �UO)a �a �/ *j �UE` �O_ E` �O_ �_ !a h E` +O)a �a �/ _ �a �a �a �_ �a M �UE` �O)a �a �/ !*a �_ �a �_ +a �_ �a �_ �a � �UE` �O \k_ �j Zkh )a �a �/ *a �_ �a /�/a �a �a M �UE` �O)a �a �/ *a �_ �a �_ �a M �U[OY��Oa �j Oa � *j 7Omj �O*j �O*j 7Olj �UOa � -*a �a �/ !*a �k/ *a �k/ *j �Omj �UUUUOa � �*a �a �/ }*a �k/ s*a �k/ i*a �k/ _*a �k/ U*a �a �/ I*a �k/ ?*a �k/ 5) 	a �j UO*j �Oa hj �O) 	a �j UO*j �Okj �UUUUUUUUUOa � *j 7Oa �j �UOa �j Oa ��*a �a �/�a �a �a �l �Okj �O_ �j �Okj �Oa �a �a �kvl �Okj �OjvE` �Onl*a �k/a �k/a �k/a �k/a �k/a �-j Zkh 6*a �k/a �k/a �k/a �k/a �k/a Ǣ/a �k/a �k/a �,E` �O_ �a � �_ �_ �6FOa �a �a �l �Okj �O_ �j �Oa �j �Okj �O*a �k/a �a �/a �a �/ *j �UO*a �k/a �a �/a �a �/a �a D/ *j �UO*a �k/a �k/a �k/a �k/a �k/a Ǣ/a �k/a �k/ *j �UO*a �k/a �k/a �k/a �k/a �k/a Ǣ/a �k/a �a �/ *j �UOkj �O_ �j �Okj �Y hW X  hOP[OY��UUOa � *j 7UOa �j OjvE` �Oa � �*a �a �/ �a �a �a �l �Okj �O_ �j �Ojj �Oa �a �a �kvl �Ojj �O �k_ �j Zkh a �j �Ojj �Oa �a �a �l �Ojj �O*a �k/ 8*a �k/ .*a �k/a �,E` �O*a �k/ *a �k/ *a �,E` �UUUUO_ �j �Ojj �O_ �_ �lv_ �6F[OY�|UUOa �j Oa 6 �*j 7O*a `a 8l �O*a 8k/ �*a ?k/ �*a Ck/ �*a Xk/ a �*a Fk/a G,FOa �*a Fl/a G,FUO �k_ �j Zkh _ �a /�/a /k/E` �O_ �a /�/a /l/E` �O*a X�k/j � *a `a Xl �Y hO*a X�k/ _ �*a Fk/a G,FO_ �*a Fl/a G,FUOP[OY��UUUUOPY�a ��*j 7O*a �-a S[a 9,\Za �@1j �O*a �k/a S[a �-a �,\Za �@1E` �O_ �k Kk_ 5j Zkh *a �_ 5a /�/E/j � "*a `a �a �a 9_ 5a /�/la M �Y h[OY��O_ [a `a /l Zkh *a ��a /k/E/j � Y*a ��a /k/E/ F*a `a �a �a ��a /k/a ��a /m/a �a /l/a�a /l/_ �a � afa �a M �UY ��a /k/a f-E` �O_ �a /k/a%_ �a /l/%a%_ �a /m/%a%_ �a /a M/%E` �O*a �_ �/ F*a `a �a �a ��a /k/a ��a /m/a �a /l/a�a /l/_ �a � afa �a M �U[OY��OPUUOa �jvE`O*a �k/a S[a �-a �,\Za@1E` �O_ � a*a �-a S[a 9,\Za@1E`	O E_	[a `a /l Zkh �a �-E` �O  _ �[a `a /l Zkh �_6F[OY��[OY��UO�_[a `a /l Zkh �a �,EO�a ,E`
O_
a,a 1&E`
O_
a  ;aaelaaaa_ +a Maaakaaa ��a,FY_
a  ;aaelaaaa_ +a Maaakaaa ��a,FY �_
a  ;aaelaaaa_ +a Maaakaaa ��a,FY �_
a  ;aa elaaaa_ +a Maa!akaaa ��a,FY F_
a"  ;aa#elaaaa_ +a Maa$akaaa ��a,FY hOP[OY��UOa%j  ascr  ��ޭ