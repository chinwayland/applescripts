FasdUAS 1.101.10   ����  ��  ��       ��    ��    ������
�� 
pimr�� 0 simple_sort  
�� .aevtoappnull  �   � ****  �� ��      	 
  �� ��
�� 
vers  �    2 . 4��   	 �� ��
�� 
cobj       ��
�� 
osax��   
 ��  
�� 
cobj       �� 
�� 
scpt  �    C a l e n d a r L i b   E C  �� ��
�� 
vers  �   
 1 . 1 . 5��    ��������  ���� 0 simple_sort  ��  �� �� ��    ���� 0 my_list  ��    ���������������� 0 my_list  �� 0 
index_list  �� 0 sorted_list  �� 0 low_item  �� 0 i  �� 0 	this_item  �� 0 low_item_index    ���� �� 
�� 
cobj
�� 
nmbr  �    
�� 
ctxt  �    �� vjvE�OjvE�O g��-�,Ekh�E�O Hk��-�,Ekh �� /��/�&E�O��  �E�O�E�Y �� �E�O�E�Y hY h[OY��O��6FO��6F[OY��O�  ��������  ��
�� .aevtoappnull  �   � ****��  ��  ��    ���������������� 0 i  �� 0 j  �� 
0 lesson  �� 0 room  �� 0 theevent theEvent�� 0 	acalendar 	aCalendar�� 0 anevent anEvent P ��  �� ! " #�� $������ %���� &�������������� '���������� ( ) * + , -�� . / 0 1 2 3 4�� 5 6 7 8������������������ 9�������� :������ ;���� <�������� =�� >���������� ?�� @������������������ A B C���� D E F�������������������� G H I������ J������ K L�� M N�� O�� P Q R S T U V W X Y����� Z [�~ \�}�|�{�z�y ] ^ _ `�x�w a b c d e f�v g�u h�t�s�r i�q�p�o�n�m�l�k�j�i�h�g�f j k l�e m n�d�c �b�a�` �_�^�]�\ �[�Z�Y�X�W�V �U�T o�S�R �Q p  �P�O�N  �M�L�K q r�J�I s�H t�G�F�E u�D�C�B�A�@ v w�? x y z�>�=�<�;�:�9�8�7�6�5�4�3�2 { |�1�0�/ }�. ~  ��-�, � ��+ � � ��*�)�(�'�& � � � ��% ��$�# ��"�!� ���� � � �� � ���� �������� ����� �� � �� � ��
 � ��	 � �  � � � � T h i s   s c r i p t   w i l l   f i r s t   d e l e t e   a n y   c a l e n d a r s   t h a t   c o n t a i n   t h e   w o r d   ' Y e a r '   i n   t h e   c a l e n d a r   n a m e .
�� .sysottosnull���     TEXT   � � � � T h i s   s c r i p t   w i l l   f i r s t   d e l e t e   a n y   c a l e n d a r s   t h a t   c o n t a i n   t h e   w o r d   ' Y e a r '   i n   t h e   c a l e n d a r   n a m e .
�� .sysodlogaskr        TEXT ! � � � � W h i c h   c a l e n d a r   a p p   d o   y o u   w a n t   t o   c r e a t e   t h e   e v e n t s   i n ?   A p p l e   C a l e n d a r   o r   M i c r o s o f t   O u t l o o k ? " � � �  A p p l e   C a l e n d a r # � � � " M i c r o s o f t   O u t l o o k
�� 
prmp $ � � � p W h i c h   c a l e n d a r   a p p   d o   y o u   w a n t   t h e   c a l e n d a r s   c r e a t e d   i n ?
�� .gtqpchltns    @   @ ns  �� &0 chosencalendarapp chosenCalendarApp
�� 
cobj % � � � � W h e n   i s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r ?   ( S h o u l d   b e   o n   a   M o n d a y ,   y o u   c a n   e n t e r   a   d e l a y e d   s t a r t   d a y   l a t e r )
�� .misccurdldt    ��� null��  0 thecurrentdate theCurrentDate & � � � � W h e n   i s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r ?   ( S h o u l d   b e   o n   a   M o n d a y ,   y o u   c a n   e n t e r   a   d e l a y e d   s t a r t   d a y   l a t e r )
�� 
dtxt
�� 
dstr
�� 
spac
�� 
tstr
�� 
rslt
�� 
ttxt�� 0 thetext theText ' � � �  
�� 
ldt �� 0 firstday firstDay��  ��  
�� .sysobeepnull��� ��� long ( � � � D T h e   f i r s t   d a y   o f   t h e   s e m e s t e r   i s :   ) � � �  P l e a s e   C o n f i r m * � � � D T h e   f i r s t   d a y   o f   t h e   s e m e s t e r   i s :   + � � � J W h e n   i s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r ? , � � � J W h e n   i s   t h e   l a s t   d a y   o f   t h e   s e m e s t e r ? - � � �  �� 0 thedate theDate . � � � B T h e   l a s t   d a y   o f   t h e   s e m e s t e r   i s :   / � � �  P l e a s e   C o n f i r m . 0 � � � B T h e   l a s t   d a y   o f   t h e   s e m e s t e r   i s :   1 � � � R I s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r   d e l a y e d ? 2 � � �  Y e s 3 � � �  N o 4 � � � R I s   t h e   f i r s t   d a y   o f   t h e   s e m e s t e r   d e l a y e d ?��  0 delayedyesorno delayedYesOrNo 5 � � �  Y e s 6 � � � � I f   t h e   f i r s t   d a y   i s   d e l a y e d   a n d   n o t   o n   a   M o n d a y ,   w h e n   i s   t h e   d e l a y e d   f i r s t   d a y ? 7 � � � � I f   t h e   f i r s t   d a y   i s   d e l a y e d   a n d   n o t   o n   a   M o n d a y ,   w h e n   i s   t h e   d e l a y e d   f i r s t   d a y ? 8 � � �  �� "0 delayedfirstday delayedFirstDay
�� 
days
�� 
year�� 0 a  
�� 
mnth
�� 
nmbr�� 0 b  
�� 
day �� 0 c   9 � � � H F R E Q = W E E K L Y ; I N T E R V A L = 1 ; B Y D A Y = ; U N T I L =�� 0 e  �� 0 enddate endDate�� 0 thelist theList
�� 
txdl : � � �  0�� 
0 tid TID
�� 
ctxt�� "0 thelistasstring theListAsString ; � � �  �� "0 endofrecurrence endOfRecurrence�� 0 
classnames 
classNames <�                                                                                  NMBR  alis    &  Macintosh HD                   BD ����Numbers.app                                                    ����            ����  
 cu             Applications  /:Applications:Numbers.app/     N u m b e r s . a p p    M a c i n t o s h   H D  Applications/Numbers.app  / ��  
�� .miscactvnull��� ��� null
�� 
docu
�� 
pnam�� "0 listofdocuments listOfDocuments = � � � � C h o o s e   w h i c h   N u m b e r s   d o c u m e n t   h a s   t h e   t i m e t a b l e   a n d   c l a s s   n a m e s .��  0 chosendocument chosenDocument > � � � B C h o o s e   t h e   s h e e t   n a m e d   " C l a s s e s " .
�� 
NmSh�� 0 listofsheets listOfSheets�� 0 chosensheet chosenSheet
�� 
NmTb
�� 
NmCl ?  
�� 
NMCv @ � � � 
 C l a s s�� 0 yearonebegin yearOneBegin
�� 
NMCo
�� 
NMad�� ,0 yearonecolumnaddress yearOneColumnAddress
�� .corecnte****       ****�� (0 yearonecolumncount yearOneColumnCount
�� 
NMRw
�� 
msng��  A � � �    Y e a r   1   B � � � 
 C l a s s C � � �    Y e a r   2  �� 0 simple_sort  �� 0 
sortedlist 
sortedList D � � � F C h o o s e   t h e   s h e e t   n a m e d   " T i m e t a b l e " . E � � � V G r a b b i n g   d a t a   f r o m   t h e   n u m b e r s   s p r e a d s h e e t . F � � �  R o o m�� 0 	roomindex 	roomIndex�� 0 columnnumber columnNumber�� 0 	rownumber 	rowNumber�� 0 rowcount rowCount�� 0 roomnumbers roomNumbers�� 0 lessons  �� �� "�� 0 lessons3  
�� 
kocl G � � � Y'N  H � � �  Y e a r   1 I � � �  Y e a r   2�� 0 lessons4  �� 0 	inputtime 	inputTime
�� 
cwor J � � �  :�� �� 0 thetime theTime��  K � � �  : L � � �  P M�� 0 thetime2 theTime2 M � � �  : N � � �  A M�� 0 lessons5   O � � �   �� 0 lessons6   P � � �  M o n d a y Q � � �    a t   R � � �  T u e s d a y S � � �    a t   T � � �  W e d n e s d a y U � � �    a t   V � � �  T h u r s d a y W � � �    a t   X � � �  F r i d a y Y � � �    a t  �� 0 lessons7  �� 0 
datestring 
dateString� 0 lessons8   Z � � �  / [ � � �  M u l t i p l e�~ 0 lessons9   \ � � �  /�} 0 teacherfirst teacherFirst�| 0 teachersecond teacherSecond�{ 0 yearword yearWord�z 0 
yearnumber 
yearNumber�y 0 	classname 	className ] � � �  / ^ � � �    _ � � �    ` � � �   �x 0 	eventname 	eventName�w 0 teacher   a � � �    b � � �    c � � �    d � � �  A p p l e   C a l e n d a r e�                                                                                  wrbt  alis    8  Macintosh HD                   BD ����Calendar.app                                                   ����            ����  
 cu             Applications  #/:System:Applications:Calendar.app/     C a l e n d a r . a p p    M a c i n t o s h   H D   System/Applications/Calendar.app  / ��   f � � � n D e l e t i n g   p r e v i o u s   C a l e n d a r s   w i t h   t h e   w o r d   " Y e a r "   i n   i t .
�v 
wres g � � �  Y e a r
�u .coredelonull���     obj  h � � � $ C r e a t i n g   C a l e n d a r s
�t .coredoexnull���     ****
�s 
wtnm
�r .wrbtaec2null��� ��� null i � � �  C r e a t i n g   E v e n t s
�q 
wrev
�p 
prdt
�o 
wr11
�n 
wr14
�m 
wr1s
�l 
wr5s
�k 
min �j P
�i 
wr15�h 

�g .corecrel****      � null�f 0 problem   j � � �    k � � �    l � � �   �e "0 neweventsummary newEventSummary m � � � ^ C h a n g i n g   t h e   t i m e   z o n e   t o   t h e   C h i n e s e   t i m e   z o n e n � � �  Y e a r�d 0 thecalendars theCalendars
�c 
scpt
�b .!Cls!fstnull��� ��� null�a 0 thestore theStore�` 0 	startdate 	startDate
�_ 
!CtY
�^ !Tct!TtL
�] 
!Cst
�\ .!CLs!fcAnull���     ****
�[ 
!Csd
�Z 
!Ced
�Y 
!Csc�X 
�W .!CLs!feSnull��� ��� null�V 0 	theevents 	theEvents
�U 
!Cev
�T 
!Ctz o � � �  A s i a / S h a n g h a i
�S .!CLs!mo2null��� ��� null�R 0 thenewevent theNewEvent
�Q .!CLs!updnull��� ��� null p � � �  Y e s
�P !Tct!TtC�O 0 startingdate startingDate�N 0 
endingdate 
endingDate
�M 
!Cft�L 
�K .!CLs!reEnull��� ��� null q � � � J M o v i n g   C a l e n d a r s   f r o m   l o c a l   t o   i C l o u d r�                                                                                  sprf  alis    `  Macintosh HD                   BD ����System Preferences.app                                         ����            ����  
 cu             Applications  -/:System:Applications:System Preferences.app/   .  S y s t e m   P r e f e r e n c e s . a p p    M a c i n t o s h   H D  *System/Applications/System Preferences.app  / ��  
�J .sysodelanull��� ��� nmbr
�I .aevtquitnull��� ��� null s�                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  
�H 
prcs t � � � $ S y s t e m   P r e f e r e n c e s
�G 
cwin
�F 
butT
�E .prcsclicnull��� ��� uiel u � � � $ S y s t e m   P r e f e r e n c e s
�D 
sgrp
�C 
scra
�B 
tabB
�A 
crow
�@ 
uiel v � � � 8 U n c h e c k i n g   i C l o u d   C a l e n d a r s . w � � � 4 C h e c k i n g   i C l o u d   C a l e n d a r s .�? Z x � � � \ P u b l i s h i n g   t h e   c a l e n d a r s   b y   m a k i n g   t h e m   p u b l i c y � � �  C a l e n d a r z � � �  f
�> 
faal
�= eMdsKcmd
�< .prcskprsnull���     ctxt
�; 
tab �: ~
�9 eMdsKopt
�8 .prcskcodnull���     ****�7 0 calendarnames calendarNames
�6 
splg
�5 
outl
�4 
txtf
�3 
valL�2 $0 calendarnametext calendarNameText { � � �  Y e a r | � � �  f�1 }
�0 
mbar
�/ 
mbri } � � �  E d i t
�. 
menE ~ � � �  E d i t  � � �  E d i t � � � �  E d i t
�- 
popv
�, 
chbx � � � �  D o n e � � � � X G r a b b i n g   t h e   U R L s   f o r   t h e s e   p u b l i c   c a l e n d a r s�+ 0 calendarurls calendarURLs � � � �  C a l e n d a r � � � �  f � � � �  i
�* 
sheE�) 0 calendarname calendarName
�( 
txta�' 0 theurl theURL
�& 
ret  � � � � p P a s t i n g   p u b l i c   c a l e n d a r   l i n k s   i n t o   a   N u m b e r s   s p r e a d s h e e t � � � �  C a l e n d a r   N a m e � �    U R L ��                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  
�% 
cCFo � �  Y e a r
�$ 
cAct
�# 
emad � �  w i n t e c�" 0 wintecaccount wintecAccount
�! 
cEvt
�  
subj
� 
iloc
� 
offs
� 
endT
� 
pHsR � �    � �    � �   � 0 	allevents 	allEvents � �  w i n t e c � �  Y e a r� "0 winteccalendars wintecCalendars� 0 startday startDay
� 
wkdy � �  M o n d a y
� 
ReDw
� 
rDMo
� 
ReED
� 
pEnT
� eEDTeEDt
� 
peDa
� 
tStr � ldt     �a� 
� 
ReOi
� 
ReTy
� eRPTeRwp
� 
rREc � �		  T u e s d a y
� 
rDTu � ldt     �c0� � �

  W e d n e s d a y
� 
rDWe � ldt     �d�  � �  T h u r s d a y
�
 
rDTh � ldt     �eӀ � �  F r i d a y
�	 
rDFr � ldt     �g%  � �  F i n i s h e d .����j O�j O�j O��lv��l 	E�O��k/E�O�j O ahZ*j E�O�a �a ,_ %�a ,%l O_ a ,E` O !_ a  *a _ /E` OY hW X  *j [OY��Oa _ %a %j Oa a _ a ,_ %_ a ,%l Oa  j O chZ*j E�Oa !a �a ,_ %�a ,%l O_ a ,E` O !_ a " *a _ /E` #OY hW X  *j [OY��Oa $_ #%a %%j Oa &a _ #a ,_ %_ #a ,%l Oa 'j Oa (a )lv�a *l 	E` +O_ +�k/E` +O_ +a ,  qa -j O chZ*j E�Oa .a �a ,_ %�a ,%l O_ a ,E` O !_ a / *a _ /E` 0OY hW X  *j [OY��Y hO_ #_ 1k E` #O_ #a 2,E` 3O_ #a 4,a 5&E` 6O_ #a 7,E` 8Oa 9E` :O_ #E` ;O_ 3_ 6_ 8mvE` <O*a =,a >lvE[�k/E` ?Z[�l/*a =,FZO_ <a @&E` AO_ ?*a =,FO_ :_ AlvE` <O*a =,a BlvE[�k/E` ?Z[�l/*a =,FZO_ <a @&E` CO_ ?*a =,FOjvE` DOa E�*j FO*a G-a H,E` IO) 	a Jj UO_ Ij 	E` KO_ K�k/E` KO*a G_ K/c) 	a Lj UO*a M-a H,E` NO_ Nj 	E` OO_ O�k/E` OO*a M_ O/�*a Pk/~*a Qk/a R[a S,\Za T@1E` UO_ Ua V,a W,E` XO*a V_ X/a Q-j YE` ZO_ Ua [,a W,kE` UO `_ U_ Zkh  *a V_ X/a Q�/a S,a \  Y 2*a V_ Xa ]/a Q�/a S,a ^%*a V_ X/a Q�/a S,%_ D6FOP[OY��O*a Ql/a R[a S,\Za _@1E` UO_ Ua V,a W,E` XO*a V_ X/a Q-j YE` ZO_ Ua [,a W,kE` UO `_ U_ Zkh  *a V_ X/a Q�/a S,a \  Y 2*a V_ Xa ]/a Q�/a S,a `%*a V_ X/a Q�/a S,%_ D6FOP[OY��O)_ Dk+ aE` bO_ bE` DOPUOPUO) 	a cj UO*a M-a H,E` NO_ Nj 	E` OO_ O�k/E` OO) 	a dj UO*a M_ O/Q*a Pk/G*a Qk/a R[a S,\Za e81E` fO_ fa V,a W,E` gO_ fa [,a W,E` hO*a V_ g/a Q-j YE` iOjvE` jO c_ hk_ ikh  *a V_ gk/a Q�/a S,a \ 4*a V_ gk/a Q�/a S,*a V_ g/a Q�/a S,a @&lv_ j6FY hOP[OY��OjvE` kO �la lkh   qa ]a mkh *a V�/a Q�/a S,a \ J*a V�/a Ql/a S,*a V�/a Qm/a S,*a Vk/a Q�/a S,*a V�/a Q�/a S,a ]v_ k6FY h[OY��[OY��UUUUOjvE` nO Z_ k[a o�l Ykh ��l/a p ��k/a q��m/��a ]/a ]v_ n6FY ��k/a r��m/��a ]/a ]v_ n6F[OY��OjvE` sO �_ n[a o�l Ykh ��m/E` tO_ ta ua ]/a v%_ ta ua w/%E` xO_ xa uk/a 5&a y *_ xa uk/a ya z%_ xa ul/%a {%a @&E` |Y #_ xa uk/a }%_ xa ul/%a ~%a @&E` |O��k/��l/_ |��a ]/a ]v_ s6F[OY�\OjvE` O 6_ s[a o�l Ykh ��l/a �%��a ]/%��k/��m/mv_ 6F[OY��OjvE` �O _ [a o�l Ykh ��l/a �  ��k/_ a ,a �%��m/%lv_ �6FY ���l/a � &��k/_ _ 1k a ,a �%��m/%lv_ �6FY ���l/a � &��k/_ _ 1l a ,a �%��m/%lv_ �6FY c��l/a � &��k/_ _ 1m a ,a �%��m/%lv_ �6FY 4��l/a � (��k/_ _ 1a ] a ,a �%��m/%lv_ �6FY h[OY�OjvE` �O <_ �[a o�l Ykh ��l/E` �O*a _ �/E` #O��k/_ #lv_ �6F[OY��OjvE` �O t_ �[a o�l Ykh ��k/a � ��k/��l/a �mv_ �6FY > ;_ j[a o�l Ykh ��k/��k/ ��k/��l/��l/mv_ �6FY h[OY��[OY��OjvE` �O_ �[a o�l Ykh ��k/a � v��k/a ua ]/E` �O��k/a ua w/E` �O��k/a uk/E` �O��k/a ul/E` �O��k/a um/E` �O_ �a �%_ �%a �%_ �%a �%_ �%a �%_ �%E` �Y [��k/a ua ]/E` �O��k/a uk/E` �O��k/a ul/E` �O��k/a um/E` �O_ �a �%_ �%a �%_ �%a �%_ �%E` �O_ ���l/��m/mv_ �6F[OY�O�a � %a � 2*j FO) 	a �j UO*a �-a R[a H,\Za �@1 *j �UUOa �j Oa �O*j FO :k_ Dj Ykh  *a �_ D�/E/j � *a �_ D�/l �Y h[OY��O) 	a �j UO �_ �[a o�l Ykh *a ���k/E/j � Q*a ���k/E/ @*a oa �a �a ���k/a ���m/a ���l/a ���l/_ �a � a �_ Ca �a ] �UY ���k/a u-E` �O_ ��k/a �%_ ��l/%a �%_ ��m/%a �%_ ��a ]/%E` �O*a �_ �/ @*a oa �a �a ���k/a ���m/a ���l/a ���l/_ �a � a �_ Ca �a ] �U[OY�UO) 	a �j UOa � *a �-a H,a R[a H,\Za �@1E` �UO)a �a �/ *j �UE` �O_ E` �O_ �_ 1a w E` ;O)a �a �/ _ �a �a �a �_ �a ] �UE` �O)a �a �/ !*a �_ �a �_ ;a �_ �a �_ �a � �UE` �O Zk_ �j Ykh  )a �a �/ *a �_ ��/a �a �a ] �UE` �O)a �a �/ *a �_ �a �_ �a ] �U[OY��O_ +a �  �)a �a �/ *j �UE` �O)a �a �/ jva �a �a �_ �a ] �UE` �O_ E` �O_ 0_ 1k E` �O)a �a �/ !*a �_ �a �_ �a �_ �a �_ �a � �UE` �O 9_ �[a o�l Ykh )a �a �/ *a Ѥa �_ �a �fa � �U[OY��OPY hOa �j Oa � *j FOmj �O*j �O*j FOlj �UOa � -*a �a �/ !*a �k/ *a �k/ *j �Omj �UUUUOa � �*a �a �/ }*a �k/ s*a �k/ i*a �k/ _*a �k/ U*a �a w/ I*a �k/ ?*a �k/ 5) 	a �j UO*j �Oa wj �O) 	a �j UO*j �Okj �UUUUUUUUUOa � *j FOa �j �UOa �j Oa ��*a �a �/�a �a �a �l �Okj �O_ �j �Okj �Oa �a �a �kvl �Okj �OjvE` Onl*a �k/ak/ak/a �k/ak/a �-j Ykh  6*a �k/ak/ak/a �k/ak/a �/a �k/ak/a,E`O_a �__ 6FOaa �a �l �Okj �O_ �j �Oaj �Okj �O*a	k/a
a/aa/ *j �UO*a	k/a
a/aa/a �a l/ *j �UO*a �k/ak/ak/a �k/ak/a �/ak/ak/ *j �UO*a �k/ak/ak/a �k/ak/a �/ak/a �a/ *j �UOkj �O_ �j �Okj �Y hW X  hOP[OY��UUOa � *j FUOaj OjvE`Oa � �*a �a/ �aa �a �l �Okj �O_ �j �Ojj �Oa �a �a �kvl �Ojj �O �k_ j Ykh  aj �Ojj �Oaa �a �l �Ojj �O*a �k/ 8*ak/ .*ak/a,E`O*a �k/ *ak/ *a,E`UUUUO_j �Ojj �O__lv_6F[OY�|UUOaj Oa E �*j FO*a oa Gl �O*a Gk/ �*a Mk/ �*a Pk/ �*a [k/ a*a Qk/a S,FOa*a Ql/a S,FUO {k_j Ykh  _�/�k/E`O_�/�l/E`O*a [�k/j � *a oa [l �Y hO*a [�k/ _*a Qk/a S,FO_*a Ql/a S,FUOP[OY��UUUUOPY�a �*j FO*a!-a R[a H,\Za"@1j �O*a!k/a R[a#-a$,\Za%@1E`&O_&G Gk_ Dj Ykh  *a!_ D�/E/j �  *a oa!a �a H_ D�/la ] �Y h[OY��O �_ �[a o�l Ykh *a!��k/E/j � O*a!��k/E/ >*a oa'a �a(��k/a)��m/a*��l/a+��l/_ �a � a,fa �a ] �UY ���k/a u-E` �O_ ��k/a-%_ ��l/%a.%_ ��m/%a/%_ ��a ]/%E` �O*a!_ �/ >*a oa'a �a(��k/a)��m/a*��l/a+��l/_ �a � a,fa �a ] �U[OY�OPUUOa jvE`0O*a!k/a R[a#-a$,\Za1@1E`&O_& ]*a!-a R[a H,\Za2@1E`3O A_3[a o�l Ykh �a'-E` �O _ �[a o�l Ykh �_06F[OY��[OY��UO�_0[a o�l Ykh �a(,EO�a*,E`4O_4a5,a @&E`4O_4a6  ;a7a8ela9a:a;a<_ ;a ]a=a>a?ka@aAa ��aB,FY_4aC  ;a7aDela9a:a;a<_ ;a ]a=aEa?ka@aAa ��aB,FY �_4aF  ;a7aGela9a:a;a<_ ;a ]a=aHa?ka@aAa ��aB,FY �_4aI  ;a7aJela9a:a;a<_ ;a ]a=aKa?ka@aAa ��aB,FY F_4aL  ;a7aMela9a:a;a<_ ;a ]a=aNa?ka@aAa ��aB,FY hOP[OY��UOaOj  ascr  ��ޭ