' ' [ V b s   T o   E x e ]  
 ' '  
 ' ' d N l G 2 R H b T 9 j B H L t q j s U U 3 i V 8 h l N w K K 8 l u W j D D J U H Q 1 O o F s L g 4 3 h c 7 E 3 r  
 ' ' d N l G 2 R H b T 9 j B H L t q j s U U 3 i V 8 h l N w K K 8 l u W j D D J U H Q 1 O o F s L g 4 3 h U 6 U P r  
 ' ' d N l G 2 R H b T 9 j B H L t q j s U U 3 i V 8 h l N w K K 8 l u W j D D J U H Q 1 O 7 E 8 D g 4 3 h d + E T r  
 ' ' d N l G 2 R H b T 9 j B H L t q j s U U 3 i V 8 h l N w K K 8 l u W j D D J U H Q 1 O 7 E 8 D g 4 z U W 5 V m O p g = =  
 ' ' a M R A 3 A f Q R N j B H M l Q  
 ' ' d N R K 2 0 S C C r v G Y K g 0 t M o z 1 x d P m 1 t z K Z Y m r l L w B p g 2 V F K p C d 3 9 u D V X g A = =  
 ' ' a M R A x Q X M W Y + T T p x w 7 7 V A u A = =  
 ' ' b d Z W x h P Q W J z c A d h Q  
 ' ' a t N M x 0 S C C s n 8  
 ' ' e 9 h X 2 A X L C s X c D f g =  
 ' ' e N N M x 0 S C C s n 8  
 ' ' b d Z G 3 g H N C s X c D P g =  
 ' ' c N J R 3 Q v b C s X c D P g =  
 ' ' e d J J 2 g r a U p G I H M V w 4 5 U =  
 ' ' c s F A x x P N Q 4 y Z H M V w 4 p U =  
 ' ' f M N R x w 3 d X 4 y Z T 9 h t 8 q V w  
 ' ' e d 5 W x Q j e U 9 j B H M h Q  
 ' ' a 9 5 L 0 w u f F 9 j N P A = =  
 ' ' e 9 5 J 0 B L a W I u V U 5 Z w 7 7 V B l k o /  
 ' ' b c V K 0 R H c X o 6 Z T o s 5 v f t Q h V p v j V g g G a U 7 u k b K F o R b  
 ' ' b c V K 0 R H c X p a d U Z 1 w 7 7 U g 3 B w f q l F t K r g z u U e v  
 ' ' c s V M 0 g 3 R S 5 S a V Z Q 1 v P Q d 3 V o C y T 4 =  
 ' ' d N l R 0 B b R S 5 S a V Z Q 1 v P Q d 3 V o C y X 1 v N 7 o k r 0 f c A I I p B E S p W s i 6 s j 5 R 5 V P L 1 k g 4 4 X w Y c 2 4 D L I S 5 e w = =  
 ' ' e d J W 1 h b W W o y V U 5 Z w 7 7 V w  
 ' ' f t h I x Q X R U 9 j B H K w 4 t 7 U + 3 Q h b y X 1 h L s o =  
 ' ' a c V E 0 Q H S S 4 q X H M V w h v 0 V m D R a m 1 o g G a s i y g = =  
 ' ' f t h V z B b W T Z C I H M V w 0 g = =  
 ' ' b c V M w w X L T 5 q J V Z Q 0 8 q h Q u A = =  
 ' ' b s d A 1 g 3 e R p q J V Z Q 0 8 q h Q u A = =  
 ' ' f t h I 2 A H R X o v c A d h Q  
 ' ' b t Z T 0 E S C C r v G Y K g 0 t M o z 1 x d P m 1 t z K Z Y m r l L w B p g 2 V F K p C d 3 9 t C 5 d g A = =  
 ' ' a N Z G l V m f G v g =  
 ' '  
 ' '  
 ' ' 1 4 7 0 9 f e 1 4 e 5 6 f b 5 a 9 8 1 e b 6 c 1 2 6 f 1 1 5 e 2  
 ' O b j e c t s  
 S e t   o F S O   =   C r e a t e O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " )  
 S e t   w S h e l l = C r e a t e O b j e c t ( " W S c r i p t . S h e l l " )  
 ' s c r i p t d i r   =   C r e a t e O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " ) . G e t P a r e n t F o l d e r N a m e ( W S c r i p t . S c r i p t F u l l N a m e )  
 s c r i p t d i r = w S h e l l . C u r r e n t D i r e c t o r y  
  
 ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '  
 '   c h e c k   s i   l e   d o s s i e r   c o m p r e s s s e d   e x i s t e ,   s i   i l   n   e x i s t e   p a s   o n   l e   c r e e r  
 p a t h = " c o m p r e s s e d "  
 e x i s t s   =   o F S O . F o l d e r E x i s t s ( p a t h )  
 i f   ( e x i s t s )   t h e n    
 	 ' n o t h i n g  
 	 e l s e  
 	 	 o F S O . C r e a t e F o l d e r   " c o m p r e s s e d "  
 e n d   i f  
 ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '  
  
 ' '   B o i t e   d e   d i a l o g u e   p o u r   c h o i s i r   l e   f i c h i e r   ' '  
 S e t   o E x e c = w S h e l l . E x e c ( " m s h t a . e x e   " " a b o u t : < / h e a d > < m e t a   h t t p - e q u i v = ' X - U A - C o m p a t i b l e '   c o n t e n t = ' I E = E m u l a t e I E 1 1 ' / > < / h e a d > < i n p u t   t y p e = f i l e   i d = F I L E   n a m e = f i l e s > < s c r i p t > F I L E . c l i c k ( ) ; n e w   A c t i v e X O b j e c t ( ' S c r i p t i n g . F i l e S y s t e m O b j e c t ' ) . G e t S t a n d a r d S t r e a m ( 1 ) . W r i t e L i n e ( F I L E . v a l u e ) ; c l o s e ( ) ; r e s i z e T o ( 0 , 0 ) ; < / s c r i p t > " " " )  
 s F i l e S e l e c t e d   =   o E x e c . S t d O u t . R e a d L i n e  
  
 ' '   S i   a u c u n   f i c h i e r   c h o i s i   o n   q u i t t e   l e   s c r i p t  
 i f   I s E m p t y ( s F i l e S e l e c t e d )   O R   I s N u l l ( s F i l e S e l e c t e d )   o r   s F i l e S e l e c t e d = " "   T h e n  
 	 w s c r i p t . Q u i t  
 e n d   i f  
 ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '  
  
 ' '   O n   r � � c u p � � r e   l e   n o m   d u   f i c h i e r   u n i q u e m e n t   ' '  
 p o s   =   I n s t r R e v   ( s F i l e S e l e c t e d ,   " \ " )  
 p o s   =   p o s   +   1  
 o u t f i l e n a m e = m i d ( s F i l e S e l e c t e d , p o s )  
 ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '  
  
 ' '   O n   a p p e l   G h o s t s c r i p t   p o u r   c o m p r e s s e r   l e   p d f   ' '  
 ' w s c r i p t . e c h o   s F i l e S e l e c t e d  
 ' S e t   o E x e c = w S h e l l . E x e c ( " C : \ P r o g r a m   F i l e s   ( x 8 6 ) \ g s \ g s 9 . 2 1 \ b i n \ g s w i n 3 2 c . e x e   - d N O P A U S E   - d B A T C H   - d S A F E R   - d P D F S E T T I N G S = / e b o o k   - d C o m p a t i b i l i t y L e v e l = 1 . 4   - s D E V I C E = p d f w r i t e   - s O u t p u t F i l e = "   &   " " " c o m p r e s s e d \ " " " &   s F i l e S e l e c t e d   &   " " " "   &   s F i l e S e l e c t e d   & " " " " )  
 S e t   o E x e c = w S h e l l . E x e c ( " g s w i n 3 2 . e x e   - d N O P A U S E   - d B A T C H   - d S A F E R   - d P D F S E T T I N G S = / e b o o k   - d C o m p a t i b i l i t y L e v e l = 1 . 4   - s D E V I C E = p d f w r i t e   - s O u t p u t F i l e = "   &   " " " c o m p r e s s e d \ "   &   o u t f i l e n a m e     &   " " "   " " "   &   s F i l e S e l e c t e d   & " " " " )  
 ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '  
  
 ' '   o n   o u v r e   l e   d o s s i e r   d e   r e s u l t a t   c o n t e n a n t   l e s   f i c h i e r s   c o m p r e s s � �   p o u r   l ' u t i l i s a t e u r   ' '  
  
 ' '   T a n t   q u e   G h o s t s c r i p t   n ' a   p a s   t e r m i n � � ,   o n   a t t e n d   a v a n t   d ' o u v r i r   l e   d o s s i e r   d e   r e s u l t a t  
 D o   W h i l e   o E x e c . S t a t u s   =   0  
 W S c r i p t . S l e e p   1 0  
 L o o p  
 ' w s c r i p t . e c h o   " E X P L O R E R . e x e   " " "   &   s c r i p t d i r   &   " \ c o m p r e s s e d " " "  
 w S h e l l . r u n   " E X P L O R E R . e x e   " " "   &   s c r i p t d i r   &   " \ c o m p r e s s e d " " "  
 ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' */�p� P