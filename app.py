# -*- coding: utf-8 -*-
"""S-OIL Inventory Risk Analyzer | single-file Streamlit app.
Run: streamlit run app.py
Required packages: streamlit, pandas, plotly
Data: S-OIL DART filings and Korea National Oil Corporation Opinet exports.
"""
import base64
import json
import zlib
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="S-OIL Inventory Risk Analyzer", page_icon="🛢️", layout="wide", initial_sidebar_state="expanded")

# The following is compressed embedded source data, NOT a download dependency.
# Raw data: 국제_원유가격20240101_20261012.csv, 국제_석유제품가격20240101-20261012.xls
PRICE_DATA_B85 = 'c-oD<%Z_ZzZY20C!U1|F1)TI3dB=^aK>_^_YTyDApjHW0U892jy~5Xu9Z4E3Gje+NKDjO37wLQ$jsN>EZ94v`$v-vy>u>-2umAgh{lEYGuYdh*+e7~GumAeb|MS29+rR$%-~Yj%e`fo!=l$RQ_22$M+JF4V|M%D5q_xL<{@4HKpa0`u{{MgZ=jZ>ZkJbP2umAS<|MB1d=imSR@AuzW8h_UO>u)ZPChwos9&P?#&;Mu-9sl^(fB9d3|F?g5{^!5`w*0ZifBxea!X6RMhZ;hhkBta@KO_Ns{%aZmKL5f*eCCf?@r{Qm;F}iljeLS@AU=Pnvm>;Tk1!+r^EdJl_Q&Yt2yI4)&!6QhLYxue^G2=S@}FbP5YK(RAwGX|H-tD=z2#v9qIiNRj~x+OMra!muE!RJP}V1eW{9vooE_qraL*nQdPKP3YVOalm1}l?Gyu3a0@PVQ#9lwdeOXih_`LRrP+C4i_j=2n+VWc2RAxpfZ#40e*DeU3um>WPh)^=ZUIEJfc**-dv3!P|eTGghc{?EN=?s%Qd|n<xe11EfKpfY{^94~J2L7-=W`Q{0*7F7N{LO}^*dxUG>6-1~1fo8qM~MA$fcX@?KpfZiixAuKxC!Bk)+`IGxmTBH4B;`}FNEXKJAR376`^@W_;fF5EIK2^oy~dgdWwxWocpVaUVDts-(r6VJH(bDKJPL?1fMuZ{T5}iV?>C1;4y0b6#xAAe2NVa_W1OP=cmu`U2%Oow}`M`UqpKPq+E1)`cfqFhWPxf^(UXB3UA*+j1pgiaM9<Jh)%~7#;3E}Rw0Vfi*`81^et-7PRG!qpSRPEI-Qj*8ctE)aNZEFJ6-G;&u^2C(PI>*BV(MqVhlewT(!ehi++5dF9*{ZNS$zyJAsU@(HcBk2jq$oUO33P^Fj9V1TJ!rZVaO>2qRjI#Weax`+7#dFg`7vT=w&uiLux-5N?JqO!l+$1&18u1n+ph?({`Q!?s=z@^W)pQQGy}f1$LJ7mWA&8BoqEIz*H^d!OgEX!Jh4<&*S!kRl9X&J1SxA_v6quiIwcXI;NdioAzLj4|S|=O2<Ua?#r!@J0H-+uoai43`padpUsAFBhp#oS)2QzTdQbk_}v>mCF`p+l?KhZZw-7F+!f|iWr~q49T+QW$-cn9*U+t9R0Xmf2#&&8#4q+=LnJ6sIMV=ugCS7enGTHiEbBW2+z-kJ0FZ3-6KGV%iQWAvgmd_@C*$f8m&i2PVi~`4ktJwzMU_5O2wxXpFx6fGMWr<l1mn5D}q<NB0|g0VApCYV=;Sn$@#$v+p|4;NwWN1Ljj-D0pm^>J^RI`)#ngP@#|<)Gy&h?e1E<l!-;+Q{gm#HXe!b__-<ovQRGUZVus*g%^N<N*CR+?7r#>Y1hMT1ae(Wdh|$?q%auMyh-L_})pxqb-}jBN6n5n4P$NbNcvw_6Hw`8rMOYsl_w(0v8cHnu8Z{l`%E#Dt{Thc*h%<kVHVZ};H*SV;V)1E3i%EIR6QizrjH@%hGk=b4d(Bs`f)90#Zh$eJqqdB(<}r7=T+gw2_B_k+9Otfg-RM;hh#D=rZ<TD+V^sI;H2lU?%Qn8z#j2kkVj6p+`vi<*MvR(|agl&bl*JJzB1jIryn&tOG0yI~;a&G=)+r%0r5-$_>j9*}2iXcpQ{uu^7)Ws8&+j&3T-o&ut>5ERmJm*St4rs=iSJ<v!hPiVC5-WoZaBw9Uc&XObKGdygQW<ovvb4`(v#6xHkwjbex@-j_#RiqK4R=8<`E}PVWWw)BS?)7QX!-$4TNJggO#F|ffFy$H*?}9tQJ))JA`H(vKYA<RQN&C0>y%NW|dmhBHv>hpX(F*6pJIok<)4t-WvCh3L!;tJtVr-$=u@{kP3$Sch4!zxoEG4IB4m%E99%C%`q!PS9Z?=zHugm>yvNnu{9Cvd4RB&PEm7G(4&>xiT$>bFoNId<gHd=r)W?BFNIs#B#oK0{%ue^35AKC6B+C8#!yORI{OxLsx_~p$gWXHBEu}YW5_dNKHYP>=y{C9T(;Bt_s9^t<1w~hrK<u7|9SP}AqoT&)pw7_rcWT0^0uNu*S4zDS6VL{`mXicA<@!z=oul(V&-BB7>ipJ7LyHTF5cl=>@#9aZgGGa-ru6@=ce1fGFWK3k#CWOWc_k$0ppCZZHG*2yuau8OgE<(N5)vpWi0Xn`uas|`+RiMIVPK)P&K+UR^@4o<f3iO{P!d&D+|D%4k2Imen$5ZWC|tL+sQ<bo#unnGDroZS=YZ@Kz3R}Sq30mw3?i3ty+k>vN)mT#?<auB~2jN*7cRp2z)HKZ;#((&bL-}YA=wJhK@OxS%s6(*?o_13}*^}T+v77glj^v^@Y+04RAMjpwjRVC9Q-><DRaF(jv;4QO*KVu3W0R^6Y|Ei7JGZ0Pc!Bsor(~X+45m`M$+|ec_46bAI{$i{y1iHb#7L7T8XT!*kL#$&2HZY@}C6{nNAkWpc^DdHcVieW#Do@=4Z?hrf9h86$j>t3V|Ab<8^gLPHl6dYJ>#nO&Xv`PG%$jevEJa@qA5SF95Tvr)u2BgSOgUpZ00sI>k#Ai-YOXPYrP1L?Fa1Ieq!*3rlWLi2b2l2C4?Pmbqsd?~(*<cr+z8fSiTAoEIulmWS+ze@=TWupbYg{?wCl%z4dBBdapgvRjU?M0N)PkO6Vg+Doe?G-x2ysMPO_h4t)Sy{YgNBNBD(v4e2**bCOPJ6=(iuc%d-ES^uI!mQJ5ID={r<W|;Gs-;}O4@hrkb{X}C^R-f$(ovXbt$0?{F#xHFmILPb3KBz%?=>3J~iR%d0jZm6!%rSOCVRos2L-%@DZ6g(RSdz=O~eH1!=L>94&m}R~hd+rw|U&s^(-M$L9*ny<Ph3B=r4a(L#e$@T=Sr<;rJSdLdUhF(^ud3oi{D@;KVNu?8s9g_}I0fhYFfG4Rbm(lWm7pM#9?E@Wnqw6Z%kf{e~G=VQC3I#=jRbMKMwdq|0d775ontyO9AmDVD?en399-w$C9eU-6t_fMtWL$&P$BcZQ5aKoSHsL{N03b6=@#;Kod&}zVr_9^-piPkz}HVfKz=-3SC*ciCqKcTtO3OTTmU{1j+6Y^>nc3L6daMeAX<H#Uw5s<9UsSr|5aFc<kknXzCVvsqUq({@vnV7*3etnO;*PuUH_Nn{)He0;8p811}J~Ag_Rw>w5>5#IA4HB^?pARy1Bx4a;0p!#O(g7sFxG3>8a|7lTUMGyoFxJaH-!4K-`9o&#8E|wovqJ$RKwMzhBzu@~4Cf1?W%t;l@0pO|3>X@^J70~BA7S4sMZ1SkM~Z%AZe~GE;BaLUzMk<3AtcqIkJHzH;iH{Ta9pDtSlzYq9X@~jMyDeQ;DX{VDL~h9zySyh6)@*<jWd*1h!KE!e(MP^6!9lCYuE{P^12*0I&#<-TyY}+$xO<fdZ#P4bb^*k_?5NS-mW|VTrdWJtP1b%F_Ab1K(AN@-<62Z^8r>yGfTE^iplqkxIl4;tj8MCp1G)jj&KKjX)ek}efC@lb^XLIBLmD;(sRIMMxiMI*wGEH1hD5`oRI?@)P@C^6yxI;xJaeh3qknl6(c@qTf5W`a54+YX-m6YO2StN3Wu+zYy@+EoBmBmta_!Q7^~(T?#IZqRtOIVoi_4G@GZ>rL!o{G2d>vl|6Fh<0CNwn!K~5gE)rsJcLH!@(zl`y6!3wS6{UoP^1NXO6Pb9!WJIPmWDa`-la$cen$VA#kfJHh#eSJ^P5{ayvT7dX`!NVU!*tZ5h2&1msFnG1!9vb&05?w^&aZxflU?uxqUa$ACV=2#bXF6t#bmz&tM~<Oe1*$2hj+rkJdb=ak#xhzwzWEsXAGdgBbU2^;qCqcsO%N))rOZxfcwKp%mFCu1i?fEPtYd;$WAbKsGl~Huh{Z8SiPXJXLw&RdWr-$49W{TL1rYw;Ts(3iXA$EMh4|4i4(fySQ>IBgIf9h4OSw;3Cay(eTi4xAy7~<gg#lQs&Inv4YEL?u}Ee@C&Lm?dBQYNuLN70@p%g0APgU<;5R50Ahlu$wGsn(K{g^aQm|Bn{sPzyLRHL0#;o(*odL{sngdWGzyT95h_6R-_8)vVXF5*kv;?m~$Tsis;rh-sct_5BoJDvu55Q=-;Bk$%06xv<w>FGL_*XR|Y7Erja~GJ8KaJBK0j@m%qT~y#Gfq!{sYmq$2p6n%X!^{E*KAI&w?FmBXS(2>X~`v<vl%V#55RH#Ch3MT7Ob(f5OUyoM-X2;z?_V<8YvhXhj_iS5#O2^>l2BFN#J?G7cE$TMO%e97#;acF#?4@x1?U74!$CbJR*v1(hw)CLb+S3o?qpPol5V9XBZ!mt`(mw@`!lppTrLk33r9ec6@Ya<ph{7fMkgbHn|Xj#tE=DoZwD343SnJ<dvJ`1P7d8D>rNv;AR>e9ed2_hCeY=&BAcXSCo*|Ed)4~&wsVDdlP};3!Y=Iu_t^`?&u`jq-E6xO+3Z=G|M(QpE7fMPkCAcLWMYH5=Ej(`%WA94(%7BPJ>REmtKTK{_!dqlW9#zHtjk&OaSJ#^79zytSOUI5tE6*%#H|wl{yzmnf!Ct>E_p_FOlSQvws!{IRMic0%=W2VXzN&&jJVydrn3wgOf^ISpPmLS<opNmHp?91cVnvXhJl+X73f^qSx#LGtcO-6>1;eYwoVJ-!!pLQfTJ&O<APKf6?#oQ|t*NFnfzyjVo0M=}>}OWcuA%1%%5Im8<^Bo;DEH82h+?$25nYm=%SPK?XB8?QdbFI*_b3baVEb(iyuYN{7~R2DJ%+lFU#pdW`djX7`2^YCXv4h%<oMUREK?Y9%hYQ*yLoN9Spb8WE0+kj)59DnK>m=Z~=x6WwCvHb86}N)2Ild&3yhu9|$`K<q#=U(ox!`t`qq*kKmz6dTf~2ozM(b=uTgCR}tYV@zyTmgy47qqG)#L<seoMf5BdCx(Dy>}3|L3F(~BP^Yv%L&CgEZ>MySbBz2IEe`0<PRE)OYfZ5JF){@tIz<7eSVaW%$~5a4cANz380HvxTQ-d%_V-Qr4H*^l+k;NFw5z5NA?`xBAR~X1a+C&k@IwD&m*Jb`r2MbsmUB{iv*dn`#CEJQ2VFz#apxn1{!Y=mJE<h=fC{h=Wnj)i0*Ki^${Jc4Jl)XW=`hgxYSFS5XilmI3|PHdXr%#Vtn`X?IgP$Xpm@EBG7Y=n4Q#$@#`S`$rpmFVh4A8&hdEyjW3_T%&#eHw=8!Tka=)YlZ!lD4Ee8##YS4ELvuV-BfoF*Ar}~skR@Yp>ggj-#=%y9x!bDb+)EpKFLf+%i>HBMgbZ9qg-~qk_Z{(s}Eh_m#2)h20%wiazO??U4{lqbg5#0;Q+7<U_{%$EJd=nG`7Lt(4TtJ=%{Q}srDx4xh)U@lQyU1%^@crpR_d;IbK3UBRtA>&;*g;{DFc>yFKvZb%aDbcGhtL}qIKUp=U+O$<gPcqwe=ik<QA(ph@6(79X=6<jhnld)ASSh)nD!%8xb_Jgv-=xJHTCLVI_bU-NJvi5`-W#lzHSn=*}=?rqj?B6-^VZQS8=#7Oq&^K*Lex`59^wBj&sRcflu!P($~9WZLBHQVa&eG2ro6AMwxvyWgU}tx!J@%O4-~w^wJdDWXkN;0?xV~7U8X||6W8~tjF=2uby2;Zi)cD;{9wBX{JLSLRF7I-!f3GV|{HU8u|wykU!c0fKKo^!3Sf)5=8Z^NUO7sSXDZHeMw2^SfHmkh;BJ+DfS{x?*X5HvLo&26nOngM<YC7H!8I6UoB%J_<&qCqB^EQL#LG~YuiLq_Pni7W6gQG52~E+jJuUbVa^z0kN7~DS9?7_VWfbL!Ji>>Mk^_xDja#40PnHqZXj50Bz)<yh4A&dp^UW-xZ&cT%~;R#?h+<|@%<3!X;%j`iJ<!H_4)x{e~|N&&frbdxlI@Bpr%eZ0j%Y&&`=x;NDHNco+(UUq=kw8Wkv^++D&J$bCU<806bkXpl;>(`3a`19XKW10`Rq>B&&+X2QKxRj%WsXL1rf}+(6CFg)bcO&~>uk3}?Mh0W1R4e1e7EXhOJU&mX&C=|Gk}Jtv3}g3TCfe4U`$_oCGX3gzpu>j5@)fU<tl%J)_qBnQeO2M7X=E!KN2Cm`EM$CK$`zQ`o042ZOpL1-y$5*t}XwJHWNsma7bK&CqA3WI<H-{4-kdEdRkDEp<!gds*?F1q$^%%P{P4~$?|2>}-~kd2}-&p$}`VEIaQF+xTl%wQx{phRit>~Dbeer=&VzQ=a9l?1MmFGsBFgdiFU9BV_6WhBTVKG1=DdA7KGJt7V{$f7}N?I0cx0*{D;hcT%KNP_$Y1;De^WJY}8VanK_^|%rse1Y%=W00@J!&DWC<F|1a$Byt)kkyN{_oZ@F(4iVog92~hqJ6;=LRK(T3|%v-;yL7Sy^hFkCQtPu>!tNc3INOayRq#_CxlD*Mr;l{ByqitLhnVC$8F>nZaGw7Y!V0E9!B5|1$==v=&;>MEA@A?)kJ3&28f~g_(rKen^gCFXlw&dFMfS_vlAiZ?M1YIPul6gx=iRg2@e_(TD`bC@b$HLyLqoutL3Xb3E{{jue4Yj8il7*h39wZcEvsKD^*VT0t>|<#gwXHvDagg<6FqN%+{;aIQk;zmo+xlih7ZrQ_Pdf*5j*Fuy%{KM&qyn_bUt^ftM+bYeyYMhXhIMaQj}QiD?2NKMmxujil&DNO%)~xqSLOzm)*)>|J`V)?0-&Ov~`dGyquaL)qgdFq&~JO(Q-7?^h0xO+HwME_H+UmXQK)7#~ay#&BVm#ONTWm;-{2JHA>@+_OUhGB&v9BBl}4w$l4nmhY_yFz1+~ii#a)(h0zhBa#LVx{Nv7jM?fp@7qqC7>5LDkR_Eu1Zd7b4>z3K{ACE`g>Nu?8@1Wf0S>s~5<#~(6%}q6ij<kmXrqew<Pm^y!<Ym2Xu}FsXeR(`5B(5l;D$3}=W99W7rsG8TMTyuoHd^@Kvi&pokUnw5z>0D;-4W?d_>ChQlQ=EAmK+Zol{H&^_5o$NG@1m2)3h-DZDWg!fdM{)(%6u4jEFfwDUu@Vit^p1vx`HWCjvrRi@6X$e5QFzRyouZX=y)R0(1_nf^MJ4_NS-2mDOx_MU6DkX-P<SMF={;Q+)cQF%p-`IuXGDD|l`zy(u0w~3H2aN)+3C?#|Nf)gxcr_E+K&%g`F^gt3fNu6ylCp|%E_&jqf0)*+WqF#5N_r<bw;M{dArIJE)-OuC=J2e_~>kdBHe<Z!)8-sL`GS6h;><Ez!5Ar{$PNWEMrBBeHPi-ViXLYRclai|onc!?u<mU??4Bu#&Z8`X$R6fW7k|n*ZG;1J^@wx)Q7?tf*G>bBP1c*C&g#=p8DybcqY<c|4oK#YcFeOgc6PpcROq~NzpjlHdn0|o<DpoL_)@U7ze}P34J{=&K@FF+7*#SDL5dyxl@5r2lF|Ub=bbg-jK^W7k3nXQ>1sr$<Cd?^HG8f3|Y>JiC8DF`a%Z<{_+(x8EY(kw<K+otzS>z4zL0oicR=0$21HM&LrLoBY3S=T#xU*>d-f=3av`vuJI+J)i0<;XU4sbToEqqONe&3Eep8#nl4F(W4J7}PDJOSozSFd|eHe>79EX|}mA!(~D@OVKf__XXSi#ME`UA?l2P79XRQVsf<PzqdC5b8ww>=W$x$j#^ib1pZls_}MK?6n7&a2{hh-8v-)xLG^1H2l*AkBqN-#aHES%wWxXf8ZNu1YeC}tjuqYAW3%cK`pur&TS(jy0$_8Tl1y_+G-f2-G4ev1MrP#!MWKbRlRR7B1nHHlHW_AoOoc@YSChuW=8|I7zO5LSF3FD%|sMzKKf64wQQ4Km$FQKqS)>UiUYo!$&n2_uSFa+Q%Anj1_O3dZZmSfO}-CO1rj;-{W%+9Z$)*rMA2=MH`p0rVq=V;5g;^^W_>|A6QO9JgCdxPtofT-)=O$S2dOLKgBXSdsWs7bDMK}&p&-zfx~qYej^KhK4?RZFQo2b>GGRAz!}G_rZbdozzF`{X>;CZ*#3)AVC%6V9On{P1M}R`gmBKicj(h}5R}2;D>SLT~3bu#-IAq_5j0ckxG{+(PsC4thCt$;6P;(m#1976JRm#pJ<-n>*{1SEjW;^Tw2rZ==r!M};3FwSKxpI?opb{YzTL3VwQa<>4##d~ZH(d7v;7xo)$Q#o8M?%r5Y+HmkGQ9DdhTP6+`r5FQ>hL|h!Xq(s39;hn<AFD!7D!I7&IubxZmxwv{;rGq&Kt~bcNETXd_r;-LOAbx6H>H+B;#6G?cE@7JCmSdIzhqD?@KQ20Q}fcH^RueInPOk2lMnY5odA7S}1d1J7i0Mrx3$1|5sIMUvoC8*xNCPCXBnm_&Qs@U66eM3INpK1Tp4W&28QOF25io$`=OE3ZW=oa8Bb!twx_HrjsB+VP>U6j8HiVE3RG&g`$p|EaMAOvJlcSoa_m1aDF(|;Dpjf&(Hgh7YzNRUj>1;v`n;*yl(2yz?glop#g|sqK3?Y2i>H~DsRjKS}t>(MtBDo&1(6SGuy|vyB7~Y$l{Dd`J25V6F4yoU39Hvb(SXWTFJ^T20UQZS!&EhTUPa5ceb5Rkho}LkDOzk`9+Hsg?^I4I>%Q_n7ttjPS6W)n*dzw4Vjo|szreP18v6|qebuQK~1!xfmh4`vPukO!Y#HNv?Ow!FyE_>Ay=#z#|+Z-AMA#)*t^pK&d)boH>ke2n(+YjbEZB>zs7E+ZswKu2yo>a-lWvVJB9$Tm1=LjV8Sb3;AEaxk<QNn9(m&$gXaKml#Z)ce4Tb&k+(A=KqLmP`Kw~|&cyr@W}U2@lsY;y;f4no?^hs-sdAA_{^AvdrZfbzB$fv=0~z1)vJN~bMK99M9BguBLy+Y-cTq<R7pY`9Y>YZVbo4Jdi6gwR?*>;OhFO>d3xmh-t@$nM{osRUQk`4rO!wmERwQB|M$wxTiilikdtk4orSk(`;3S_Pi?LF(>-^GiyNk4vChN6!LCRBt2Su1k=m?F)_^DZVV$||{%wHxUf-%+?ulE3i5XMLdQ_EPd9&_RpVJW}u>i?WYse)w$+W{b^af`KTSd|Fg-^@p#?^|bMAqK)Vr_2%I^$)yMO1i7|YyD%5s54TaW2dHt6+s6d#4lyn+et!LCkYLtoFA(qkXGDXu*v5aLsSI^Q%d<JY^=d^1b7_G#0w_v>cIdGi~v2|Fcf;(1H2f(1}2OFmJxi7;lm`e1#}%IGdp9=vN}0*x)gJ`MT)3Zmgu#1BPW^h^s=x7>Mh(3P}pd{!DxRe^6GE=-{~{p@%>rrF#oAZW!#aWsoE(#T8uw!wmF|CsggUxH}U-bu%d5#Ff=vE`%9DuI!spg2a<JWeSMV;E*B=kpBmr|Zvnhn;Vs4%+t3A(k4=X45CX521bsz#JHzWFc<GGGCO-|`h8gdldmt`Ygl_LX>x~bGGD&n=<z2m(r6)zWp)*c7(^_{7uab4ZP$nbW?T=#hb+1yRC8EkleeoW@H(w;%Ik{odCcJ0uQk&3VD!8#5rpgFdqx>V_djkZm@G`Y$n@Kk~v<V0HF&8QFI<XMLjkRaUE#d>Ia3(E5-j4;!fc=5e7j<mKNR#3G9@+oR_hzpf-voR_%A790oBpHvBO@Fe6Z_wwt$5U5Mg+J%DRD1mB6rn3UGPM$b;dX*x^@kv)>}HjPkW!}-2Amv7N|rhY)7PjNTVNh_x2ZI`@)x%%sY{G8f2^LOL;Oe%E3O2G#Q-XYXEOV`&HP-(#Y66(WTtQJ``p57?DGWeK;PUCVU_Zs6wQZnC_YNAC<9CpYcUyW5RI_b@%>Eecu4V0SW<FCm)Z{H$(^vrB|HHovDA`EVtj<CO5zV>oCYzI$bA%;z@?$I>rDR0oYJ_Z^jratCtI@AxyTPw1YB2huKUjuCv*KQpLSlr+OR#S_Y^@zJ@T~dy$st_y}(@*-w_xD(2`b2hpf*=R~=INrwQL%>msa;$n?y6G4sW)E?gGum|FmjLJ7E^%A8@QFkH$jV+96-!0s^0Rka^kRV?fS*%PxB3aD)Z-#!n<=O9}L);@iP@n4LejAmB0jkqyeKAUxJxLs@7w(8}eLg>7HKGOh7M`E7PB5!dLX<z>T$mf+&H%Ul$^}a9(QJh(rIX?Q6KybQ9ehJ{1bijRC-Pm55fQfkDBmhoevb)kd#lp2Gr|M+LPF@q>^F-)_PW=w_uq^QobS|dolX4y1jH#DlO|dOph1Y$I{;+<dC><1wOiMU0HMg?O#h5rufw_`L;G0d1t#wYZrj;k;sZCn@yR|W+Zw1~ANNMdyb_I+VW@-A@`R)V1OuMyS6STcfKv>ER2?V!Y>Qb4=<m%p8TAGTg$^S-(H2vv6bQ0(etd+D=-)P+iCb@~&W#V$gISv1XdMd<5H>#bx_$OVpsDg*3+1&v092F@CQfFvN|DNsbmLCmIAaoq^1vW<#{*OtDlBX}n@mHqe_E`m%tf<0k)yu0MHi@G!HtM_4c5jqh_O1>QQ6EHZ3CnHa(qEDlWkY=>TeuiIXo<*8mVGA8)s<fU(ySrQ0CSBo2ly3q!CnKF-p_fOg+wUU~0w~TA|R}gbdwIM$@sM?)Swa{smd~*7T<W+c!XPfR(GLY~svIjJmK<Lqh2O3p(_S2iPGA>7)**Bu2G+G4GEfNh;)qjnT=m&K4}VLlV;JV%%+PWNPI!eHj9ruqgXhC?fKE6ZX`cKA}f|6Zq~z=7JeO?>~|OXI+4>osYE+{f5qa16=t8m8pj?fN5xNZR8@1x~Jwazm4d_lv;V$2ynwx9McbBT()E*muLzA92MUt@K#5W<7gtt#Og)}`3znMDnecn61vjq{5k=+^_!{38=!oe7+7Sr@XG?y7l`^RMS*XVAnT>ml6xT%@oC1ln8c!fl2}{@n}ttyFV#YXDHu9V#myN3f=F8<bIZb}M*T;!&ZwKFtdT;S4WuIUM}h|vHIul&#q@O?CK^`J80*9ZuFn_T-(;*CATr9?OkDqkhUP?xtVRTi8q?5Jl)gwF2I*Lhu$?N~?zlhznERP8%lwtd=PpY5Jp%MkGKNM(W?xma#AjP5p+t1}l!>5CY`BgH5b0P>=vZ~l&%ONados)<LsN|nrVm`ypl|Jv8y*Z58lgi^CeDF$&xLH+vWUZkiOkMQc$aj3<AX}7v8A}L6^scWN;gzSzA_V;_{!;9=I#avlfhLdn#tGL-6l<pHo!Ls-%j>Nfv5A0#YY7BVk8sk44xr8!#I`1`Hb$q`%g~@vc;GN^I-5`k4`U>{r;<okcE6A1^czkce3A#M9cbd70f1l_f&QNHc|STt9}9;Fxo$fNMMg8#*ipMF;Gh;M^p?a%an?IQYcoABC}mb=Lee@uq<w}yb|6Q0Z4VM%QBAe$PGy4W*w}~;(6cA)GQ1W7W8gkCd@1t9ZWhuOyr1CFUgbc6&hdg6d9>f2ni?6CPJ=LxQ5}4F52%JM7l<ZeVnd*G~l@)i#uJP0#9RS{*TUa8Tx*jxkrG}fY+=I|E-k|eIDf)!ORWs$0j3Zo6W0il!kepkx9fCBbY@Zm=nKOfb*I~{Q@|F(pL0L=j|q|Q1pOi(!mhTBc*X`l03Z;%Gd@z@4Pp2q#NG7-f(T#><x3l&lA*m!zqn>{-G0qbuRd%j4K1oQS{!Gle$?9v&*WG8QXCq=OJrVPS$B82syzd(6$<M8&3983JIf%S-<{XGU0rHGvh0ir^AJ)9=gHgcY}WGCfjcJAnTG|0_J%rDZX)uGj;&>3Xn`awmkVg9U5K-V7cK*g1`;s=u?wO_>vQ}I`p##n>iDhOc!YQhtT#5-(n5xhEr1(XsIJd`OAn2PnLv8Rvr)>x@dJKX)h*1+E54ABBf|#Qr7uHC2H?X%)fbl;-DHWq|mWhp`*d73(xnPI9bbySovnk4Zt{gzD$q?D>NQol{O<fxUGLP)pY|*9%es?@`^Gs`kvu|7`@b5bdZ<K_lqL4p5V<4Pg$XjZW};=9nP=S;~Q+*fa1_z3b9g<zgE?H->e<{y1o=*VT9!spHO@T=Qq;#=lO+O#C|)w;4HMQs8rtSRLFM*oQIOm{BFG9G*$EbcP4Srm353U!V(y(p-~poZrGTFoOKspQP#{!$kM9u$wYcz!lV&jW|eRJl<x<m&V_93)~q9QAMr&CnN9@Sq&74!-dTR&vR3{k+p2dbtnG~gd`n}9Fmo?tNvR^db|nHn)J15S)q7(1U9Y$HTVH=M1^J$?H$A|;u1Y%Lu#o#*aJ?9!7FK>3DN5H1_~v_G;f>Go=O^U+>Qv!tGif?MfdgzL!ZLP%KV=~kvXD+<=t-7I99{(&a0&8D8EdboM*J0{=7rx+FC6ABW^OgCNyL*D+cf|()=NS7VtZW(vl%bn8Vtr=VM+uDL4K28txLL8ZsHRM${~%uA&V_sF^plxiuh)};axA7$PP4!5hRx>!;+R{A?yVKajds0!i{bLI<oJFOtm&p7em6tiKYg@=LBCK3ppWRbkG7%aLwUnlw|q>*ug+%$>ErbF}BYMbuTOeOpS&$h=l8ik!1h~Ph8e{-$rwg&I|-}FFgCeoP;c*A-vdVe}h&Hj6R+(kOZmp3!ErDFw+4le1elCA(Oel4p}-@vg)#nUCV&maO0W1*A?FQdd8<LiDPeEhbqD-E}CR<b#i4JuAH(+dYZY@usTt$VZ7cAq^Bhi<%f_t9v~?N8Z^6(4pV&O6l#_2H!Im*(%UIWHH;F753Je)uU2Gw<I)q7(~!<~o{9b91Kb#-Zsuax7np}f&-o<_@%#q&-&!z4SsmYyPVToeA3e}QcEp#n5T@Wd2%qCzOFlnA11X%v%~*jOXGin>E=WRp<)ofjN#>#()Y<Y<yJH~4!&GUaSG_BRse#qIawlXBc-vnta1($!J`^fTvy}!}^^sd~e17F#$Vr5iM1ic3^#P!h6SNp`e0_qBVT@V?IKjn`N|w+#CgW@P2&>ftY^)L^Xd)pBT#(0&x|H|3wUljUfF0W*PHcEEfE)FvJlP2Nm_kT6u9fXA>_0l`pxgH|sSPV_m(W(gm%Aa0Xe=7LV1-)w%?U{O80V=mfCIy<odC+GTKRr3d=pC1uv4s&t$>Y)klyl9*+)kU7A*A=%2*w2cySG5o&=;oB?=D0^mPVK7s^|I?*<5*wm}A%J5%}WAjAv+Cs9eAWO3^m#}HsGMh}CVM+v;fwrVdBfOHhd>isuM3%!B7&Wg?SRrJLcncZ{3%wIrBEh6RGe>MMoZ)41}nb83X@HwXZYBXu&N?Tcc>+;*-;0{mR)zFCmK|-W1WLt#K;RYVsLBM7bG6Jh3_{Od2JN7~r5ZV?IK(o42(RUGOTPT20&(GNfZ=@`)w}RY_57-==OeT?aT_Xz+EJpzmt+aA;!U*(%0uwW`ju-T5tmc{Ljbxn4T#3;~A*Hg-yo`}mLWCY2AU12L%r=vl%`conh+u;#7|L6fCF`i;)=W(R7c{%b6U-$@A<x9c<mCm~&YB?$S6A;iUcDgAGzdKI7ntp&bSP!@{b{dzek;GuuVWD1;}e|IB=P~C<beTf01!F73?JQ@IS^ksznvLv3$y<l-&!9Tl~<D|8;DnYeScgua8|o(#D^h_*Y8gmp=}!AlaHim8e*Y}^-1FrKxkLjEQu~1by)7MAe3(te$tB+O(GbC*CM)+(M==%Wqg|KI_Pk#F<b4vnLc0#c%d1*p_k-zW$R!HvkLg2o6eT8HRk^um)_1OVY5L5tx<r=1Qs{68Uz)h8`;WBx1BlGIaV@Y)QBcGyx4Sge&GgcL>t+3A;f}RFo{dkt-DMnDOpK5nZU%*e1Z!X7BUM~Hr`qYPho(5m&v3?T;QZv{8NZ9Nh#kMJcE_3$H7>HH(+jsjQuc~z)Y7IRRgy!!)jb6u$6Xex?k8;Bfsm#p@k?8PObU@ZniYIY-+elRT1FIKE^A);U86)?`it%V#pIG?1C545iDA;fDfGH`(ukjvIf0Ckbm%5#n<R6jt=<3*vc-ue)js!WSn>C1`D$CNtzkN4YEQ<d4l5?sJJ=Zelwcs0%I3~kH{v4Kw}ajjl>A?l~esZ`^aQm2h_?3mJwwArjS~cJ}=&G#Mj}GW8E(H2^s*zp@B|H<wVG}flS*Pkx;Z@E=j5m?p|b5hzqdt<;q3#0TLHO>o6O`n84WCFqb3`f{*A+2<(C>LcJTllpCM2g^vT^&G>}it{RcwRPxVf4gGYJPcqhfv#rbi!0*qJeSpsUBu~g@G0es~CBAV&f-I_`RKnC9pO7qMJ%~_rn#%vO;9Q8gS0(0kNs|;(nZyn@3ngAK3R$ylN&Axy^P0eypxe%rGz(uhk%H8~Nwu;Pb(5XYRH}F97{+kxd?Qtk@lm7g+9(s@p?_g7FdOTFBFrvZU`UZe*q)z-A=XI82O(A`_Cd3_0adJpGW46MH#D|RaXZ>a#0TOHot&StWt4`La?F99bPB6Zsx9Ok&-lPuNi)0riqN(|d;lr4r;0#2aZdUmhFPY}h|fN8@)zGazaYfQDvL+9nmlleO$MuC86kdw0%f8`)Dh^?^Z|Sr=)@5Ts*@!t8=hAnrG{Rx)m9f1I9aSt38Yk%zbK5c7q_WSE+gkN-EgK<?#K9*ya}o!G}K?wx-|W~aD$A_Z#=(HXER__?mAf=s7raaA6l)t0N^H1urc|IJo<f=Od`OUFIa!o3hvO#6p-33qUZMolS1zuT;vZh;j>SL^7UX-EAnHGB(>2%bEI?5IGNcKJO=|NQy0DnKFpCe-fw4x@=eYUsxW7rPXMS;C^urjWG-%Aky`?Yt(r+Jc+<7>c)t;zK$(}c(i=gCbHPR3#7P9RnKDA3Kbc1+RQh{zGUD^`QM#dHoqQ5)34O3SKEfN({V7B+m9Xtx`XXnT9Yd^1*0oBMWdBZj;7m`a^ZOXWIIChUj8I3>4q2oRM<W=-7y3vBeJ-!7k}?A*sMm~qfW4^K%wIKXCJ2FFgC7x~Wq_O5FB$8a(8=srW*elJ-dKb%uRc(oO<xh9e;Vy~wrua-;1I7^P~d42{U5@poONs}WFZbo5{{#4RPX&>yk_87>i`v(`KIXA2<H0_&_P*3DmrUSBk-2%jQAWoN!0@?n+e%C_xKFH)A7N!)~UY|uW@LEOeC}IBG8t%E`?ojrHpD3AGG2M*XablOvEsU&>aVZCkavIo2ByuUf{racc4=`Y1BANk_^~70~PM(gVZ=C%Yw54-%9E=4oQ-esSsmqMa$wVtCvQN5c}0%P$e|7ogCPd%maRNeg+Ok<x>{?-2vux94m>#ICSDznb+CEWrHf#KzX7`)^P}FCB1}>62NrB*oj;In_ec$>l`0b_5l+4DYnf{@Qq6*ED&UoDwZ%E*9PY|#z%oMVL|>K4y7-!MIVi=IHHb4#eepOGT{2W&6pL5N(R@{=Z6xXWSJs8K8Gshz+V3*_y&7H93zAf13J(t8(DZJtkDG{7ZV$qY+q|<TpdRT-J^)Bmqy$mWVXp(xB=hawvYQfHM*)V`<NKzhrw2dB-87@$IMX*A0YIq%eH0Ks-|3>4ff&Dzo7KRlHZWU?z}TOX#4;d$$_a*^!TBO<uG+-R0LO<e%yo?q_l#^{#BLjW|(FRiTEIdVSIdLA7BRyDLT<a<(Phdr4ig>2#8tcBuShuwr_M~e2$&!`8$bEN3zb0QtDo^Ot!97*sg{pKLK=kkPy)JZ~8Qn4$wb^SRM!7Lnj?wph6>Q>2Jy$r%{6k7!4$|DZMT<=8X^-Ix7+Crc@4<&t$+zm5>H(%0)iF#g<Jqn9A&!ByFhY+yE#002O0=kN6aZ3P~G;+3y5topOAHCn1chxV{xXJLq5{j!exf3r`z>!}xxYndQdvBgx-hoLN5bRzAE%CJHxI_XHwa3ICjQ`DqspYD5y&Q67LB06OU+O-3e%Hv06BTt(W3+Yd~Tdcp_lfVO^BgEknK>L9$r0>G2Wu<x(UZ{vdl(6jlD{wwJzSeLmv)?f&a!y1_xeKoYRGE4+0pgX`$aIzwdf-naL_{)CPMcUZMY=I1Y@WI&pAi@o_3XNTcQ8LW&8z$CkPxKeS8I>gtzyaDxgz3aY=of0+4JfiXVPbVQ5SIh2xx-NK!8ZXPIl#^OtXbg`wTTS0eqSp7!bILPJ`2)G2j9y01{3dEq)oU%JE`LV4x*CJRQ$6%<+Q-1Jot8&PU=J)kcivef<RH864<i6rs#XKHlYF;b;gCBPk_0PL>7APK{k%s$ZXGPn1ir7ewy3}ear;_S3E&f<{0D^p8M0Di+OQl@+<tyH5|D}K2J?1!h=5gCPhH9PO(7<Q$~WFmdP7qlQ(>XqXFZ<85`&#5wwCjrNYW;AqwUCV7u7zl7<xptPi{mUF!WSc(*4{4+t((T%^P{oN#$zT0lYyd;l6K>yR27p(=J5xy3&JS=PpCfGkCV8f6Q8>H`q8iG%<gjLa+0r*`=MoKWqQWk@-~KfW};>+Mqr<BUqF5TpuZAl~XmZGEhdtL&zSahd~9;QJ#hBnVY|lzAlBhAJe&5;oDQ5CR_gzpv{R$Uzn@LDwA@K7c9KaTl8sjeLDCnaRuI76AN`=RE)iq;Q)|OtE#w^3%AP6Zug(FiLM&BEAVp2<Zb?-5MXo!{j3LsezCsA=^IgTYkPEE8+uDjzu1j$wrn94p4C)w2=c;wh%RI$DJ209UurNIotDNWawL99w`%HoM@DrkZpw%)_Km;zX<qHqP#FA4j0<V&UBO(%Q5sB<z`Quu$k`-5zLn|(E*4O<&3O7$=3M+NfNB{Br~^`w7S0lP7DE~1q&2m>NZMY(+ma{(y#Lh_!fCWddwF7;KP~Vxx>QN0}!$}HZ)ornLf%J+{@1IUy*fAG?Yyw(s>U+P$t^s6IAvIHsC{NvZ7Dc=pySzI}D}I6{9SHEM(M5(NW0R1mF1i5uWuC?nE8XkxTi+H$38lAd699$V%<?z!)nBft$#hFe-=Ni+j@jj`RUSjJ3(gDA@^pk9MSamE;1YH4;WkB;I}rUqF};r@o&EQ2=rfLiuD)x}!2tc7k48U9=K&zGVe~EgoUcGnE8di_j@?Yx-*S8J_(&iMj#bc#fE63gk&jim8i9q}l`E4Qw>9kvKXrkhjInkslES1=mh0Z9_ly3R|7{v0qIEFN_jKb}<Tt7#qJ_=s6%w-&o6MWD_#gI<e5vgZ9(`xIZa;vN-FeY9tZnQcB&wVmsl&(dPkbx?ua9$F=#fn%}V7?{pcPF}4nbF*9gmi9#=F5XPsoQs4$gZ-Z3nRI3T0jVZ4<et{uQ)grLSq;(eCIO^zO6q>;{eygohZg@w45XN+p%1u`5UD-lFBzo78iy$1G*8vJx$k)o-gKt2XMKUFFfHP5!r=xh>rXwqK*&cj=J-TRxOs>;)kCPps!OTy?17ss1bX{Rb7nnXku;9+LP1wj!Zl<>f3STH2`C%?v3=CK!Jck#K>@gvGI7V4BP{?dd`e(3}!I{l=zQ1T60(&G>ez|A`n;})XhD;T*k)%A*FneIL9ehwKpTE=P(A<rCrmy~sRs%2PCwQk^6X7XrX&@tAlkG7vSk<LS*tY&FvW@~9>j1O}aAtr?KET2BTa$gvjUQkq-9lpyTJVB3sHh)+EAQcO(yUobCLi#17@4YjE+*T8;85wE8noFCyq)1yKJOq|=5acfx9x$vqL3JM5E1o_t%iA~B_qJ(0NGgZAZZdR2+v(N{V;&FR%xi)Fh}gRD&%qP&1p!x&NdU-hJA%FH3iJpd(jG!@xRwQeeYouOt`Lu{$`U-*L;8zr}m0|>vy}@6aiYi;oR*~vg@q#PcZK*+&Hx-bQesel)p6j(hbk}$YB~+)1*$9VYQX<4IGP8vTTnmB^OMCN9?C-Nt3iL1Ds4DIzJ(|zD?dGroS=a*T(LIucz~ag_KHG5ltrJpBR<ZLf9AB`KF7!t7i7>1am)$y>dcGPaD*o=Eli+sd;c`yQUl){$~U*;=>@p2|l(?F>ULb_s*ld?QB+n#+G$;wBVPES<e&>vZO{R(Zcs2;3Lv;p=_P(Ck?sROi>Q<XT+CwVHx=V8>#ia)>Fids31Vvhy^qGxQ#<A&sX@5^vVqof`nRW6T&HfEqad!xXTAfcHYugc;F-VN5b$1xZwox3tVipbwep|E__59Q4eX4^>N$TaWM&>WqgG*yVzv#YjtDt(J14;cfDa@>*5-k$e&8g4G^LPBPQa^#gvt@W#=;48r;~XhP<rigfSxk#7LjHK=uVXB<lD9vX#6>ml2_h9`XEgyGBT*O-U4Unv)KN1VQPBUe9l^I3FkpUjlbGKnM~BxnR0+bU-C$EEJh#09huNMLy?xOH|$fKeFzG6?<>7&G-do+IwiuBp<`}iDh2?+;4S`8z6)+oornmZ0GI9ToctphzXNYmSDAfy@|6oJ_uu+ydsOOP0v2YP$jjOhTN98+&3(eFgQPrU!b;6th>f`2E(F;X~-?_WFM9?-lt;Ab-im@H$Z6D)WHVfyhTS|CDH<`zD_49!X4WU`(5m~@j<ayNg%DPR@reUB@%mugjs8SZa4Tl5xD_Erc_?x-PlI90h1SZ-#8y2S?W2NKoKatxA&@n8z3y~Vm^kL$jm;bN<2*>ve=nS>a5d+$;ae|jn-WVyRqTD0ZnyWOn%BmgV+dD0D^Tt?GfV`Y}hRF%u)S@1tMecicGrLeRd`xg4{|^2ME&xoh0jPev~Tjn5=&-*!#JlWro{o=6ky90*Kte$rj7f=_!~GvGu9ru)o>1vgS8x;AdeMBufM6A}=@1wjp$o>Mi2~b!nV(m6(MC*6tA?EJkIy2o>3Z`KhMp1&N6mQyD8NUI`woMYU=n&U)4g<MR{eYZj>%t;t%ND5DwTEbJ(45(u(-Td|Od#f1GJJL9O9f5&}TArn#H25!PT>7--mFrPHAmas`6%@$X`)>5SLOw))DJdBZoFIB+MA4Q55SzJx`wjycb8iEV>az|;CwWy0NlR$xqd$D1bU?Br&oefxb6#7i~<kK>V$p`2+;B%btA4CV2DV1y4%NzY4iWpYX&LHGtb0FTW2g)25&cxnC=J>T&E8&aoMSvS!8^UC7`nkOIqe9WDqZhj~#>Bds1$omzG+zPpZKRk=855|_vE<|QStrzEn>b$mHT-h}1PLk`zac%Hk%S53-wv!NV;dD2JwFrgjuP7p02LHsHfaejLN3qj`~~l?vvN7>_`Bk+0e`D#-uOVMoXL=m@lol0A*)Tu0XmycC(AA*;v1mQLCE=pSg9ZczCh`nlHY~I&~<uczNx)8K-A&$U(H`=RyxLZ17ihpA!dOOFeowK_|_+-e2>eBzdJyK`9fE^4{Or{G~728ePGIxworwH#m0XCAIz`x8snQj*!YSfybvMu96^_T_IgX5mk3&9=1R=zKO`qDvg#XmPPKSxk%7CvE#^r*2hpZAPja>uP!xzoXFMuhvh)VmF)!MD6(c=AE0-)Uo!W>o8nTgG&YeUuGT<C-0XasyRQiduL2Ow@1I@~71zrRhjrlC~Lt~M)h~}cqo~s_W3UUD1<lUD40x9iH4Gkct0?X*I{1eD%(H8?7gMN`tHrU7nX~@H^P1n6zWipB~w<w3*o9<iM7g^?p8K36!ec--w>CE%U;hDzhK}10wO%8CZFv!@1vL!EaQoz~mt)ze9Od<BM+j^Q_JMQw@aQBAk3smN0Xq{mf-Icd>&OPtEleXXU1vf}ZARU~W-u|Qt{>#Zdqm4_GEL9+L1CQx~$yD`abtu?Ko3X0dbG$8ml+Mm^DN;}Kp705G-g9q9o7ad(25X*(_dG2m+TRn;H^}HNkB2EdQR$0BJ%S5lM6vcI?nQY?1`B-THGv#(o*Rr7l%ntBW6INbJX+-=iQlU50q2Inm*bE|InS#PR!zvSpB4}F`Rz*tx)V?$)qSka$bsfWS`;fK2WR~{r^es<eV`zgB4Tg3nH#1hn3EV$XS2KqW(s!?&MY}|a29&Q^(-)$IieKVZ{5`!XY!n+>LgQ0*-3Te^&`@eW}~{vW4R_#uT|}^O83}N`aV-`Bhw({zks%y-Kc67Ae7#rz;ph}@Pf0G3`bR<#&hf_Ii-~2$nx6Hy_{e@(4m?3b-v=A?WhkRpd6sP5%Q*vdwMD!=y*iGlyaQ|sj&NG2yyZ%o`T_H?wR_a&v>XjMPrdkHx<j9U8J3OFQS~udn&UcsmkrzKsoz-_F0=>z5?_wlzkEk$@zUIgdgp#kAA~U&hz=6l+5;ei>)5!f)`UA)|hEyaZ*>JFr_lflldhBeFII-Qwb&UyqR^`l|HRIwGoaPWij{T5uBZ9Eo5lFH#A2x(pYQmx!G|A{7&3w?Z@cUJi~RqVsP_A%OK|YKD`tqbqcfvFLRHtQ6jLNw@S!kRQp<bo+tyD;69(mQD4nx08tzK*X_Xb-ePZN<Z+~9W=^7z)~i>V=svM}aBvoV#S<;*u3l-N`)o5>(#RR`Nk+@#J_a)+*9RVFo(?{#Tj^w;w{2H1@A4?YQusOtWJTI)7(ayQ0O{7eQQ|zBc??|D>U_<4*{t6B8c@!<3IQ5%ZumNl#Y*Hcny{GT097z|uRQaU7uCkddABl4@BnQIXo{Aa(7b5SU@0X=Ck(?Nc~%<FRgI{@{7qPiU#Gl>=5ClNSZYer%oeu*zukNTE7Ad6ne!5qT@#o0BrtE1SjOoVaOO0#H6RLI?-5(LbAU<?=6eCE^>y#PDp81`Sj#hAv3Qy}HEOE}avzYy*4xe72_P2?shyXwP~^{P){sullm*;+f9n<8I8&(9dR<_7j;R)gPu7)cAjbkHI8Wx4)W~Bb_tvEBJ7@Hri3fMSiBcXuR#bu|Ic?jN8Og#+Z&EkHU$1gK&pDwDbyWdI#JK=xJ<diNp>jw{>EnFfaxT`KxUX8#Gt8-P(McheBF>(yXoIyEr=&`ZNM0-c63+Y<J>TE$<kl<wR^|c<{W0eiAET5Rxm7C?BeMJFo?3F94XsE{{j+|YgVjNX0DI-=ne>QKa<#LJvmKa~s(#Y%w4>~@&a4Zu$xQFD$jUUls(=IPof2&*G3Bl}zufsC3TCo)E-pxeHsmP|7JZG&((9G?EVO%y;2tq%zd0(A;FxC3*=0*s5<T%19HrG()s+zt`@yxBm;EH1*>7%QKZAXl8?XHNW%PWOJ-@v3&I7AfiA=okE8<M%v*9$Ayh3&F!g)7ZMv1fBr@CCmJkDF$CacPhg45G!+Q-veg#K{8%*3xcsc@c&Kfj6!InP>GINxR-%9G}itUEG~?l_gQDd!})l?^Mer8?oYY;f(nZUQ-F-?o|8#N)I_%!hJ*Da5yTHXi!NKmtgEZuIe*qr%)|fbbN<;0=<mDb^>>nLAnM4KkLnSDdS5S?F`?xtAqnQR(aJByZ)omyG0|L5MhKZf2cZUnIX=!f^TmoLTwo9Fno)B)5J$s+?!oiZB<1a;xUSjKp>o$WrEZp77GzaPoE5frhKtAedl!!Wo^C?@Dz##!gxr!}D0kl^JB5!`M5tiJb6Nf!t_CGhMr7oZU%e+=TTov#mD;UeUSCqGfXxF>A<qnqM28XS(;?(rRRf`c<G)7A56G`1FRdSDKyKdt+XH;3nbAEk%3JJ*_rKokkhXm()&bYOU<$&-}<BALu3r>cxF}96X$uyD^+2zQ=EzrQzCP8^(FE%V5w|C=u2e){f`76VB{EXN`{gV9lZ>P@j_yLJC)3q~i3bo){g|!1`R^mQBpvcG>N4ve!l8&Yo4%4$cktUUl*Q#KD9OPzp+#juV3rb5c-R?Alw7j|HZ!&KKNjJxS?tE)42U#+u9sgOHM9_RTW(wK5gy;B471Ev+LPC;jP*1)p7hsC!N5YSCwN5{-qcJ0<WOwZgAZYDSZ#!f@^K7PGox3Ot*egmhI7OkQ)l={b^tZoXEzPv3sB*t5EMPdL$gHs<D4>3f6(i+t^4*OWsyS%rOmqY^7=e2h|4>zUluGK~%*!+&b(sjl$iJw^)EtX%c;!%QhXX+U*uMpI8cEjM<Qoq5jZ8`m?)p5wW`$cL$eB*H^T<1O1yia)%HAYE~9O)*Eyq;OmK3iVECKVoF;<QNrj7M?IAqV!MZBb`v2uPAqB&RpJJd)17}S;}CUa39bkIz)pv$dx-UuGM48Vh{RQS=iR@5GS6b)%aaDX|$dXyp|O9h#-58H;0faQXK_7fkLwH*V4Qk7MDS?lY!PPCFkMCim}F-?v=APXPlo?5+pEN&edOQ>$4kjcACAyeEwA_7TzPu<cH*4)=HmcXSS2x$Rl)=zsk`_73n;E!fdCfvph=J5<w;l@AZKcXH>MY@?hg@1xP+gC6(gIvWuJroj%5j5TkjQZF4UtL}K<6EAD)d_))g1JAXb(QYY>cTC0905d+H96T+FR*C%Ppd{B$G{hCV8_T7Ye%Ef%pLXZh#j5EsM%j+a-1Ea%|U_;vBtJ@9`o$;Sf8yq3dgfVfnSu;o)3)yt%?!`-!D>`#lL3o;A?WmsmZno{Syem4&yi1m}``ct?j^ps6H{W+$8=2%r8dlHflax7L-#`E<u~fKVrEG^aewvHhk?+28(P%g1k_%ZIccm+f5#IAq&5mQY3)aeZ;JE1R{W&Rr;8Au2SvD`jy}Uv8%fknhJ?4)Kmdkb$Q#N7dEt=VzXG9l{Qmd%-buVOgpaI+P$UQ`mo<S}l_uXqTQ=nS3(Y{apwm9>)*sJ(VZF22xf9YyoAX5&pm9)+q(>1(O<@NJ@k2<M&<$S{KJ0Ilp4T>0h9%I>?+&1Z~Jb(Gr>D?%Bef6#CNp>P1cuGQ<hH(0xGol=Mps!H4+^BUcY-1^@tw`0KI5GD&_DSZM@B=a$PF4yZwSwti`-?Nk@zYFNoVWOzW^vZ5I97vGX&zr3#hoYQ94EW$G`M+3+!<kgZcfK@F^FYwZb!?`;t9J{*_@44UCu}HP5Z=~u8h~OPJFg&m=6EiU35Xr@(D!f>ntE3$2ii0(%0B$_C2*+*`$i<U*AlJ0o@Q<?6QBhLR^f2f5<NJ6mh@d^<H8;{`{;-lKA{>>x})55J5=>%alG?SRwZc&7XYZu8atuMEd{yGP(d_s0r*M9By?!t#0AI!T^JZFeB`h0^>&3fL_j90pW-M!Jef8+|GiYKH&`Uc_~wp&u^MI9AxO6Pb~r5u!YR)EiZP-=>U5P%1C3>`SZ7MQ|Ww*uipqQ51MKPNJcAEa#FY#aBfddP{@M5s&M29lry<)r;_Q;Ng<mh_Q71DV!GAv@zYT$PQDovOHpF~k^l-A1@>{z3bBe?p4?+X+S+mup+tMm8Pi2aX1bi2^=Qhe%e1w6n|L+vY)~^*-7cDDghzHL#GYWx263SHNE^(EaO7;T&<NUQ$>1qv$Jg2KYHaS-wes?bu$igZVehc@NFvyf3_9V`Kci3|tca52tGq$nJy`j%InloPLtGCaN0fiqkPOZ``+O19SW$dJWv}66WM>Jh(Mr!WNLa({*@j$jk*s>NZuhV}jdJh7i5LmRK^ld)f#Pt!?r1Vu1A@q&pO0t?V-BPpruWQyu9sE1h|p|x+;Ln0`MkM^QSu0RWeCY2N=(&iCg{CgjcXz0QNq2MNE$dbvG@w6M9_$I?eTM+wq%pw(bba{Y;o7jUk97eUiX^GT-S@ntxysd)=k=j@xqn>Wf7S-LV8|oG4|=3W^J&uZZCpN1#!lzY^*x5W(vNqNo6Ud1}EH4GO~HbB@|9lWB8qzfJP``r(m;KlQW4G`@|0X@VfNlgc^+QJJG&ZoOIv0<F$<yo4=O%<PL}=l)c!_uDg$p?J$%IRr}{@#lRw^gWId0lo=bmZ;Y17Hs>S@DX2)jNEkrMB!Mhsfe%hf2u(;~PmD#8t7rXtw}NsKri~^g^Y=$7kj<P0Me!)pfSD1lywDQqnevZuq@yg-ai*P|Bw+(h!*4QCXf0=4+ejC-841)5gV=L;k(Gyz;ndlkXBtfoa!$}dJ*9bxAV;=t_I48&bozoTzumRlRS6*?(Fg<a&lFE`MUW7wm5$u0a#h0_u$JbvWjl62M%~MG6*4zt2EfDk{8G|&r$}vsvdpXcNP^zcv!gsHJ5mGc*iM^_12hEYA}qa&mXla(WDdJc)+&B}bGdUkuG}hKc(3cyuc(*yU=^%;#-qfBmsd}ztdRR6_k4kLtnOV4^U7DbqDXC||Kyfei>xn$mY3Bn5*J~}nLkICMFS4q*An;KDOTps4fS+_Fp=eC(X5IbU&Y|fP)=yNLmPGqnQ)cySae!v!BK{pB7=O94eJPNCM}rUcxjHpNjf=6M)N~<=TQfVUMQtG=Rsm97hzGs-jJ8#WJcLp62T@7kX7SJ<jHc&%NxmRL&GHpm1%s?3N!xsmMkZ4<CsjPCadY35^a2TQqPO>7bY8fI?8W7KRXKye32IC2`VRzd>ugU2r`e5O#-r$apsGKxUv_5mA!L-OPjgYn-iT}tvY%w38dsir%pY|Agk4{e9SXSFDTm}_bn`3%M}9OaTw1a(b*5`&nraD8lW!{NsiVk+4!hObT&$qbzTpQbdWQDl8s=bT+UQ~wlYS<D3YgYt&F$b;Cr0W0i})?<Ek711Hw$&S<Cn7C4sXYc+V-IB)otzpsMB)B^4~Ei!7%VB~o4!J^4n4#n8G5{vtRl4KF6Vi3uW*r<isd(ZQRb3U4S%a#aQ=BlS6*B|wZGFm@OWb;h1@W6yS`>?*ckPf2T$G|cJusPNW=R<`HPoIA@EoyceW5iyG(wfLTTl1H>?vUgkZ7dZpX)ImF=DXFq(j0fHz4HDXpDb8y%1%O=a*<Kw++I+-S51y07gHqVckkf)ys-ijy#UGf{NWn^}3Z7t9EuH0nfgxw~XLgt?PddC*HsLVkcLP{+12w=$j4m7mKH6jhSetRl7c81~#h`aJQrqvsXnoT`>e72`*E!02u&Q)dMylDG#;SdSs*9A()$^4`4jKjFyBMe-jv|suDDPd#TJ|$nofJ9G^Z=q1Cn!nlv&bHv7H$tHv$n_6oeer7_EDAKbY2IRdyS%WhEj?mXFAM0ai^Wg^o<@rrXZT(Fz>HLkW-294C$6wFY<RSbRrza!=4^(6Obgc_4&wiEfYFAHmW4>0ojR8$}ZDPNtD20IP7WAo5&53daB>0i!@~kFdLzy4cad$3$w2<(IU*Oiexpg9FvI+othTHgfjoR^GU>;Ncqb3QLa3Fm>WJ!cI8ZJNUk^e%YJU%{(P7<a2w7@EY{1&$c$Up9475f@>yUi>qgkf`#=sgz;N2X7TE@!Tigrc%X2m9oy?PA!Un33Y-mkz?^*sITqc{MA79G`vM^2s89&TC7A8bps?+hcJsuxCW+ws1m?Xkn$i=v*t~-M4@J6!9N+Jinv2lz0RgUK`;@>B|A+OIavXLi1i{V=q^>~Knwb(ysap=+!&-Rm-j>@K|xnrjH#(C3}^MJBNpP7{-rIYf|^TQmSU3yyW=?f>*7i;H*j3{4wA2?b$qMa*-Mw1*T=`wE^C-8|u`z8XHW0UP{=NE!8>Wh;aQ98{1PI8!yYI-E$GcR%kYZ9Kzpw#VCCGj7Ws~uE5+9J8dpq1*~v!yGR(-+y7dZ~a9Ch7qSDJ#tQOG7P2qU17L47!4)ckr1V1ZRW{ULEO*igGKYy!S|UgsC7ISs6J;OF~#Uob-Eq;&Xn8cO}w%X2Oeh?WcS{oP2Q3kIBWLyxBvP-QYO|pUOJ>FqNfhXKA5e6!*oEJ~$WTYUHV$bSQmf8M$+w9ly@1l`)mdvu#+md2xO%;>4V?E+#mEtg<>n+G~E*9nr&A&pA5_;}%X)%QLaw5oQjMT8-l-Re#?bM6jA-1@0L`k<W9oA&7;k?^WGi5vPM0O%lj8sv?(_9jZjiX5F*FG>p&ZX$W%$l%PA}JQ3RL&c#7Ru9xtmQgp_5rR%Av{_g+zpZ^EAW(uz'

@st.cache_data(show_spinner=False)
def load_prices():
    obj=json.loads(zlib.decompress(base64.b85decode(PRICE_DATA_B85.encode())).decode())
    df=pd.DataFrame.from_dict(obj,orient='index')
    df.index=pd.to_datetime(df.index)
    df.index.name='date'
    return df.sort_index().astype(float)

# Disclosed consolidated inventory valuation P&L: + = loss, - = reversal.
# Unit KRW 100 million (억원). Q2, Q3, Q4 standalone amounts derived by differencing cumulative disclosures.
VAL_CUM={
    '2025Q1':725.53,'2025Q2':33.55,'2025Q3':241.42,'2025Q4':250.00,
    '2026Q1':-325.64,'2026Q2':3431.26,
}
VAL_SOURCES={
 '2025Q1':'S-OIL 2025년 1분기보고서: 연결 재고자산 주석',
 '2025Q2':'S-OIL 2025년 반기보고서: 연결 재고자산 주석',
 '2025Q3':'S-OIL 2025년 3분기보고서: 연결 재고자산 주석',
 '2025Q4':'S-OIL 2025년 사업보고서: 연결 재고자산 주석',
 '2026Q1':'S-OIL 2026년 1분기보고서: 연결 재고자산 주석',
 '2026Q2':'S-OIL 2026년 반기보고서: 연결 재고자산 주석',
}
# Disclosed refining segment raw-material inventories, carrying amounts (억원), NOT physical volume.
REFINING_RAW={'2025Q4':26318.78,'2026Q1':36740.57,'2026Q2':41291.05}
REFINING_RAW_SOURCES={'2025Q4':'2025년 사업보고서','2026Q1':'2026년 1분기보고서','2026Q2':'2026년 반기보고서'}

@st.cache_data(show_spinner=False)
def load_financials():
    out=[]
    prev={}
    for period,cum in VAL_CUM.items():
        yr=period[:4]
        standalone=round(cum-prev.get(yr,0),2)
        prev[yr]=cum
        out.append({'분기':period,'평가손실_분기':standalone,'평가손실_누적':cum,'출처':VAL_SOURCES[period]})
    return pd.DataFrame(out)

def fmt(x,digits=1):
    return f'{x:,.{digits}f}'

def quarter_index(period):
    return pd.Period(period,freq='Q')

def price_stats(prices,periods):
    rows=[]
    for p in periods:
        d=prices.loc[(prices.index.to_period('Q')==quarter_index(p))]
        for col in prices.columns:
            s=d[col].dropna()
            if not s.empty:
                rows.append({'분기':p,'지표':col,'분기평균':s.mean(),'분기마지막관측값':s.iloc[-1],
                             '분기첫관측값':s.iloc[0],'관측일수':len(s)})
    return pd.DataFrame(rows)

st.markdown("""<style>
.block-container{padding-top:1.7rem;max-width:1380px}
.hero{background:linear-gradient(115deg,#112c46,#145d69);color:#fff;padding:28px 32px;border-radius:18px;margin-bottom:12px}
.hero h1{font-size:2.1rem;color:#fff;margin:0 0 7px 0}
.hero p{color:#d9e9ed;margin:0;font-size:1.02rem}
.guide-hint{padding:10px 0;color:#526474}
div.stButton>button[kind="primary"]{font-weight:700;min-height:3.4rem;font-size:1.05rem}
</style>""",unsafe_allow_html=True)

st.markdown("""<div class="hero"><h1>S-OIL Inventory Risk Analyzer</h1>
<p>국제유가·석유제품가격 변동에 따른 재고평가손익 분석 및 위험 시뮬레이션</p></div>""",unsafe_allow_html=True)
st.write('**프로그램 개요**  S-OIL의 재고평가손익 급변 현상을 공시자료와 국제 원유·제품가격을 통해 살펴보고, 재고의 원가 회수가능성이 악화되는 조건을 가상 시나리오로 분석합니다.')

if 'show_manual' not in st.session_state:st.session_state.show_manual=False
if st.button('📖  프로그램 사용설명서  —  개발 동기 · 분석 방법 · 예상 결론',use_container_width=True,type='primary'):
    st.session_state.show_manual=not st.session_state.show_manual
if st.session_state.show_manual:
    with st.container(border=True):
        st.subheader('프로그램 사용설명서')
        st.markdown("""
#### 1. 개발 동기
S-OIL의 연결재무제표를 검토하던 중 **2026년 1분기 재고평가손실 325.64억 원 환입**과
**2026년 2분기 단독 평가손실 3,756.90억 원**이라는 급격한 방향 전환에 주목했습니다.
정유회사는 원유를 구매·보관하고 정제해 석유제품을 판매합니다. 원유 취득 시점과 제품 판매 시점 사이에
가격이 크게 변하면 재고의 원가 회수가능성에 영향을 줄 수 있습니다.
이에 **실제 재무제표 수치와 국제 원유·제품가격을 교차 분석**하고, 재고평가 위험이 커지는 조건을 살펴보기 위해 개발했습니다.

#### 2. 프로그램 구성과 사용 순서
1. **종합 대시보드**: 분석 분기를 선택하고 핵심 지표를 한눈에 확인합니다.
2. **재고평가손익 분석**: 누적 공시금액을 분기 단독 손익으로 환산하고 재고 규모를 살펴봅니다.
3. **국제유가·제품가격 분석**: 두바이유, 휘발유, 경유, 나프타 등의 가격 추이와 제품별 단순 스프레드를 비교합니다.
4. **재고평가 위험 시뮬레이션**: 가상 취득원가·예상 판매가격·추가원가·가격 변동률을 조정해 평가손익의 변화를 계산합니다.
5. **재무적 시사점**: 관측된 위험 신호와 결산·재고·자금관리 관점의 검토사항을 확인합니다.

#### 3. 무엇을 분석하는가?
- **현상**: S-OIL의 재고평가손익은 어느 분기에 크게 바뀌었는가?
- **가격 환경**: 같은 시기에 원유와 석유제품의 가격·스프레드는 어떻게 움직였는가?
- **위험 노출**: 재고 장부금액은 얼마나 증가했으며, 평가손실과 어떤 시기적 관계가 있는가?
- **시나리오**: 예상 판매가치가 취득원가보다 낮아지면 평가손실이 얼마나 발생하는가?

#### 4. 도출하려는 결론과 검증할 가설
**검증할 가설**: 높은 취득원가의 재고를 보유한 상태에서 제품 판매가치가 하락하고 재고 노출금액이 클수록
재고평가 위험이 확대될 수 있습니다. 분석 결과가 이 가설을 지지하는지 확인합니다.
**기대하는 시사점**: 결산 시 재고 순실현가능가치 점검, 제품가격과 원가의 차이 확인,
재고 보유에 따른 운전자본 부담 관리의 필요성을 설명합니다.

#### 5. 자료 출처·분석상 주의사항
- **재무수치**: S-OIL 2025~2026년 연결재무제표의 재고자산 관련 주석. 2025년과 2026년 분기별 공시 수치를 사용합니다.
- **국제가격**: 한국석유공사 오피넷에서 받은 두바이유·국제 석유제품 가격 원본(2024-01~2026-10). 파일 내용을 코드에 내장했습니다.
- **단위**: 재무수치는 억 원, 국제가격은 원본 자료의 달러/배럴 기준입니다.
- **중요한 한계**: S-OIL의 개별 재고별 실제 취득원가, 생산수율, 예상 판매가격과 추가원가가 모두 공개되지는 않습니다.
  따라서 국제가격과 손실의 동시 움직임을 **인과관계로 단정하지 않으며**, 시뮬레이션 결과는 **가상 예시이지 S-OIL의 손실 전망치가 아닙니다**.
- **회계적 구분**: 재고평가손익(기간 중 손익)과 재고평가충당금(기말 잔액),
  그리고 IR 발표자료의 넓은 의미의 '재고 관련 이익'은 서로 다른 개념입니다.
- **데이터 처리**: 비거래일에 기재된 0값은 결측으로 처리하고, 공시된 누적 손익을 차감해 분기 단독 손익을 계산합니다.
""")

prices=load_prices()
financials=load_financials()
valid_periods=financials['분기'].tolist()
with st.sidebar:
    st.header('분석 설정')
    chosen=st.selectbox('분석 분기',valid_periods,index=len(valid_periods)-1)
    st.caption('연결 기준 · 금액 단위: 억 원')
    st.divider()
    st.write('**내장 가격 데이터**')
    st.write(f'{prices.index.min():%Y-%m-%d} ~ {prices.index.max():%Y-%m-%d}')
    st.caption('별도 CSV/XLS 파일을 업로드할 필요가 없습니다.')
    st.divider()
    st.caption('데이터는 원본 다운로드 당시 시점의 고정 스냅샷입니다. 실시간 시세가 아닙니다.')

qrow=financials.loc[financials['분기']==chosen].iloc[0]
qprice=prices.loc[prices.index.to_period('Q')==quarter_index(chosen)]
qdub=qprice['Dubai'].dropna()
prev_period=str(quarter_index(chosen)-1).replace('Q','Q')

tabs=st.tabs(['01  종합 대시보드','02  재고평가손익','03  국제가격 분석','04  위험 시뮬레이션','05  재무적 시사점'])
with tabs[0]:
    st.subheader(f'{chosen} 핵심 지표')
    a,b,c,d=st.columns(4)
    a.metric('분기 단독 평가손실 (+) / 환입 (-)',f'{fmt(qrow["평가손실_분기"],2)}억 원')
    b.metric('해당 연도 누적 평가손익',f'{fmt(qrow["평가손실_누적"],2)}억 원')
    c.metric('두바이유 분기 마지막 관측가격',f'${fmt(qdub.iloc[-1],2)}/bbl' if not qdub.empty else '자료 없음')
    raw=REFINING_RAW.get(chosen)
    d.metric('정유부문 원재료 재고 장부금액',f'{fmt(raw,0)}억 원' if raw is not None else '공시값 미입력')
    st.info('평가손익은 **기간 중 발생액**, 재고 장부금액은 **기말 잔액**입니다. 원유가격의 마지막 관측값은 분기말 달력 날짜와 다를 수 있습니다.')
    fig=px.bar(financials,x='분기',y='평가손실_분기',title='분기별 재고평가손실(+)/환입(-)',text='평가손실_분기',color='평가손실_분기',color_continuous_scale='RdBu_r')
    fig.update_traces(texttemplate='%{text:,.1f}',textposition='outside')
    fig.update_layout(coloraxis_showscale=False,yaxis_title='억 원',xaxis_title='',height=380)
    st.plotly_chart(fig,use_container_width=True)
    if chosen=='2026Q2':
        st.warning('2026년 2분기 단독 평가손실은 상반기 누적 3,431.26억 원에서 1분기 누적 -325.64억 원을 차감한 **3,756.90억 원**입니다. 가격 하락과 동시 발생했다는 사실만으로 인과관계가 확정되지는 않습니다.')
    st.caption('출처: '+qrow['출처'])

with tabs[1]:
    st.subheader('실제 재고평가손익의 분기별 변동')
    st.write('**양수는 평가손실, 음수는 평가손실 환입**입니다. 보고서의 누적 수치를 해당 분기 단독 금액으로 환산했습니다.')
    st.dataframe(financials.rename(columns={'평가손실_분기':'분기 단독 평가손실(+)/환입(-) [억]','평가손실_누적':'연도 누적 [억]'}),hide_index=True,use_container_width=True)
    st.markdown('#### 정유부문 원재료 재고 장부금액')
    rawdf=pd.DataFrame([{'분기':k,'장부금액_억':v,'출처':REFINING_RAW_SOURCES[k]} for k,v in REFINING_RAW.items()])
    f=px.bar(rawdf,x='분기',y='장부금액_억',text='장부금액_억',title='정유부문 원재료 재고 장부금액(기말)')
    f.update_traces(texttemplate='%{text:,.0f}',textposition='outside')
    f.update_layout(yaxis_title='억 원',xaxis_title='',height=350)
    st.plotly_chart(f,use_container_width=True)
    st.dataframe(rawdf,hide_index=True,use_container_width=True)
    st.caption('재고 장부금액의 증가는 물량·단가·제품 구성의 변화가 혼재한 결과이므로 원유 물량이 동일 비율로 늘었다고 해석할 수 없습니다.')
    st.caption('2024년 분기별 평가손익은 이 버전에서 검증된 값이 없어 표시하지 않습니다. 가격 분석은 2024년부터 가능합니다.')

with tabs[2]:
    st.subheader('국제 원유·석유제품가격 비교')
    available=prices.columns.tolist()
    default=[v for v in ['Dubai','Gasoline 92','Gasoil 10ppm','Naphtha'] if v in available]
    picked=st.multiselect('비교할 가격 지표',available,default=default)
    gran=st.radio('집계 주기',['월평균','분기평균','일별'],horizontal=True)
    start,end=st.date_input('가격 분석 기간',(datetime(2024,1,2),datetime(2026,6,30)),min_value=prices.index.min().date(),max_value=prices.index.max().date())
    subset=prices.loc[(prices.index.date>=start)&(prices.index.date<=end),picked]
    if subset.empty or not picked:st.warning('가격 지표와 날짜를 선택해 주세요.')
    else:
        if gran=='월평균':graph=subset.resample('MS').mean()
        elif gran=='분기평균':graph=subset.resample('QS').mean()
        else:graph=subset
        fig=px.line(graph.reset_index().melt(id_vars='date',var_name='지표',value_name='가격'),x='date',y='가격',color='지표',title='국제 원유·석유제품가격 추이 (USD/bbl)')
        fig.update_layout(height=440,xaxis_title='',yaxis_title='USD/bbl')
        st.plotly_chart(fig,use_container_width=True)
        st.markdown('#### 제품별 단순 가격 스프레드')
        st.caption('스프레드 = 해당 석유제품가격 − 두바이유 가격. 이는 **정제마진·영업이익·재고평가손익과 동일하지 않습니다**.')
        products=[v for v in picked if v!='Dubai']
        if 'Dubai' in subset and products:
            spread=subset[products].sub(subset['Dubai'],axis=0)
            spread=spread.resample('MS').mean() if gran=='월평균' else (spread.resample('QS').mean() if gran=='분기평균' else spread)
            f=px.line(spread.reset_index().melt(id_vars='date',var_name='제품',value_name='스프레드'),x='date',y='스프레드',color='제품')
            f.add_hline(y=0,line_dash='dash',line_color='gray')
            f.update_layout(height=370,yaxis_title='USD/bbl',xaxis_title='')
            st.plotly_chart(f,use_container_width=True)
        st.markdown('#### 분기별 가격 요약')
        periods=sorted(set(str(p) for p in subset.index.to_period('Q')))
        ps=price_stats(subset,periods)
        if not ps.empty:
            st.dataframe(ps.round(2),hide_index=True,use_container_width=True)
        st.caption('휴장일의 0값은 제외했습니다. 두바이유와 제품가격은 각기 다른 거래·평가 기준이므로 단순 차이의 해석에는 주의가 필요합니다.')

with tabs[3]:
    st.subheader('재고평가 위험 시뮬레이션')
    st.warning('**가상 사례**입니다. 실제 S-OIL의 평가손실 예측치가 아니며, 입력한 원가·판매가격·비용은 모두 사용자가 설정한 가정입니다.')
    st.write('취득원가와 순실현가능가치(예상 판매금액 − 완성·판매에 필요한 추가원가)를 비교해 재고 평가액을 계산합니다.')
    left,right=st.columns(2)
    with left:
        cost=st.number_input('재고 취득원가 (억 원)',min_value=0.0,value=1000.0,step=50.0)
        sale=st.number_input('기준 예상 판매금액 (억 원)',min_value=0.0,value=1200.0,step=50.0)
        other=st.number_input('완성·판매 추가원가 (억 원)',min_value=0.0,value=100.0,step=10.0)
    with right:
        shock=st.slider('예상 판매가격 변동률 (%)',min_value=-50,max_value=30,value=-15,step=1)
        costshock=st.slider('완성·판매 추가원가 변동률 (%)',min_value=-30,max_value=50,value=0,step=1)
        st.caption('원유가격 변동률이 곧 제품 판매가격 변동률은 아니므로 직접 구분하여 입력합니다.')
    nrv_base=max(0,sale-other)
    nrv_after=max(0,sale*(1+shock/100)-other*(1+costshock/100))
    allow_base=max(0,cost-nrv_base)
    allow_after=max(0,cost-nrv_after)
    delta=allow_after-allow_base
    x,y,z=st.columns(3)
    x.metric('기준 순실현가능가치',f'{fmt(nrv_base)}억 원')
    y.metric('시나리오 순실현가능가치',f'{fmt(nrv_after)}억 원')
    z.metric('추가 평가손실(+)/환입(-)',f'{fmt(delta)}억 원')
    st.caption('기준 평가충당금과 시나리오 평가충당금의 차이를 계산했습니다. 환입은 기존 평가손실 범위를 초과하지 않도록 제한됩니다.')
    cf=pd.DataFrame({'항목':['취득원가','기준 순실현가능가치','시나리오 순실현가능가치','기준 평가충당금','시나리오 평가충당금'],
                     '금액_억':[cost,nrv_base,nrv_after,allow_base,allow_after]})
    fig=px.bar(cf,x='항목',y='금액_억',text='금액_억',title='가상 재고 평가 결과')
    fig.update_traces(texttemplate='%{text:,.1f}',textposition='outside')
    fig.update_layout(yaxis_title='억 원',xaxis_title='',height=380)
    st.plotly_chart(fig,use_container_width=True)
    st.info('실제 회계처리에서는 원재료가 사용될 완제품의 예상 판매수익성과 원가 회수가능성도 함께 고려합니다. 이 계산기는 가상 완제품/판매가능 재고의 단순화된 평가 사례입니다.')
    st.markdown('#### 제품 판매가격 하락률별 민감도')
    shocks=[-30,-25,-20,-15,-10,-5,0,5,10]
    rows=[]
    for s in shocks:
        nv=max(0,sale*(1+s/100)-other*(1+costshock/100))
        rows.append({'제품 판매가격 변동률':f'{s:+d}%','순실현가능가치_억':nv,'평가충당금_억':max(0,cost-nv),'기준 대비 추가손실_억':max(0,cost-nv)-allow_base})
    st.dataframe(pd.DataFrame(rows).round(1),hide_index=True,use_container_width=True)

with tabs[4]:
    st.subheader('분석 결과와 재무적 시사점')
    if qrow['평가손실_분기']>0:
        st.error(f'{chosen}: 재고평가손실 {fmt(qrow["평가손실_분기"],2)}억 원 발생 — 순실현가능가치 산정 근거와 제품별 가격 변동 점검 필요')
    else:
        st.success(f'{chosen}: 재고평가손실 {fmt(abs(qrow["평가손실_분기"]),2)}억 원 환입 — 과거 평가손실 환입 근거 확인 필요')
    st.markdown("""
**① 회계·결산 관리**  재고의 취득원가와 순실현가능가치를 비교하고, 제품가격 하락 및 추가 가공·판매비용의 반영이 적절한지 점검합니다.

**② 재고·운전자본 관리**  재고 장부금액 증가가 물량·가격·구성 중 어느 요인에서 발생했는지 파악하고, 과도한 자금 묶임을 점검합니다.

**③ 수익성 분석**  평가손익이 영업이익에 미치는 영향을 구분하되, IR에서 말하는 넓은 의미의 재고 관련 손익과 혼동하지 않습니다.

**④ 위험 시나리오**  국제유가와 제품가격이 동반 하락하거나 제품가격이 상대적으로 더 빠르게 하락할 때, 보유 재고의 원가 회수가능성을 재검토합니다.
""")
    st.markdown('#### 데이터 해석 원칙')
    st.info('**관측**: 공시된 평가손익·재고금액 및 국제가격 변동.  **가설**: 가격 하락과 원가 회수가능성 악화가 평가손실에 기여했을 가능성.  **미확인**: 실제 재고별 취득원가·판매조건·생산수율·헤지 효과에 따른 정확한 원인별 기여도.')
    st.markdown('#### 데이터 출처')
    st.write('S-OIL 연결 분기·반기·사업보고서(2025~2026), 한국석유공사 오피넷 국제 원유 및 석유제품 가격 다운로드 자료(2024~2026).')
    st.caption('이 프로그램은 교육·직무 탐구용 분석이며 투자판단 또는 실제 회계감사를 위한 산출물이 아닙니다.')

st.divider()
st.caption('S-OIL Inventory Risk Analyzer · 공시 사실과 가상 시뮬레이션을 구분하는 재무분석 도구')
