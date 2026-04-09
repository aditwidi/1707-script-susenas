* Encoding: UTF-8.
*bekerja di data 72_ssn_202403_kp43.sav.
*mencari kemiskinan kabupaten/kota dengan faktor koreksi (jika sudah diketahui).
*memasukkan variabel faktor koreksi.
compute fk=0.
IF r102=7 fk=0.86987940631734.
EXECUTE.

*memasukkan variabel garis kemiskinan kabupaten/kota.

compute gkkako=0.
IF r102=7 gkkako=532432.
EXECUTE.

*hitung kapitafk dengan faktor koreksi.
compute kapitafk=kapita*fk.
execute.

*memberikan flag rumah tangga miskin.
compute mkako=0.
IF kapitafk<gkkako mkako=1.
VARIABLE LABELS  mkako 'miskin kabupaten/kota'.
value labels mkako 
1'1 Miskin' 
0'0 Tidak Miskin'.
EXECUTE.

*menghitung indeks kedalam kemiskinan P1

compute p1kako=0.
if mkako=1 p1kako=((gkkako-kapitafk)/gkkako)*100.
VARIABLE LABELS p1kako 'Indeks kedalaman kemiskinan kabupaten/kota'.
execute.

*menghitung indeks keparahan kemiskinan P2

compute p2kako=0.
if mkako=1 p2kako=((gkkako-kapitafk)/gkkako)*((gkkako-kapitafk)/gkkako)*100.
VARIABLE LABELS p2kako 'Indeks keparahan kemiskinan kabupaten/kota'.
execute.

WEIGHT by weind.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 mkako p1kako p2kako DISPLAY=LABEL
  /TABLE r102 [C] BY mkako [C][COUNT F40.0, ROWPCT.COUNT PCT40.2] + p1kako [S][MEAN F40.3] + p2kako [S][MEAN F40.3]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE.

*kapita dibagi menjadi 10 ntile.
RANK VARIABLES=kapita (A)
  /NTILES(10)
  /PRINT=NO
  /TIES=MEAN.
*di bagian variable view kemudian merubah name variabel dari kapita menjadi Nkapitakako.

*memberikan flag pengeluaran perkapita berdasarkan world bank menurut wilayah perkotaan dan perdesaan.
RECODE Nkapita (1 thru 4=1) (5 thru 8=2) (9 thru 10=3) INTO Nkapitakakowb.
VARIABLE LABELS  Nkapitakakowb 'pengelompokkan 40-40-20 berdasarkan kabupaten/kota'.
value labels Nkapitakakowb
1'1 40% terbawah'
2'2 40% menengah'
3'3 20% teratas'.
EXECUTE.

*menghitung share pengeluaran makanan terhadap total pengeluran.
COMPUTE foodkapita = food/r301.
EXECUTE.

COMPUTE nonfoodkapita = nonfood/r301.
EXECUTE.

COMPUTE sharefoodkapita = foodkapita/kapita*100.
COMPUTE sharenonfoodkapita = nonfoodkapita/kapita*100.
EXECUTE.

WEIGHT BY WEIND.

*menampilkan tabel share makanan menurut status miskin.
* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 mkako sharefoodkapita DISPLAY=LABEL
  /TABLE r102 [C] BY mkako [C] > sharefoodkapita [MEAN F40.2]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE.


*tabel 1 dikumpulkan dari data yang sudah rilis.
*Bekerja di file 72_ssn_202403_kor_ind.sav.
*Tabel 2.
use all.
WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > r405 [C][COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE.

*tabel 3-4-5.

use all.
WEIGHT BY fwt.

*recode kelompok umur.
RECODE r407 (0 thru 14=1) (15 thru 64=2) (65 thru Highest=3) INTO kelum1.
VARIABLE LABELS  kelum1 'kelompok umur produktif'.
value labels kelum1
1'0-14'
2'15-64'
3'65+'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 kelum1 r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > kelum1 [C][COUNT F40.0, COLPCT.COUNT PCT40.2] BY r405 [C] > mkako [C]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=kelum1 mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 6-7-8.

USE ALL.
COMPUTE filter_$=(r407>=10).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

*recode kelompok umur.
RECODE r407 (10 thru 18=1) (19 thru Highest=2) INTO kelum2.
VARIABLE LABELS  kelum2 'kelompok umur'.
value labels kelum2
1'10-18'
2'>18'.
EXECUTE.

* Custom Tables. 
CTABLES 
  /VLABELS VARIABLES=r102 r404 kelum2 r405 mkako DISPLAY=LABEL 
  /TABLE r102 [C] > r404 [C] > kelum2 [C][COUNT F40.0, COLPCT.COUNT PCT40.2] BY r405 [C] > mkako [C] 
  /CATEGORIES VARIABLES=r102 r404 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER 
  /CATEGORIES VARIABLES=kelum2 mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER 
  /CRITERIA CILEVEL=95. 

*Tabel 9.

USE ALL.
COMPUTE filter_$=(r403=1).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > r405 [C][COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 10-11-12.

USE ALL.
COMPUTE filter_$=(r403=1).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 r404 r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > r404 [COUNT F40.0, COLPCT.COUNT PCT40.2] BY r405 [C] > mkako [C]
  /CATEGORIES VARIABLES=r102 r404 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 13-14-15.

USE ALL.
COMPUTE filter_$=(r403=1 & r407>=15).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

*recode kelompok umur.
RECODE r407 (15 thru 24=1) (25 thru 44=2) (45 thru 64=3) (65 thru Highest=4) INTO kelum3.
VARIABLE LABELS  kelum3 'kelompok umur'.
value labels kelum3
1'15-24'
2'25-44'
3'45-64'
4'65+'.
EXECUTE.

* Custom Tables. 
CTABLES 
  /VLABELS VARIABLES=r102 kelum3 r405 mkako DISPLAY=LABEL 
  /TABLE r102 > kelum3 [COUNT F40.0, COLPCT.COUNT PCT40.2] BY r405 > mkako 
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER 
  /CATEGORIES VARIABLES=kelum3 mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER 
  /CRITERIA CILEVEL=95.

*tabel 16.

USE ALL.
COMPUTE filter_$=(r403=1).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

*recode jumlah ART.
RECODE r301 (1 thru 3=1) (4 thru 6=2) (7 thru Highest=3) INTO keljart.
VARIABLE LABELS  keljart 'kelompok jumlah ART'.
value labels keljart
1'1-3'
2'4-6'
3'>=7'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 keljart mkako DISPLAY=LABEL
  /TABLE r102 > keljart [COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=keljart mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.


*Tabel 17-18-19.

USE ALL.
COMPUTE filter_$=(r407>=5 & r407<=24).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 r610 r405 mkako DISPLAY=LABEL
  /TABLE r102 > r610 [COUNT F40.0, COLPCT.COUNT PCT40.2] BY r405 > mkako
  /CATEGORIES VARIABLES=r102 r610 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*Tabel 20-21-22.

USE ALL.
COMPUTE filter_$=(r407>=15).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

*recode ijasah tertinggi.
RECODE r614 (0=1) (25=1) (1 thru 5=2) (6 thru 10=3) (11 thru 17=4) (18 thru 24=5) INTO kelijasah.
VARIABLE LABELS  kelijasah 'ijasah tertinggi yang dimiliki'.
value labels kelijasah
1'Tidak Tamat SD/Tidak atau belum pernah sekolah'
2'SD Sederajat'
3'SMP Sederajat'
4"SMU Sederajat"
5'Perguruan Tinggi'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 kelijasah r405 mkako DISPLAY=LABEL
  /TABLE r102 > kelijasah [COUNT F40.0, COLPCT.COUNT PCT40.2] BY r405 > mkako
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=kelijasah mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 23-24-25.

USE ALL.
COMPUTE filter_$=(r407>=7 & r407<=18).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

*recode kelompok umur.
RECODE r407 (7 thru 12=1) (13 thru 15=2) (16 thru 18=3) INTO kelum4.
VARIABLE LABELS  kelum4 'kelompok umur'.
value labels kelum4
1'7-12'
2'13-15'
3'16-18'.
EXECUTE.

compute aps=0.
IF r610=2 aps=100.
VARIABLE LABELS  aps 'angka partisipasi sekolah'.
value labels aps
100'sekolah'.
execute.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 kelum4 r405 mkako aps DISPLAY=LABEL
  /TABLE r102 [C] > kelum4 [C] BY r405 [C] > mkako [C] > aps [S][MEAN]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=kelum4 ORDER=A KEY=VALUE EMPTY=INCLUDE
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 26-27-28.

USE ALL.
COMPUTE filter_$=(r407>=15 & r407<=55).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

*bisa baca dan tulis.

compute amh=0.
IF r607=1 or r608=1 or r609=1 amh=100.
IF r607=5 and r608=5 and r609=5 amh=0.
VARIABLE LABELS  amh 'angka melek huruf'.
value labels amh
100'Bisa baca tulis'
0'Tidak bisa baca tulis'.
EXECUTE.

compute amh1=0.
IF r407>=15 & r407<=24 amh1=1.
VARIABLE LABELS  amh1 'kelompok umur 15-24'.
value labels amh1
1'15-24'
0'Lainnya'.
EXECUTE.

compute amh2=0.
IF r407>=15 & r407<=55 amh2=1.
VARIABLE LABELS  amh2 'kelompok umur 15-55'.
value labels amh2
1'15-55'
0'Lainnya'.
EXECUTE.

*select AMH 15-24.
USE ALL.
COMPUTE filter_$=((r407>=15 & r407<=24) & amh1=1).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 amh1 r405 mkako amh DISPLAY=LABEL
  /TABLE r102 [C] > amh1 [C] BY r405 [C] > mkako [C] > amh [S][MEAN]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=amh1 ORDER=A KEY=VALUE EMPTY=INCLUDE
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.


*select AMH 15-55.
USE ALL.
COMPUTE filter_$=((r407>=15 & r407<=55) & amh2=1).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 amh2 r405 mkako amh DISPLAY=LABEL
  /TABLE r102 [C] > amh2 [C] BY r405 [C] > mkako [C] > amh [S][MEAN]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=amh2 ORDER=A KEY=VALUE EMPTY=INCLUDE
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 29-30-31.

USE ALL.
COMPUTE filter_$=(r403=1).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 kelijasah r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > kelijasah BY r405 [C] > mkako [C][COUNT F40.0, COLPCT.COUNT PCT40.2]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=kelijasah mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 32-33-34.

USE ALL.
COMPUTE filter_$=(r407>=15).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

compute tkerja=0.
if r706=0 tkerja=100.
VARIABLE LABELS  tkerja 'tidak kerja'.
value labels tkerja
100'tidak bekerja'
0'Bekerja'.
EXECUTE.

RECODE r707 (3 thru 4=1) (1 thru 2=2) (5 thru 6=2) INTO statuskerja.
VARIABLE LABELS  statuskerja 'formal-informal'.
value labels statuskerja
1'Formal'
2'Informal'.
EXECUTE.

compute kformal=0.
if statuskerja=1 kformal=100.
VARIABLE LABELS  kformal 'kerja di sektor formal'.
value labels kformal
100'Kerja di sektor Formal'
0'Bekerja'.
EXECUTE.

compute kinformal=0.
if statuskerja=2 kinformal=100.
VARIABLE LABELS  kinformal 'kerja di sektor informal'.
value labels kinformal
100'Kerja di sektor informal'
0'Bekerja'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 tkerja kformal kinformal r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > (tkerja [S][MEAN] + kformal [S][MEAN] + kinformal [S][MEAN]) BY r405 [C] >
    mkako [C]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 35-36-37.

USE ALL.
COMPUTE filter_$=(r407>=15).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

compute tkerja=0.
if r706=0 tkerja=100.
VARIABLE LABELS  tkerja 'tidak kerja'.
value labels tkerja
100'tidak bekerja'
0'Bekerja'.
EXECUTE.

RECODE r706 (1 thru 6=1) (7 thru 26=2) INTO sektorkerja.
VARIABLE LABELS  sektorkerja 'pertanian non pertanian'.
value labels sektorkerja
1'Pertanian'
2'Non Pertanian'.
EXECUTE.

compute ktani=0.
if sektorkerja=1 ktani=100.
VARIABLE LABELS  ktani 'kerja di sektor pertanian'.
value labels ktani
100'Kerja di sektor pertanian'.
EXECUTE.

compute kntani=0.
if sektorkerja=2 kntani=100.
VARIABLE LABELS  kntani 'kerja di sektor non pertanian'.
value labels kntani
100'Kerja di sektor non pertanian'.
EXECUTE.


* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 tkerja ktani kntani r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > (tkerja [S][MEAN] + ktani [S][MEAN] + kntani [S][MEAN]) BY r405 [C] > mkako [C]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.


*tabel 38-39-40.

use all.
WEIGHT BY fwt.


*kepemilikan JKN.

COMPUTE milikjkn=0.
IF  (r1101_a="A" | r1101_b = "B" | r1101_c="C" | r1101_d="D" | r1101_e="E") milikJKN=1.
IF  (r1101_x="X") milikJKN=0.
value labels milikJKN
1'Memiliki'
0'Tidak Memiliki'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 milikjkn r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > milikjkn [C][COUNT F40.0, COLPCT.COUNT PCT40.2] BY r405 [C] > mkako [C]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=milikjkn mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 41-42-43.

USE ALL.
COMPUTE filter_$=(r407>=5).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.
WEIGHT BY fwt.

RECODE r1207 (1=1) (2=2) (5=3) (8=3) INTO rokok.
VARIABLE LABELS  rokok 'merokok tembakau'.
value labels rokok
1'Ya, setiap hari'
2'Ya, tidak setiap hari'
3'Tidak/Tidak Tahu'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 rokok r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > rokok [COUNT F40.0, COLPCT.COUNT PCT40.2] BY r405 [C] > mkako [C]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=rokok mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*buka file 72_ssn_202403_kor_rt.sav.
*tabel 44.
use all.
WEIGHT BY fwt.

*sanitasi layak.
*1.Fasilitas buang air besar : sendiri & bersama ART tertentu & komunal dan jenis kloset: leher angsa dan tempat pembuangan akhir tinja: tangki septik, IPAL.
*2.Fasilitas buang air besar : sendiri & bersama ART tertentu & komunal dan  n jenis kloset: leher angsa dan tempat pembuangan akhir tinja: lubang tanah dan wilayah perdesaan.

compute sal=0.
If (R1809A<=3 and R1809B=1 and R1809C<=2) sal=100. 
If (R1809A<=3 and R1809B=1 and R1809C=4 and R105=2) sal=100.
Variable Labels sal "Akses Terhadap Layanan Sanitasi Layak".
Value Labels sal
0 'Tidak punya akses sanitasi layak'  
100 'punya akses sanitasi layak'.
EXECUTE.

*air minum layak.

Compute airmlayak=0.
IF  (r1810a=3 | r1810a=4 | r1810a=5 | r1810a=7 | r1810a=10) & (r1814a=3 | r1814a=4 | r1814a=5 | r1814a=7 | r1814a=10) airmlayak=100.
IF  (r1810a=3 | r1810a=4 | r1810a=5 | r1810a=7 | r1810a=10) & (r1814a=1 | r1814a=2) airmlayak=100.
IF  (r1810a=3 | r1810a=4 | r1810a=5 | r1810a=7 | r1810a=10) & (r1814a=6 | r1814a=8 | r1814a=9 | r1814a=11) airmlayak=100.
IF  (r1810a=1 | r1810a=2) & (r1814a=3 | r1814a=4 | r1814a=5 | r1814a=7 | r1814a=10) airmlayak=100.
IF  (r1810a=1 | r1810a=2) & (r1814a=1 | r1814a=2) airmlayak=0.
IF  (r1810a=1 | r1810a=2) & (r1814a=6 | r1814a=8 | r1814a=9 | r1814a=11) airmlayak=0.
IF  (r1810a=6 | r1810a=8 | r1810a=9 | r1810a=11) & (r1814a=3 | r1814a=4 | r1814a=5 | r1814a=7 | r1814a=10) airmlayak=0.
IF  (r1810a=6 | r1810a=8 | r1810a=9 | r1810a=11) & (r1814a=1 | r1814a=2) airmlayak=0.
IF  (r1810a=6 | r1810a=8 | r1810a=9 | r1810a=11) & (r1814a=6 | r1814a=8 | r1814a=9 | r1814a=11) airmlayak=0.
VARIABLE LABELS  airmlayak 'air minum layak'.
Value Labels airmlayak
0 'Tidak punya akses air minum layak'  
100 'punya akses air minum layak'.
EXECUTE.

*sintak air minum bersih :  Terdiri dari air kemasan, air isi ulang, air ledeng dan [(sumur bor/pompa, sumur terlindung serta mata air terlindung) dengan jarak ke tempat penampungan akhir tinja ? 10 m]

Compute sab=0.
if r1810a=1 sab=100.
if r1810a=2 sab=100.
if r1810a=3 sab=100.
if (r1810a=4 or r1810a=5 or r1810a=7) and r1810c=2 sab=100.
VARIABLE LABELS  sab 'air minum bersih'.
Value Labels sab
0 'Tidak punya akses air minum bersih'  
100 'Punya akses air minum bersih'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 airmlayak sab sal mkako DISPLAY=LABEL
  /TABLE r102 [C] > (airmlayak [S][MEAN] + sab [S][MEAN] + sal [S][MEAN]) BY mkako [C]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.



*tabel 45.

use all.
WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 r1802 mkako DISPLAY=LABEL
  /TABLE r102 [C] > r1802 [COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 r1802 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 46.

use all.
WEIGHT BY fwt.

*luas per kapita.
COMPUTE lkapita=r1804/r301.
VARIABLE LABELS  lkapita 'luas lantai per kapita'.
EXECUTE.

RECODE lkapita (0 thru 7.2=1) (ELSE=2) INTO klkapita.
VARIABLE LABELS  klkapita 'kelompok luas lantai per kapita'.
value labels klkapita
1'<=7,2 m2'
2'>7,2 m2'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 klkapita mkako DISPLAY=LABEL
  /TABLE r102 [C] > klkapita [COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=klkapita mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 47.

use all.
WEIGHT BY fwt.

*recode atap rumah terluas.
RECODE r1806a (1=1) (2=1) (3=2) (4=3) (else=4) INTO katap.
VARIABLE LABELS  katap 'jenis atas rumah terluas'.
value labels katap
1'Beton/Genteng'
2'Seng'
3'Asbes'
4'Bambu/kayu/sirap/jerami/ijuk/daun-daunan/rumbia/lainnya'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 katap mkako DISPLAY=LABEL
  /TABLE r102 [C] > katap [C][COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=katap mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 48.

use all.
WEIGHT BY fwt.

*recode dinding terluas.
RECODE r1807 (1=1) (2=2) (3=3) (else=4) INTO kdinding.
VARIABLE LABELS  kdinding 'jenis dinding rumah terluas'.
value labels kdinding
1'Tembok'
2'Plesteran Anyaman Bambu/Kawat'
3'Kayu/Papan'
4'Lainnya (Anyaman Bambu/batang kayu/bamboo/lainnya)'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 kdinding mkako DISPLAY=LABEL
  /TABLE r102 [C] > kdinding [COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=kdinding mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 49.

use all.
WEIGHT BY fwt.

*recode atas lantai terluas.
RECODE r1808 (1 thru 3=1) (4=2) (5=3) (else=4) INTO klantai.
VARIABLE LABELS  klantai 'jenis lantai rumah terluas'.
value labels klantai
1'Marmer/Granit/keramik/parket/vinil/karpet'
2'Ubin/Tegel/Teraso'
3'Kayu/papan'
4'Lainnya (Semen/Bata Merah/Bambu/Tanah/Lainnya)'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 klantai mkako DISPLAY=LABEL
  /TABLE r102 [C] > klantai [COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=klantai mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 50.

use all.
WEIGHT BY fwt.

*recode sumber penerangan.
RECODE r1816 (1 thru 2=1) (3=2) (4=3) INTO klistrik.
VARIABLE LABELS  klistrik 'sumber penerangan utama'.
value labels klistrik
1'Listrik PLN'
2'Listrik Non PLN'
3'Bukan listrik'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 klistrik mkako DISPLAY=LABEL
  /TABLE r102 > klistrik [COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=klistrik mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 51-52-53.
*buka file 72_ssn_202403_kor_ind.sav.

use all.
WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 sharefoodkapita sharenonfoodkapita r405 mkako DISPLAY=LABEL
  /TABLE r102 [C] > (sharefoodkapita [S][MEAN] + sharenonfoodkapita [S][MEAN]) BY r405 [C] > mkako
    [C]
  /CATEGORIES VARIABLES=r102 r405 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 54.
*buka file 72_ssn_202403_kor_kp41.sav.

USE ALL.
COMPUTE filter_$=(klp<>0).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY weind.

*menghitung konsumsi per kapita per bulan.
compute kons_art=(b41k10*(30/7))/r301.
execute.

*aggregate kelompok makanan.
AGGREGATE
  /OUTFILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\1707_ssn_202403_aggr_klpfood_kabko.sav'
  /BREAK=r101 r102 mkako klp 
  /kons_art_sum=SUM(kons_art).

*buka file 72_ssn_202403_kor_kp43.sav.

use all.
WEIGHT BY weind.

*aggregate jumlah penduduk.
AGGREGATE
  /OUTFILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\1707_ssn_202403_aggr_jmlart_kabko.sav'
  /BREAK=r101 r102 mkako
  /r301_n=N(r301).

*buka file 72_ssn_202403_aggr_klpfood_kabko.sav kemudian merge dengan 72_ssn_202403_aggr_jmlart_kabko.sav.
*hasil gabungan di file 72_ssn_202403_aggr_klpfood_kabko.sav.

COMPUTE ratakonskapita=kons_art_sum / r301_n.
VARIABLE LABELS  ratakonskapita 'rata-rata konsumsi makanan per kapita per bulan'.
EXECUTE.


* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 klp mkako ratakonskapita DISPLAY=LABEL
  /TABLE r102 > klp BY mkako > ratakonskapita [MEAN F40.0, COLPCT.SUM PCT40.2] 
  /CATEGORIES VARIABLES=r102 klp ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE.

*buka file 72_ssn_202403_kor_kp43.sav.
use all.
WEIGHT BY weind.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 foodkapita mkako DISPLAY=LABEL
  /TABLE r102 [C] BY foodkapita [S][MEAN] > mkako [C]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 55.
*buka file 72_ssn_202403_kor_kp42.sav.

USE ALL.
COMPUTE filter_$=(klp<>0 & klp<>1 & klp<>2).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY weind.

*menghitung konsumsi per kapita per bulan.
compute kons_art=(sebulan)/r301.
execute.

*aggregate kelompok non makanan.
AGGREGATE
  /OUTFILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\1707_ssn_202403_aggr_klpnonfood_kabko.sav'
  /BREAK=r101 r102 mkako klp 
  /kons_art_sum=SUM(kons_art).

*buka file 72_ssn_202403_kor_kp43.sav.

use all.
WEIGHT BY weind.

*aggregate jumlah penduduk.
AGGREGATE
  /OUTFILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\1707_ssn_202403_aggr_jmlart_kabko2.sav'
  /BREAK=r101 r102 mkako
  /r301_n=N(r301).

*buka file 72_ssn_202403_aggr_klpnonfood_kabko.sav kemudian merge dengan 72_ssn_202403_aggr_jmlart_kabko2.sav.
*hasil gabungan di file 72_ssn_202403_aggr_klpnonfood_kabko.sav.

COMPUTE ratakonskapita=kons_art_sum / r301_n.
VARIABLE LABELS  ratakonskapita 'rata-rata konsumsi non makanan per kapita per bulan'.
EXECUTE.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 klp mkako ratakonskapita DISPLAY=LABEL
  /TABLE r102 [C] > klp [C] BY mkako [C] > ratakonskapita [S][MEAN F40.0, COLPCT.SUM PCT40.2]
  /CATEGORIES VARIABLES=r102 klp ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE.

*buka file 72_ssn_202403_kor_kp43.sav.

use all.
WEIGHT BY weind.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 nonfoodkapita mkako DISPLAY=LABEL
  /TABLE r102 [C] > nonfoodkapita [S][MEAN] BY mkako [C]
  /CATEGORIES VARIABLES=r102 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE.

*tabel 56.
*buka file 72_ssn_202403_kor_rt.sav.

use all.
WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 r2203 mkako DISPLAY=LABEL
  /TABLE r102 [C] > r2203 [C][COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 r2203 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.


*tabel 57.

use all.
WEIGHT BY fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 r2207 mkako DISPLAY=LABEL
  /TABLE r102 [C] > r2207 [COUNT F40.0, COLPCT.COUNT PCT40.2] BY mkako [C]
  /CATEGORIES VARIABLES=r102 r2207 ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER
  /CATEGORIES VARIABLES=mkako ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

*tabel 58-59.

*buka file 72_ssn_202403_kor_ind.sav.
*buat CSplan.

* Analysis Preparation Wizard (dibuat manual).
CSPLAN ANALYSIS
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_1707_ssn_202403_kor_ind.csaplan'
  /PLANVARS ANALYSISWEIGHT=fwt
  /SRSESTIMATOR TYPE=WR
  /PRINT PLAN
  /DESIGN STRATA=strata CLUSTER=psu
  /ESTIMATOR TYPE=WR.

USE ALL.
COMPUTE filter_$=(r407>=7 & r407<=18).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.
* Complex Samples Descriptives.
CSDESCRIPTIVES
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_1707_ssn_202403_kor_ind1.csaplan'
  /SUMMARY VARIABLES=aps
  /SUBPOP TABLE=mkako BY r102 BY kelum4 DISPLAY=LAYERED
  /MEAN
  /STATISTICS SE CV CIN(95)
  /MISSING SCOPE=ANALYSIS CLASSMISSING=EXCLUDE.


*tabel 60-61.

USE ALL.
COMPUTE filter_$=(r407>=15).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.

* Complex Samples Descriptives.
CSDESCRIPTIVES
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_1707_ssn_202403_kor_ind1.csaplan'
  /SUMMARY VARIABLES=tkerja ktani kntani
  /SUBPOP TABLE=mkako BY r102 DISPLAY=LAYERED
  /MEAN
  /STATISTICS SE CV CIN(95)
  /MISSING SCOPE=ANALYSIS CLASSMISSING=EXCLUDE.


*tabel 62-63.

USE ALL.
COMPUTE filter_$=(r407>=15).
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.

WEIGHT BY fwt.
* Complex Samples Descriptives.
CSDESCRIPTIVES
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_1707_ssn_202403_kor_ind1.csaplan'
  /SUMMARY VARIABLES=tkerja kformal kinformal
  /SUBPOP TABLE=mkako BY r102 DISPLAY=LAYERED
  /MEAN
  /STATISTICS SE CV CIN(95)
  /MISSING SCOPE=ANALYSIS CLASSMISSING=EXCLUDE.

*tabel 64-65.

*buka file 72_ssn_202403_kor_rt.sav.
*buat CSplan.
CSPLAN ANALYSIS
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_1707_ssn_202403_kor_rt.csaplan'
  /PLANVARS ANALYSISWEIGHT=weind
  /SRSESTIMATOR TYPE=WR
  /PRINT PLAN
  /DESIGN STRATA=r105 CLUSTER=psu
  /ESTIMATOR TYPE=WR.

use all.
weight by fwt.

* Complex Samples Descriptives.
CSDESCRIPTIVES
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_1707_ssn_202403_kor_rt.csaplan'
  /SUMMARY VARIABLES=airmlayak sab sal
  /SUBPOP TABLE=mkako BY r102 DISPLAY=LAYERED
  /MEAN
  /STATISTICS SE CV CIN(95)
  /MISSING SCOPE=ANALYSIS CLASSMISSING=EXCLUDE.

*tabel 66-69.

*buka file 72_ssn_202403_kor_kp43.sav.

USE ALL.

*buat dummy persentase miskin kabupaten/kota.

compute dmkako=0.
if mkako=1 dmkako=100.
VARIABLE LABELS  dmkako 'Persentase Penduduk Miskin'.
EXECUTE.

*buat CS Plan individu.

* Analysis Preparation Wizard(buat sendiri).
CSPLAN ANALYSIS
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_72_ssn_202403_kor_kp43_ind1.csaplan'
  /PLANVARS ANALYSISWEIGHT=weind
  /SRSESTIMATOR TYPE=WR
  /PRINT PLAN
  /DESIGN CLUSTER=psu
  /ESTIMATOR TYPE=WR.

use all.
weight by weind.

*buat RSE jumlah penduduk miskin.
* Complex Samples Descriptives.
CSDESCRIPTIVES
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_72_ssn_202403_kor_kp43_ind12.csaplan'
  /SUMMARY VARIABLES=mkako
  /SUBPOP TABLE=r102 DISPLAY=LAYERED
  /SUM
  /STATISTICS SE CV CIN(95)
  /MISSING SCOPE=ANALYSIS CLASSMISSING=EXCLUDE.

*buat RSE P0, P1 dan P2.

* Complex Samples Descriptives.
CSDESCRIPTIVES
  /PLAN FILE='D:\bps_lebong\2025\Publikasi\Profil Kemiskinan 2024\Trial and error\CSPlan_72_ssn_202403_kor_kp43_ind12.csaplan'
  /SUMMARY VARIABLES=dmkako p1kako p2kako
  /SUBPOP TABLE=r102 DISPLAY=LAYERED
  /MEAN
  /STATISTICS SE CV CIN(95)
  /MISSING SCOPE=ANALYSIS CLASSMISSING=EXCLUDE.

*buka file 72_ssn_202403_kor_rt.sav.
*grafik 1.

use all.
weight by fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 Nkapita airmlayak DISPLAY=LABEL
  /TABLE r102 [C] BY Nkapita [C] > airmlayak [S][MEAN]
  /CATEGORIES VARIABLES=r102 Nkapita ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER.



*grafik 2.

use all.
weight by fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 Nkapita sab DISPLAY=LABEL
  /TABLE r102 [C] BY Nkapita [C] > sab [S][MEAN]
  /CATEGORIES VARIABLES=r102 Nkapita ORDER=A KEY=VALUE EMPTY=EXCLUDE TOTAL=YES POSITION=AFTER.



*grafik 3.

use all.
weight by fwt.

* Custom Tables.
CTABLES
  /VLABELS VARIABLES=r102 Nkapita sal DISPLAY=LABEL
  /TABLE r102 [C][COUNT F40.0, ROWPCT.COUNT PCT40.1] BY Nkapita [C] > sal
  /CATEGORIES VARIABLES=r102 Nkapita ORDER=A KEY=VALUE EMPTY=EXCLUDE
  /CATEGORIES VARIABLES=sal ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER.

