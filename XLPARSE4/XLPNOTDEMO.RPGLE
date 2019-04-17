000100000000      ** This program demonstrates parsing a spreadsheet with the
000200000000      **  XLPARSER4 service program and uses the xlparse_notify()
000300000000      **  routine to determine where rows start/end.
000400000000      **                             Scott Klement, November 25, 2008
000500000000      **
000600000000      ** To compile:
000700000000      **    (first follow the instructions to compile XLPARSER4)
000800000000      **>   CRTBNDRPG XLPNOTDEMO SRCFILE(QRPGLESRC) DBGVIEW(*LIST)
000900000000      **
001000000000     H DFTACTGRP(*NO) BNDDIR('XLPARSE') ACTGRP('KLEMENT')
001100000000
001200000000     FQSYSPRT   O    F  132        PRINTER OFLIND(OverFlow)
001300000000
001400000000      /copy XLPARSE_H
001500000000
001600000000     D XLPNOTDEMO      PR                  ExtPgm('XLPNOTDEMO')
001700000000     D   peXSSF                       1n   const options(*nopass)
001800000000     D XLPNOTDEMO      PI
001900000000     D   peXSSF                       1n   const options(*nopass)
002000000000
002100000000     D IFSDIR          C                   '/tmp'
002200000000
002300000000     D clear_struct    PR
002400000000     D   Sheet                     1024A   varying const
002500000000     D   Row                         10I 0 value
002600000000
002700000000     D print_struct    PR
002800000000     D   Sheet                     1024A   varying const
002900000000     D   Row                         10I 0 value
003000000000
003100000000     D character       PR             1N
003200000000     D   Sheet                     1024A   varying const
003300000000     D   Row                         10I 0 value
003400000000     D   Column                       5I 0 value
003500000000     D   Value                    32767A   varying const
003600000000
003700000000     D numeric         PR             1N
003800000000     D   Sheet                     1024A   varying const
003900000000     D   Row                         10I 0 value
004000000000     D   Column                       5I 0 value
004100000000     D   Value                        8F   value
004200000000
004300000000     D mon             ds                  qualified
004400000000     D   part                        10a
004500000000     D   desc                        30a
004600000000     D   saleA                       11p 2 inz
004700000000     D   saleB                       11p 2 inz
004800000000     D   saleC                       11p 2 inz
004900000000     D   saleD                       11p 2 inz
005000000000
005100000000     D firstpage       s              1n   inz(*ON)
005200000000     D Month           s              9a
005300000000     D Year            s              4s 0
005400000000
005500000000     D filename        s            200a   varying
005600000000     D msg             s             50A
005700000000     D wait            s              1A
005800000000
005900000000      /free
006000000000
006100000000         xlparse_notify( %paddr(clear_struct)
006200000000                       : %paddr(print_struct) );
006300000000
006400000000         if (%parms>=1 and peXSSF);
006500000000            filename = IFSDIR + '/november_sales.xlsx';
006600000000         else;
006700000000            filename = IFSDIR + '/november_sales.xls';
006800000000         endif;
006900000000
007000000000         if xlparse_workbook( filename
007100000000                            : %paddr(Numeric)
007200000000                            : %paddr(Character)
007300000000                            : *NULL ) <> 0;
007400000000             msg = 'Error occcured. See job log.';
007500000000             dsply msg ' ' wait;
007600000000         endif;
007700000000
007800000000         // this disables notifications:
007900000000
008000000000         xlparse_notify( *null: *null );
008100000000
008200000000         *inlr = *on;
008300000000         return;
008400000000
008500000000       /end-free
008600000000
008700000000     OQSYSPRT   E            Heading        1  3
008800000000     O                       *DATE         Y     10
008900000000     O                                           +1 'ACME Widget Sales for'
009000000000     O                       Month               +1
009100000000     O                                           76 'Page'
009200000000     O                       page          Z     80
009300000000
009400000000     O          E            item
009500000000     O                       mon.part            10
009600000000     O                       mon.desc            +1
009700000000     O                       mon.saleA     L     +1
009800000000     O                       mon.saleB     L     +1
009900000000     O                       mon.saleC     L     +1
010000000000     O                       mon.saleD     L     +1
010100000000
010200000000
010300000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
010400000000      *  This is called once for each character cell in the spreadsheet
010500000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
010600000000     P Character       B
010700000000     D Character       PI             1N
010800000000     D   Sheet                     1024A   varying const
010900000000     D   Row                         10I 0 value
011000000000     D   Column                       5I 0 value
011100000000     D   Value                    32767A   varying const
011200000000      /free
011300000000          if (row = 1  and  column = 2);
011400000000             month = value;
011500000000          endif;
011600000000
011700000000          if (row > 4);
011800000000
011900000000             select;
012000000000             when column = 0;
012100000000               mon.part = value;
012200000000             when column = 1;
012300000000               mon.desc = value;
012400000000             endsl;
012500000000
012600000000          endif;
012700000000
012800000000          return *OFF;
012900000000
013000000000      /end-free
013100000000     P                 E
013200000000
013300000000
013400000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
013500000000      *  This is called once for each numeric cell in the spreadsheet
013600000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
013700000000     P Numeric         B
013800000000     D Numeric         PI             1N
013900000000     D   Sheet                     1024A   varying const
014000000000     D   Row                         10I 0 value
014100000000     D   Column                       5I 0 value
014200000000     D   Value                        8F   value
014300000000
014400000000      /free
014500000000          if (row = 1  and  column = 3);
014600000000             year = value;
014700000000          endif;
014800000000
014900000000
015000000000          // Sales figures are in columns C,D,E,F of each row
015100000000          // beyond row 4
015200000000
015300000000          if (row > 4);
015400000000
015500000000             select;
015600000000             when column=2;
015700000000                mon.saleA = value;
015800000000             when column=3;
015900000000                mon.saleB = value;
016000000000             when column=4;
016100000000                mon.saleC = value;
016200000000             when column=5;
016300000000                mon.saleD = value;
016400000000             endsl;
016500000000
016600000000          endif;
016700000000
016800000000          return *off;
016900000000      /end-free
017000000000     P                 E
017100000000
017200000000
017300000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
017400000000      *  This is called before each row that has data
017500000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
017600000000     P clear_struct    B
017700000000     D clear_Struct    PI
017800000000     D   Sheet                     1024A   varying const
017900000000     D   Row                         10I 0 value
018000000000      /free
018100000000         if Row = 1;
018200000000            Month = *blanks;
018300000000            year  = 0;
018400000000         endif;
018500000000
018600000000         if Row > 4;
018700000000            clear mon;
018800000000         endif;
018900000000      /end-free
019000000000     P                 E
019100000000
019200000000
019300000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
019400000000      *  This is called after each row that has data
019500000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
019600000000     P print_struct    B
019700000000     D print_Struct    PI
019800000000     D   Sheet                     1024A   varying const
019900000000     D   Row                         10I 0 value
020000000000
020100000000      /free
020200000000         if row > 4;
020300000000             if (firstpage or overflow);
020400000000                except Heading;
020500000000                firstpage = *off;
020600000000                overflow  = *off;
020700000000             endif;
020800000000
020900000000             except item;
021000000000         endif;
021100000000      /end-free
021200000000     P                 E
