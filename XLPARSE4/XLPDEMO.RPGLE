000100000000      ** This program demonstrates parsing a spreadsheet with the
000200000000      **  XLPARSER4 service program.
000300000000      **                                Scott Klement, January 8, 2004
000400000000      **
000500000000      ** To compile:
000600000000      **    (first follow the instructions to compile XLPARSER4)
000700000000      **>   CRTBNDRPG XLPDEMO SRCFILE(QRPGLESRC) DBGVIEW(*LIST)
000800000000      **
000900000000     H DFTACTGRP(*NO) BNDDIR('XLPARSE') ACTGRP(*NEW)
001000000000
001100000000     FQSYSPRT   O    F   80        PRINTER OFLIND(*INOF)
001200000000
001300000000      /copy XLPARSE_H
001400000000
001500000000     D XLPDEMO         PR                  ExtPgm('XLPDEMO')
001600000000     D   peXSSF                       1n   const options(*nopass)
001700000000     D XLPDEMO         PI
001800000000     D   peXSSF                       1n   const options(*nopass)
001900000000
002000000000     D IFSDIR          C                   '/tmp'
002100000000
002200000000     D charcell        PR             1N
002300000000     D   Sheet                     1024A   varying const
002400000000     D   Row                         10I 0 value
002500000000     D   Column                       5I 0 value
002600000000     D   Value                    32767A   varying const
002700000000
002800000000     D numbercell      PR             1N
002900000000     D   Sheet                     1024A   varying const
003000000000     D   Row                         10I 0 value
003100000000     D   Column                       5I 0 value
003200000000     D   Value                        8F   value
003300000000
003400000000     D msg             s             50A
003500000000     D wait            s              1A
003600000000
003700000000     D year            s              4P 0
003800000000     D part            s              8A
003900000000     D descr           s             28A
004000000000     D sales           s              7P 2 dim(4)
004100000000     D myfile          s           1024a   varying
004200000000
004300000000      /free
004400000000
004500000000         // FIXME: Change the path below to the place where you put
004600000000         //       xlpdemo.xls in your IFS!
004700000000
004800000000         if (%parms>=1 and peXSSF);
004900000000            myfile = IFSDIR + '/xlpdemo.xlsx';
005000000000         else;
005100000000            myfile = IFSDIR + '/xlpdemo.xls';
005200000000         endif;
005300000000
005400000000         if xlparse_workbook( myfile
005500000000                            : %paddr(NumberCell)
005600000000                            : %paddr(CharCell)  ) < 0;
005700000000             msg = 'Error occcured while parsing spreadsheet.';
005800000000             dsply msg ' ' wait;
005900000000         endif;
006000000000
006100000000         *inlr = *on;
006200000000         return;
006300000000
006400000000       /end-free
006500000000
006600000000     OQSYSPRT   E            Heading        1  3
006700000000     O                       *DATE         Y     10
006800000000     O                                           +1 'ACME Widget Sales for'
006900000000     O                       year                +1
007000000000     O                                           76 'Page'
007100000000     O                       page          Z     80
007200000000
007300000000     O          E            item
007400000000     O                       part                10
007500000000     O                       descr               +1
007600000000     O                       sales(1)      L     +1
007700000000     O                       sales(2)      L     +1
007800000000     O                       sales(3)      L     +1
007900000000     O                       sales(4)      L     +1
008000000000
008100000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
008200000000      *  This is called once for each character cell in the spreadsheet
008300000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
008400000000     P charcell        B
008500000000     D charcell        PI             1N
008600000000     D   Sheet                     1024A   varying const
008700000000     D   Row                         10I 0 value
008800000000     D   Column                       5I 0 value
008900000000     D   Value                    32767A   varying const
009000000000      /free
009100000000
009200000000          // We're only interested in the "part no" and "description"
009300000000          // columns which are in rows 5 - 20, columns A & B.
009400000000
009500000000          if (row>=4 and row<=19);
009600000000
009700000000             select;
009800000000             when column = 0;
009900000000               part = value;
010000000000             when column = 1;
010100000000               descr = value;
010200000000             endsl;
010300000000
010400000000          endif;
010500000000
010600000000          return *OFF;
010700000000
010800000000      /end-free
010900000000     P                 E
011000000000
011100000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
011200000000      *  This is called once for each numeric cell in the spreadsheet
011300000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
011400000000     P numbercell      B
011500000000     D numbercell      PI             1N
011600000000     D   Sheet                     1024A   varying const
011700000000     D   Row                         10I 0 value
011800000000     D   Column                       5I 0 value
011900000000     D   Value                        8F   value
012000000000
012100000000     D firstpage       s              1N   static inz(*ON)
012200000000     D qtr             s             10I 0
012300000000
012400000000      /free
012500000000
012600000000          // we're interested in getting the year from the first row.
012700000000
012800000000          if (row=0 and column=7);
012900000000              year = value;
013000000000          endif;
013100000000
013200000000          // We're also interested in the sales figures which are
013300000000          // in columns C-F and rows 5-20
013400000000
013500000000          if (row>=4 and row<=19);
013600000000
013700000000             if (column>=2 and column<=5);
013800000000
013900000000                // save sales figures into an array.
014000000000
014100000000                qtr = column - 1;
014200000000                sales(qtr) = value;
014300000000
014400000000                // if this is the last column, print this item, and
014500000000                //   clear out our variables.
014600000000
014700000000                if (column = 5);
014800000000
014900000000                    if (firstpage or *inof=*on);
015000000000                       except heading;
015100000000                       firstpage = *Off;
015200000000                    endif;
015300000000
015400000000                    except item;
015500000000
015600000000                    for qtr = 1 to %elem(sales);
015700000000                       sales(qtr) = 0;
015800000000                    endfor;
015900000000
016000000000                    part = *blanks;
016100000000                    descr = *blanks;
016200000000
016300000000                endif;
016400000000
016500000000             endif;
016600000000          endif;
016700000000
016800000000          return *off;
016900000000      /end-free
017000000000     P                 E
