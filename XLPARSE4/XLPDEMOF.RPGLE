000100000000      ** This program demonstrates parsing a spreadsheet that has
000200000000      ** formulas with the XLPARSER4 program.
000300000000      **                                Scott Klement, July 12, 2007
000400000000      **
000500000000      ** WARNING: Formula support should be considered "experimental"
000600000000      **          at this point.
000700000000      **
000800000000      ** To compile:
000900000000      **    (first follow the instructions to compile XLPARSER4)
001000000000      **>   CRTBNDRPG XLPDEMOF SRCFILE(QRPGLESRC) DBGVIEW(*LIST)
001100000000      **
001200000000     H DFTACTGRP(*NO) BNDDIR('XLPARSE') ACTGRP(*NEW)
001300000000
001400000000     FQSYSPRT   O    F   80        PRINTER OFLIND(*INOF)
001500000000
001600000000      /copy XLPARSE_H
001700000000
001800000000     D XLPDEMOF        PR                  ExtPgm('XLPDEMOF')
001900000000     D   peXSSF                       1n   const options(*nopass)
002000000000     D XLPDEMOF        PI
002100000000     D   peXSSF                       1n   const options(*nopass)
002200000000
002300000000     D IFSDIR          C                   '/tmp'
002400000000
002500000000     D charcell        PR             1N
002600000000     D   Sheet                     1024A   varying const
002700000000     D   Row                         10I 0 value
002800000000     D   Column                       5I 0 value
002900000000     D   Value                    32767A   varying const
003000000000
003100000000     D numbercell      PR             1N
003200000000     D   Sheet                     1024A   varying const
003300000000     D   Row                         10I 0 value
003400000000     D   Column                       5I 0 value
003500000000     D   Value                        8F   value
003600000000
003700000000     D formulacell     PR             1N
003800000000     D   Sheet                     1024A   varying const
003900000000     D   Row                         10I 0 value
004000000000     D   Col                          5I 0 value
004100000000     D   Value                        8F   value
004200000000     D   Nan                          5I 0 value
004300000000     D   Formula                  32767A   varying const
004400000000
004500000000     D prt             ds            80
004600000000
004700000000     D filename        s            200a   varying
004800000000     D msg             s             50A
004900000000     D wait            s              1A
005000000000
005100000000      /free
005200000000
005300000000         // FIXME: Change the path below to the place where you put
005400000000         //       TestFormula.xls in your IFS!
005500000000
005600000000         if (%parms>=1 and peXSSF);
005700000000            filename = IFSDIR + '/TestFormula.xlsx';
005800000000         else;
005900000000            filename = IFSDIR + '/TestFormula.xls';
006000000000         endif;
006100000000
006200000000         if xlparse_workbook( filename
006300000000                            : %paddr(NumberCell)
006400000000                            : %paddr(CharCell)
006500000000                            : %paddr(FormulaCell)
006600000000                            : *OFF ) < 0;
006700000000             msg = 'Error occcured while parsing spreadsheet.';
006800000000             dsply msg ' ' wait;
006900000000         endif;
007000000000
007100000000         *inlr = *on;
007200000000         return;
007300000000
007400000000       /end-free
007500000000
007600000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
007700000000      *  This is called once for each character cell in the spreadsheet
007800000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
007900000000     P charcell        B
008000000000     D charcell        PI             1N
008100000000     D   Sheet                     1024A   varying const
008200000000     D   Row                         10I 0 value
008300000000     D   Column                       5I 0 value
008400000000     D   Value                    32767A   varying const
008500000000      /free
008600000000
008700000000          prt = 'Row=' + %char(row) + ', Col=' + %char(Column)
008800000000              + ', value = ' + Value + ' (charcell)';
008900000000          write QSYSPRT prt;
009000000000
009100000000          return *OFF;
009200000000
009300000000      /end-free
009400000000     P                 E
009500000000
009600000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
009700000000      *  This is called once for each numeric cell in the spreadsheet
009800000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
009900000000     P numbercell      B
010000000000     D numbercell      PI             1N
010100000000     D   Sheet                     1024A   varying const
010200000000     D   Row                         10I 0 value
010300000000     D   Column                       5I 0 value
010400000000     D   Value                        8F   value
010500000000
010600000000     D firstpage       s              1N   static inz(*ON)
010700000000     D qtr             s             10I 0
010800000000
010900000000      /free
011000000000
011100000000          prt = 'Row=' + %char(row) + ', Col=' + %char(Column)
011200000000              + ', value = ' + %char(%dec(Value:3:0))
011300000000              + ' (numbercell)';
011400000000          write QSYSPRT prt;
011500000000
011600000000          return *off;
011700000000      /end-free
011800000000     P                 E
011900000000
012000000000     P formulacell     B
012100000000     D formulacell     PI             1N
012200000000     D   Sheet                     1024A   varying const
012300000000     D   Row                         10I 0 value
012400000000     D   Col                          5I 0 value
012500000000     D   Value                        8F   value
012600000000     D   Nan                          5I 0 value
012700000000     D   Formula                  32767A   varying const
012800000000      /free
012900000000
013000000000          if (Nan = 0);
013100000000             prt = 'Row=' + %char(row) + ', Col=' + %char(Col)
013200000000                 + ', value = ' + %char(%dec(Value:3:0))
013300000000                 + ' (formulacell)';
013400000000             write QSYSPRT prt;
013500000000          else;
013600000000             prt = 'Row=' + %char(row) + ', Col=' + %char(Col)
013700000000                 + ', value = NaN (Not a Number)'
013800000000                 + ' (formulacell)';
013900000000             write QSYSPRT prt;
014000000000          endif;
014100000000
014200000000         return *Off;
014300000000      /end-free
014400000000     P                 E
