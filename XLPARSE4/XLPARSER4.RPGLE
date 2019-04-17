000100000000     /*-                                                                            +
000200000000      * Copyright (c) 2003-2010 Scott C. Klement                                    +
000300000000      * All rights reserved.                                                        +
000400000000      *                                                                             +
000500000000      * Redistribution and use in source and binary forms, with or without          +
000600000000      * modification, are permitted provided that the following conditions          +
000700000000      * are met:                                                                    +
000800000000      * 1. Redistributions of source code must retain the above copyright           +
000900000000      *    notice, this list of conditions and the following disclaimer.            +
001000000000      * 2. Redistributions in binary form must reproduce the above copyright        +
001100000000      *    notice, this list of conditions and the following disclaimer in the      +
001200000000      *    documentation and/or other materials provided with the distribution.     +
001300000000      *                                                                             +
001400000000      * THIS SOFTWARE IS PROVIDED BY THE AUTHOR AND CONTRIBUTORS ''AS IS'' AND      +
001500000000      * ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE       +
001600000000      * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE  +
001700000000      * ARE DISCLAIMED.  IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE     +
001800000000      * FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL  +
001900000000      * DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS     +
002000000000      * OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)       +
002100000000      * HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT  +
002200000000      * LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY   +
002300000000      * OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF      +
002400000000      * SUCH DAMAGE.                                                                +
002500000000      *                                                                             +
002600000000      */                                                                            +
002700000000      * Service Program for parsing Excel spreadsheets (Event Model)
002800000000      *                                 Scott Klement, January 2004
002900000000      *
003000000000      *  Licensed program requirements to compile:
003100000000      *   - ILE RPG Compiler (57xx-WDS, opt 31)
003200000000      *   - IBM Developer Kit for Java (57xx-JV1, *base)
003300000000      *   - One of the Java Developer Kit 5.0 or higher (57xx-JV1, opt 7+)
003400000000      *   - QShell (57xx-SS1, opt 30) not required, but highly recommended
003500000000      *
003600000000      * Requirements for parsing Binary Excel (XLS) workbooks:
003700000000      *   - Scott Klement's XLPARSE (from March 2010 or later)
003800000000      *       = xlparse.jar
003900000000      *   - Apache POI version 3.6 (or higher)
004000000000      *       = poi-3.6-20091214.jar
004100000000      *
004200000000      * Requirements for parsing Excel XML (XLSX) workbooks:
004300000000      *   - Scott Klement's XLPARSE (from March 2010 or later)
004400000000      *       = xlparse.jar
004500000000      *   - Apache POI version 3.6 (or higher)
004600000000      *       = poi-3.6-20091214.jar
004700000000      *       = poi-ooxml-3.6-20091214.jar
004800000000      *       = poi-ooxml-schemas-3.6-20091214.jar
004900000000      *   - Apache DOM4J 1.6.1 (or higher)
005000000000      *       = dom4j-1.6.1.jar
005100000000      *   - Apache XmlBeans 2.5.0 (or higher)
005200000000      *       = jsr173_1.0_api.jar
005300000000      *       = xbean.jar
005400000000      *
005500000000      *  All of the above JAR files should be in your CLASSPATH
005600000000      *  before the JVM is loaded in the job.
005700000000      *
005800000000      * To compile:
005900000000      *    - verify that the XLPARSE_H member has been uploaded
006000000000      *        to a QRPGLESRC file in your *LIBL.
006100000000      *    - verify that the XLPARSER4.bnd file has been uploaded
006200000000      *        with member name XLPARSER4 to a QSRVSRC file in
006300000000      *        your *LIBL.
006400000000      *    - run the following commands:
006500000000      *>      CRTRPGMOD MODULE(XLPARSER4) DBGVIEW(*ALL)
006600000000      *>      CRTSRVPGM SRVPGM(XLPARSER4) MODULE(XLPARSER4) -
006700000000      *>                EXPORT(*SRCFILE) SRCMBR(XLPARSER4) -
006800000000      *>                ACTGRP(XLPARSE) BNDDIR(QC2LE)
006900000000      *> ign: CRTBNDDIR BNDDIR(XLPARSE)
007000000000      *> ign: ADDBNDDIRE BNDDIR(XLPARSE) OBJ((XLPARSER4 *SRVPGM))
007100000000      *
007200000000      *
007300000000     H NOMAIN OPTION(*SRCSTMT: *NODEBUGIO: *NOSHOWCPY)
007400000000
007500000000      /copy xlparse_h
007600000000
007700000000     D XLPARSE_CLASS   C                   'com.iseriesnetwork.clubtech.-
007800000000     D                                     xlparse.ParseSheet'
007900000000     D XSSFPARSE_CLASS...
008000000000     D                 C                   'com.iseriesnetwork.clubtech.-
008100000000     D                                     xlparse.ParseSheetXssf'
008200000000     D CBCLASS         C                   'com.iseriesnetwork.clubtech.-
008300000000     D                                     xlparse.ParseSheetCallback'
008400000000
008500000000     D callbackStringCell...
008600000000     D                 PR                  ExtProc(*JAVA: CBCLASS:
008700000000     D                                             'callbackStringCell')
008800000000     D                                     static
008900000000     D  Sheet                      1024A   const varying
009000000000     D  Row                          10I 0 value
009100000000     D  Column                        5I 0 value
009200000000     D  Value                     32767A   const varying
009300000000
009400000000     D callbackNumericCell...
009500000000     D                 PR                  ExtProc(*JAVA: CBCLASS:
009600000000     D                                             'callbackNumericCell')
009700000000     D                                     static
009800000000     D  Sheet                      1024A   const varying
009900000000     D  Row                          10I 0 value
010000000000     D  Column                        5I 0 value
010100000000     D  Value                         8F   value
010200000000
010300000000     D callbackFormulaCell...
010400000000     D                 PR                  ExtProc(*JAVA: CBCLASS:
010500000000     D                                             'callbackFormulaCell')
010600000000     D                                     static
010700000000     D  Sheet                      1024A   const varying
010800000000     D  Row                          10I 0 value
010900000000     D  Column                        5I 0 value
011000000000     D  Value                         8F   value
011100000000     D  Nan                           5I 0 value
011200000000     D  Formula                   32767A   const varying
011300000000
011400000000     D StringProc      PR             1N   ExtProc(pStringProc)
011500000000     D   Sheet                     1024A   varying const
011600000000     D   Row                         10I 0 value
011700000000     D   Column                       5I 0 value
011800000000     D   Value                    32767A   varying const
011900000000
012000000000     D NumericProc     PR             1N   ExtProc(pNumericProc)
012100000000     D   Sheet                     1024A   varying const
012200000000     D   Row                         10I 0 value
012300000000     D   Column                       5I 0 value
012400000000     D   Value                        8F   value
012500000000
012600000000     D FormulaProc     PR             1N   ExtProc(pFormulaProc)
012700000000     D   Sheet                     1024A   varying const
012800000000     D   Row                         10I 0 value
012900000000     D   Column                       5I 0 value
013000000000     D   Value                        8F   value
013100000000     D   Nan                          5I 0 value
013200000000     D   Formula                  32767A   varying const
013300000000
013400000000     D EndRowProc      PR                  ExtProc(pEndRowProc)
013500000000     D   Sheet                     1024A   varying const
013600000000     D   Row                         10I 0 value
013700000000
013800000000     D NewRowProc      PR                  ExtProc(pNewRowProc)
013900000000     D   Sheet                     1024A   varying const
014000000000     D   Row                         10I 0 value
014100000000
014200000000     D Escape          PR             1N
014300000000     D   StackCnt                    10I 0 value
014400000000     D   MsgTxt                     256A   const
014500000000
014600000000     D open_fds        PR
014700000000     D isXssf          PR            10i 0
014800000000     D   pathName                  1024a   const varying
014900000000
015000000000     D Last            ds                  qualified
015100000000     D   Sheet                     1024a   varying inz('*NONE')
015200000000     D   Row                         10i 0 inz(-1)
015300000000
015400000000     D pStringProc     s               *   procptr inz(*NULL)
015500000000     D pNumericProc    s               *   procptr inz(*NULL)
015600000000     D pFormulaProc    s               *   procptr inz(*NULL)
015700000000     D pEndRowProc     s               *   procptr inz(*NULL)
015800000000     D pNewRowProc     s               *   procptr inz(*NULL)
015900000000
016000000000
016100000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
016200000000      * xlparse_notify():  Enable notifications.  This will cause
016300000000      *      XLPARSER4 to call subprocedures before and after
016400000000      *      each row that's read from the spreadsheet.
016500000000      *
016600000000      *    StartRow = (input) routine to call when a row begins
016700000000      *      EndRow = (input) routine to call when a row ends
016800000000      *
016900000000      * You can set one/both of these procedure pointers to *NULL
017000000000      * if you do not want to receive notification callbacks for
017100000000      * before/after events.
017200000000      *
017300000000      * A callback for either routine (before or after) should be
017400000000      * prototyped like this:
017500000000      *
017600000000      *      D mycallback      PR
017700000000      *      D  Sheet                      1024A   const varying
017800000000      *      D  Row                          10I 0 value
017900000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
018000000000     P xlparse_notify...
018100000000     P                 B                   EXPORT
018200000000     D xlparse_notify...
018300000000     D                 PI
018400000000     D   startRow                      *   value procptr
018500000000     D   endRow                        *   value procptr
018600000000      /free
018700000000         pNewRowProc = startRow;
018800000000         pEndRowProc = endRow;
018900000000      /end-free
019000000000     P                 e
019100000000
019200000000
019300000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
019400000000      * xlparse_workbook(): Parse the contents of an XLS workbook
019500000000      *
019600000000      *      Filename = /path/to/my/file.xls
019700000000      *       Numeric = procedure to call for each numeric cell in sheet
019800000000      *        String = procedure to call for each string cell in sheet
019900000000      *       Formula = procedure to call for each formula cell in sheet
020000000000      *
020100000000      * Returns 0 if successful, or -1 when a Java exception occurs
020200000000      *
020300000000      *  Callback procedures should be prototyped as follows:
020400000000      *
020500000000      *  D MyString        PR             1N
020600000000      *  D   Sheet                     1024A   varying const
020700000000      *  D   Row                         10I 0 value
020800000000      *  D   Col                          5I 0 value
020900000000      *  D   Value                    32767A   varying const
021000000000      *
021100000000      *  D MyNumeric       PR             1N
021200000000      *  D   Sheet                     1024A   varying const
021300000000      *  D   Row                         10I 0 value
021400000000      *  D   Col                          5I 0 value
021500000000      *  D   Value                        8F   value
021600000000
021700000000      *  D MyFormula       PR             1N
021800000000      *  D   Sheet                     1024A   varying const
021900000000      *  D   Row                         10I 0 value
022000000000      *  D   Col                          5I 0 value
022100000000      *  D   Value                        8F   value
022200000000      *  D   Nan                          5I 0 value
022300000000      *  D   Formula                  32767A   varying const
022400000000      *
022500000000      *  The following information is sent to them:
022600000000      *
022700000000      *          Sheet = name of worksheet
022800000000      *            Row = row number (where 0 is the first row)
022900000000      *            Col = column number (where 0=A, 1=B, etc)
023000000000      *          Value = numeric or string of formula value of cell
023100000000      *
023200000000      *   You may return *ON to cancel processing, or *OFF otherwise
023300000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
023400000000     P xlparse_workbook...
023500000000     P                 B                   EXPORT
023600000000     D xlparse_workbook...
023700000000     D                 PI            10I 0
023800000000     D    Filename                 1024A   const varying
023900000000     D    Numeric                      *   procptr value
024000000000     D    String                       *   procptr value
024100000000     D    Formula                      *   procptr value options(*nopass)
024200000000     D    RetainLog                    N   const options(*nopass:*omit)
024300000000
024400000000     D unlink          PR            10I 0 ExtProc('unlink')
024500000000     D   path                          *   Value options(*string)
024600000000
024700000000     D ParseSheet_parse...
024800000000     D                 PR                  ExtProc(*JAVA
024900000000     D                                            : XLPARSE_CLASS
025000000000     D                                            : 'parse201003' )
025100000000     D                                     static
025200000000     D    Filename                 1024A   const varying
025300000000
025400000000     D ParseSheetXssf_parse...
025500000000     D                 PR                  ExtProc(*JAVA
025600000000     D                                            : XSSFPARSE_CLASS
025700000000     D                                            : 'parse' )
025800000000     D                                     static
025900000000     D    Filename                 1024A   const varying
026000000000
026100000000     D Retain          s                   like(RetainLog)
026200000000     D RC              s             10I 0
026300000000
026400000000      /free
026500000000            if (%parms() >= 4);
026600000000                pFormulaProc = Formula;
026700000000            else;
026800000000                pFormulaProc = *null;
026900000000            endif;
027000000000
027100000000            if (%parms() >= 5 and %addr(RetainLog)<>*NULL);
027200000000                Retain = RetainLog;
027300000000            else;
027400000000                Retain = *off;
027500000000            endif;
027600000000
027700000000            pNumericProc = Numeric;
027800000000            pStringProc = String;
027900000000
028000000000            RC = 0;
028100000000            reset Last;
028200000000
028300000000            monitor;
028400000000               open_fds();
028500000000               if isXssf(filename) = 1;
028600000000                  ParseSheetXssf_parse(Filename);
028700000000               else;
028800000000                  ParseSheet_parse(Filename);
028900000000               endif;
029000000000            on-error;
029100000000               RC = -1;
029200000000            endmon;
029300000000
029400000000            if (Last.Row>=0 and pEndRowProc<>*NULL);
029500000000               EndRowProc(Last.Sheet: Last.Row);
029600000000            endif;
029700000000
029800000000            return RC;
029900000000
030000000000      /end-free
030100000000     P                 E
030200000000
030300000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
030400000000      * hssf_callback_numeric_cell():  This is called back by Java
030500000000      *   for each numeric cell in a spreadsheet.
030600000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
030700000000     P callbackNumericCell...
030800000000     P                 B                   EXPORT
030900000000     D callbackNumericCell...
031000000000     D                 PI                  static
031100000000     D  Sheet                      1024A   const varying
031200000000     D  Row                          10I 0 value
031300000000     D  Col                           5I 0 value
031400000000     D  Value                         8F   value
031500000000      /free
031600000000          if (Last.Row<>Row or Last.Sheet<>Sheet);
031700000000              if (Last.Row>=0 and pEndRowProc<>*NULL);
031800000000                 EndRowProc(Last.Sheet: Last.Row);
031900000000              endif;
032000000000              if (pNewRowProc<>*NULL);
032100000000                 NewRowProc(Sheet: Row);
032200000000              endif;
032300000000              Last.Sheet = Sheet;
032400000000              Last.Row   = Row;
032500000000          endif;
032600000000
032700000000          if (pNumericProc <> *NULL);
032800000000             if (NumericProc(Sheet: Row: Col: Value) = *On);
032900000000                escape(1: 'End Of Data');
033000000000             endif;
033100000000          elseif (pStringProc <> *NULL);
033200000000             if (StringProc(Sheet: Row: Col: %char(Value)) = *On);
033300000000                escape(1: 'End Of Data');
033400000000             endif;
033500000000          endif;
033600000000
033700000000      /end-free
033800000000     P                 E
033900000000
034000000000
034100000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
034200000000      * hssf_callback_string_cell():  This is called back by Java
034300000000      *   for each string cell ("label") in a spreadsheet.
034400000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
034500000000     P callbackStringCell...
034600000000     P                 B                   EXPORT
034700000000     D callbackStringCell...
034800000000     D                 PI                  static
034900000000     D  Sheet                      1024A   const varying
035000000000     D  Row                          10I 0 value
035100000000     D  Col                           5I 0 value
035200000000     D  Value                     32767A   const varying
035300000000
035400000000      /free
035500000000          if (Last.Row<>Row or Last.Sheet<>Sheet);
035600000000              if (Last.Row>=0 and pEndRowProc<>*NULL);
035700000000                 EndRowProc(Last.Sheet: Last.Row);
035800000000              endif;
035900000000              if (pNewRowProc<>*NULL);
036000000000                 NewRowProc(Sheet: Row);
036100000000              endif;
036200000000              Last.Sheet = Sheet;
036300000000              Last.Row   = Row;
036400000000          endif;
036500000000
036600000000          if (pStringProc <> *NULL);
036700000000             if (StringProc(Sheet: Row: Col: Value) = *On);
036800000000                escape(1: 'End Of Data');
036900000000             endif;
037000000000          endif;
037100000000
037200000000      /end-free
037300000000     P                 E
037400000000
037500000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
037600000000      * hssf_callback_formula_cell():  This is called back by Java
037700000000      *   for each formula cell in a spreadsheet.
037800000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
037900000000     P callbackFormulaCell...
038000000000     P                 B                   EXPORT
038100000000     D callbackFormulaCell...
038200000000     D                 PI                  static
038300000000     D  Sheet                      1024A   const varying
038400000000     D  Row                          10I 0 value
038500000000     D  Col                           5I 0 value
038600000000     D  Value                         8F   value
038700000000     D  Nan                           5I 0 value
038800000000     D  Formula                   32767A   const varying
038900000000      /free
039000000000          if (Last.Row<>Row or Last.Sheet<>Sheet);
039100000000              if (Last.Row>=0 and pEndRowProc<>*NULL);
039200000000                 EndRowProc(Last.Sheet: Last.Row);
039300000000              endif;
039400000000              if (pNewRowProc<>*NULL);
039500000000                 NewRowProc(Sheet: Row);
039600000000              endif;
039700000000              Last.Sheet = Sheet;
039800000000              Last.Row   = Row;
039900000000          endif;
040000000000
040100000000          if (pFormulaProc <> *NULL);
040200000000             if (FormulaProc(Sheet: Row: Col: Value: NaN: Formula) = *On);
040300000000                escape(1: 'End Of Data');
040400000000             endif;
040500000000          elseif (NaN=0 and pNumericProc <> *NULL);
040600000000             if (NumericProc(Sheet: Row: Col: Value) = *On);
040700000000                escape(1: 'End Of Data');
040800000000             endif;
040900000000          elseif (pStringProc <> *NULL);
041000000000             if (StringProc(Sheet: Row: Col: Formula) = *On);
041100000000                escape(1: 'End Of Data');
041200000000             endif;
041300000000          endif;
041400000000
041500000000      /end-free
041600000000     P                 E
041700000000
041800000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
041900000000      * Escape():
042000000000      *   This sends back a fatal error message (an "escape" message)
042100000000      *   with MSGID CPF9897 to the specified level in the call stack,
042200000000      *   ending any higher level programs.
042300000000      *
042400000000      *   peStackCnt = (input) Specifies which program's msgq to send
042500000000      *       *ESCAPE message to by specifying a number of levels back...
042600000000      *   peMsgTxt   = (input) Text of message to send
042700000000      *
042800000000      * Returns *OFF if it failed.  If it succeeds, this routine
042900000000      *   does not return, since it's call level is ended.
043000000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
043100000000     P Escape          B
043200000000     D Escape          PI             1N
043300000000     D   StackCnt                    10I 0 value
043400000000     D   MsgTxt                     256A   const
043500000000
043600000000     D QMHSNDPM        PR                  ExtPgm('QMHSNDPM')
043700000000     D   MessageID                    7A   Const
043800000000     D   QualMsgF                    20A   Const
043900000000     D   MsgData                    256A   Const
044000000000     D   MsgDtaLen                   10I 0 Const
044100000000     D   MsgType                     10A   Const
044200000000     D   CallStkEnt                  10A   Const
044300000000     D   CallStkCnt                  10I 0 Const
044400000000     D   MessageKey                   4A
044500000000     D   ErrorCode                 1024A   options(*varsize)
044600000000
044700000000     D dsEC            DS
044800000000     D  BytesProv                    10I 0 inz(0)
044900000000     D  BytesAvail                   10I 0 inz(0)
045000000000
045100000000     D MsgLen          S             10I 0
045200000000     D MsgKey          S              4A
045300000000
045400000000      /free
045500000000
045600000000         MsgLen = %len(%trimr(MsgTxt));
045700000000         if (MsgLen < 1);
045800000000            return *off;
045900000000         endif;
046000000000
046100000000         QMHSNDPM( 'CPF9897'
046200000000                 : 'QCPFMSG   *LIBL'
046300000000                 : MsgTxt
046400000000                 : MsgLen
046500000000                 : '*ESCAPE'
046600000000                 : '*'
046700000000                 : StackCnt
046800000000                 : MsgKey
046900000000                 : dsEC );
047000000000
047100000000         return *off;
047200000000
047300000000      /end-free
047400000000     P                 E
047500000000
047600000000
047700000000      *-----------------------------------------------------------------
047800000000      *  open_fds():  Make sure we have 3 FDs open before starting JVM
047900000000      *
048000000000      *  The JVM will encounter strange errors (I/O errors?) if
048100000000      *  there aren't at least 3 file descriptors open.  This code
048200000000      *  verifies that there are at least 3.
048300000000      *-----------------------------------------------------------------
048400000000     P open_fds        B
048500000000     D open_fds        PI
048600000000
048700000000     D started         s              1n   inz(*off) static
048800000000
048900000000     D SndPgmMsg       PR                  ExtPgm('QMHSNDPM')
049000000000     D   MessageID                    7A   Const
049100000000     D   QualMsgF                    20A   Const
049200000000     D   MsgData                     80A   Const
049300000000     D   MsgDtaLen                   10I 0 Const
049400000000     D   MsgType                     10A   Const
049500000000     D   CallStkEnt                  10A   Const
049600000000     D   CallStkCnt                  10I 0 Const
049700000000     D   MessageKey                   4A
049800000000     D   ErrorCode                32767A   options(*varsize)
049900000000
050000000000     D my_open         PR            10I 0 ExtProc('open')
050100000000     D  path                           *   value options(*string)
050200000000     D  openflags                    10I 0 value
050300000000     D  mode                         10U 0 value options(*nopass)
050400000000     D my_unlink       PR            10I 0 ExtProc('unlink')
050500000000     D   path                          *   Value options(*string)
050600000000     D my_dup          PR            10I 0 ExtProc('dup')
050700000000     D   fildes                      10I 0 Value
050800000000
050900000000     D my_Mode         C                   const(438)
051000000000     D my_tmpnam       PR              *   extproc('_C_IFS_tmpnam')
051100000000     D   string                      39A   options(*omit)
051200000000     D my_close        PR            10I 0 ExtProc('close')
051300000000     D  fildes                       10I 0 value
051400000000     D O_RDWR          C                   4
051500000000     D O_CREAT         C                   8
051600000000
051700000000     D ErrorNull       ds
051800000000     D   BytesProv                   10I 0 inz(0)
051900000000     D   BytesAvail                  10I 0 inz(0)
052000000000
052100000000     D fd              s             10I 0
052200000000     D filename        s               *
052300000000     D key             s              4A
052400000000
052500000000      /free
052600000000         if (started);
052700000000           return;
052800000000         endif;
052900000000
053000000000         fd = my_open('/dev/null': O_RDWR);
053100000000         if (fd = -1);
053200000000             filename = my_tmpnam(*omit);
053300000000             fd = my_open(filename: O_RDWR+O_CREAT: my_Mode);
053400000000             my_unlink(filename);
053500000000         endif;
053600000000
053700000000         if (fd > 2);
053800000000            my_close(fd);
053900000000         endif;
054000000000
054100000000         dow ( fd < 2 );
054200000000            if (fd = -1);
054300000000                SndPgmMsg( 'CPF9897'
054400000000                         : 'QCPFMSG   *LIBL'
054500000000                         : 'Unable to open three descriptors!'
054600000000                         : 80
054700000000                         : '*ESCAPE'
054800000000                         : '*PGMBDY'
054900000000                         : 1
055000000000                         : Key
055100000000                         : ErrorNull );
055200000000                return;
055300000000            endif;
055400000000            fd = my_dup(fd);
055500000000         enddo;
055600000000      /end-free
055700000000     P                 E
055800000000
055900000000
056000000000     P isXssf          B
056100000000     D isXssf          PI            10i 0
056200000000     D   pathName                  1024a   const varying
056300000000
056400000000     D my_open         PR            10I 0 ExtProc('open')
056500000000     D  path                           *   value options(*string)
056600000000     D  openflags                    10I 0 value
056700000000     D  mode                         10U 0 value options(*nopass)
056800000000     D my_read         PR            10I 0 ExtProc('read')
056900000000     D  fildes                       10i 0 value
057000000000     D  buf                            *   value
057100000000     D  bytes                        10U 0 value
057200000000     D my_close        PR            10I 0 ExtProc('close')
057300000000     D  fildes                       10I 0 value
057400000000     D O_RDONLY        C                   1
057500000000
057600000000     D fd              s             10i 0
057700000000     D len             s             10i 0
057800000000     D data            s              8a
057900000000
058000000000      /free
058100000000         // can't open file.  Never a good sign.
058200000000         fd = my_open(%trimr(pathName): O_RDONLY);
058300000000         if (fd = -1);
058400000000            return -1;
058500000000         endif;
058600000000
058700000000         len = my_read(fd: %addr(data): %size(data));
058800000000         my_close(fd);
058900000000
059000000000         // file isn't even 8 bytes long.  Can't be
059100000000         // a valid Xssf file, then.
059200000000         if (len < %size(data));
059300000000            return 0;
059400000000         endif;
059500000000
059600000000         if data = x'D0CF11E0A1B11AE1';
059700000000            return 0;
059800000000         else;
059900000000            return 1;
060000000000         endif;
060100000000      /end-free
060200000000     P                 E
