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
002700000000      /if defined(XLPARSE_H)
002800000000      /eof
002900000000      /endif
003000000000      /define XLPARSE_H
003100000000
003200000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
003300000000      * xlparse_workbook(): Parse the contents of an XLS workbook
003400000000      *
003500000000      *      Filename = /path/to/my/file.xls
003600000000      *       Numeric = procedure to call for each numeric cell in sheet
003700000000      *        String = procedure to call for each string cell in sheet
003800000000      *       Formula = procedure to call for each formula cell in sheet
003900000000      *
004000000000      * Returns 0 if successful, or -1 when a Java exception occurs
004100000000      *
004200000000      *  Callback procedures should be prototyped as follows:
004300000000      *
004400000000      *  D MyString        PR             1N
004500000000      *  D   Sheet                     1024A   varying const
004600000000      *  D   Row                         10I 0 value
004700000000      *  D   Col                          5I 0 value
004800000000      *  D   Value                    32767A   varying const
004900000000      *
005000000000      *  D MyNumeric       PR             1N
005100000000      *  D   Sheet                     1024A   varying const
005200000000      *  D   Row                         10I 0 value
005300000000      *  D   Col                          5I 0 value
005400000000      *  D   Value                        8F   value
005500000000      *
005600000000      *  D MyFormula       PR             1N
005700000000      *  D   Sheet                     1024A   varying const
005800000000      *  D   Row                         10I 0 value
005900000000      *  D   Col                          5I 0 value
006000000000      *  D   Value                        8F   value
006100000000      *  D   Nan                          5I 0 value
006200000000      *  D   Formula                  32767A   varying const
006300000000      *
006400000000      *  The following information is sent to them:
006500000000      *
006600000000      *          Sheet = name of worksheet
006700000000      *            Row = row number (where 0 is the first row)
006800000000      *            Col = column number (where 0=A, 1=B, etc)
006900000000      *          Value = numeric or string of formula value of cell
007000000000      *
007100000000      *   You may return *ON to cancel processing, or *OFF otherwise
007200000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
007300000000     D xlparse_workbook...
007400000000     D                 PR            10I 0
007500000000     D    Filename                 1024A   const varying
007600000000     D    Numeric                      *   procptr value
007700000000     D    String                       *   procptr value
007800000000     D    Formula                      *   procptr value options(*nopass)
007900000000     D    RetainLog                    N   const options(*nopass:*omit)
008000000000
008100000000
008200000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
008300000000      * xlparse_notify():  Enable notifications.  This will cause
008400000000      *      XLPARSER4 to call subprocedures before and after
008500000000      *      each row that's read from the spreadsheet.
008600000000      *
008700000000      *    StartRow = (input) routine to call when a row begins
008800000000      *      EndRow = (input) routine to call when a row ends
008900000000      *
009000000000      * You can set one/both of these procedure pointers to *NULL
009100000000      * if you do not want to receive notification callbacks for
009200000000      * before/after events.
009300000000      *
009400000000      * A callback for either routine (before or after) should be
009500000000      * prototyped like this:
009600000000      *
009700000000      *      D mycallback      PR
009800000000      *      D  Sheet                      1024A   const varying
009900000000      *      D  Row                          10I 0 value
010000000000      *+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
010100000000     D xlparse_notify...
010200000000     D                 PR
010300000000     D   startRow                      *   value procptr
010400000000     D   endRow                        *   value procptr
