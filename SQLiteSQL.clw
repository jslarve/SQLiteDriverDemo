                    PROGRAM

!Updated 2020.06.22 - Minor cosmetic changes - Also loading a pre-built SQLIte db into memory instead of using INSERT script. You can load a whole lot of data very quickly.
!Updated 2020.06.14 - testing scope and lifetime using threads. Thanks a whole lot to Federico Navarro for the tips on using :Memory: and for creating the test window.
!Updated 2020.06.13 - Increased the query size, improved some of the error checking and column sizing. Works better now if you do a query such as 'PRAGMA database_list'.
!Updated 2020.06.12 - Used PRAGMA table_info to get the column information instead of parsing the create() statement. Added an option to include data type.

    INCLUDE('KEYCODES.CLW'),ONCE
    INCLUDE('SystemString.inc'),ONCE   


    OMIT('***')
 * Created with Clarion 11.0
 * User: Jeff Slarve
 * Date: 6/9/2020
 * Time: 11:45 PM
 * 
!MIT License
!
!Copyright (c) 2020 Jeff Slarve
!
!Permission is hereby granted, free of charge, to any person obtaining a copy
!of this software and associated documentation files (the "Software"), to deal
!in the Software without restriction, including without limitation the rights
!to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
!copies of the Software, and to permit persons to whom the Software is
!furnished to do so, subject to the following conditions:
!
!The above copyright notice and this permission notice shall be included in all
!copies or substantial portions of the Software.
!
!THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
!IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
!FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
!AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
!LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
!OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
!SOFTWARE.
 ***

                    MAP
                        SqliteTest(STRING pQuery)
                        MODULE('')
                            InvalidateRect(ulong, LONG, UNSIGNED),PASCAL !To help clear artifacts on resize
                        END
                    END

TURBO_FIELD_SIZE    EQUATE(2001) !Field size within the MyTurbo table. Set to suit your needs for accomodating your query results.
SPECIAL_DUMMY_TABLE EQUATE('DummyTable123xyz') !Random name to try not to collide. This table gets created in memory.
DEFAULT_QUERY       EQUATE('SELECT * FROM mem.Data WHERE slang LIKE("%data%")<13,10>--This query shows all rows that have "data" within the "Slang" column.<13,10>--Try a query of your own and press Ctrl+Enter or F8 to execute.')

DummyDBOwner        STRING(FILE:MaxFilePath+1) !Just a filename for a dummy SQLite table. Could be created under %TEMP% or something.

MyTurbo             FILE,DRIVER('SQLite','/TURBOSQL=TRUE'),THREAD,PRE(MyTurbo),OWNER(DummyDBOwner),CREATE !Just a dummy table, but it ain't no dummy.
RECORD                  RECORD
F1                          CSTRING(TURBO_FIELD_SIZE)
F2                          CSTRING(TURBO_FIELD_SIZE)
F3                          CSTRING(TURBO_FIELD_SIZE)
F4                          CSTRING(TURBO_FIELD_SIZE)
F5                          CSTRING(TURBO_FIELD_SIZE)
F6                          CSTRING(TURBO_FIELD_SIZE)
F7                          CSTRING(TURBO_FIELD_SIZE)
F8                          CSTRING(TURBO_FIELD_SIZE)
F9                          CSTRING(TURBO_FIELD_SIZE)
F10                         CSTRING(TURBO_FIELD_SIZE)
F11                         CSTRING(TURBO_FIELD_SIZE)
F12                         CSTRING(TURBO_FIELD_SIZE)
F13                         CSTRING(TURBO_FIELD_SIZE)
F14                         CSTRING(TURBO_FIELD_SIZE)
F15                         CSTRING(TURBO_FIELD_SIZE)
F16                         CSTRING(TURBO_FIELD_SIZE)
F17                         CSTRING(TURBO_FIELD_SIZE)
F18                         CSTRING(TURBO_FIELD_SIZE)
F19                         CSTRING(TURBO_FIELD_SIZE)
F20                         CSTRING(TURBO_FIELD_SIZE)
                        END !Add as many columns here as you need, but other things might need updating to support it.
                    END


c                   CLASS
CheckError              PROCEDURE(STRING pMessage),LONG !Generic error handler
InvalidateWindow        PROCEDURE                       !calls invalidaterect to refresh window on resize
                    END

MainWindow          WINDOW('Simple SQLite Tester Main Window'),AT(,,261,247),GRAY,SYSTEM, |
                        FONT('Segoe UI',9)
                        PROMPT('SQLiteDB file or :memory:    (Owner ID)'),AT(5,3,159), |
                            USE(?DBOwnerPrompt)
                        ENTRY(@s100),AT(5,14,235,9),USE(DummyDBOwner)
                        BUTTON,AT(243,12,14,12),USE(?LookupFileButton),ICON('openfold.ico')
                        BUTTON('&Test on Main Thread'),AT(5,27,50,20),USE(?TestButton)
                        STRING('SQLiteTest(''Query'') !Call SQLiteTest Procedure'),AT(63,33,191), |
                            USE(?TestOnMainThreadString),FONT('Consolas')
                        BUTTON('&Test on New Thread'),AT(5,52,50,20),USE(?TestNTButton)
                        STRING('START(SQLiteTest,25000,''Query'')'),AT(63,59,191,9),USE(?TestOnMainThreadString:2) |
                            ,FONT('Consolas')
                        BUTTON('&Create DB && Dummy Table'),AT(5,76,50,20),USE(?CreateDBButton)
                        STRING('CREATE(MyTurbo) !Not needed for :memory:'),AT(63,82,191,9), |
                            USE(?TestOnMainThreadString:3),FONT('Consolas')
                        BUTTON('&Create blank SqliteDB'),AT(5,101,50,20),USE(?CreateDBPropertyButton)
                        STRING('MyTurbo{{PROP:CreateDB} !Not needed for :memory:'),AT(63,107,191,9), |
                            USE(?TestOnMainThreadString:4),FONT('Consolas')
                        BUTTON('&Open File on Main Thread'),AT(5,125,50,20),USE(?OpenFileButton)
                        STRING('OPEN(MyTurbo)'),AT(63,133,191,9),USE(?TestOnMainThreadString:5), |
                            FONT('Consolas')
                        BUTTON('&Close File on Main Thread'),AT(5,150,50,20),USE(?CloseFileButton)
                        STRING('CLOSE(MyTurbo)'),AT(63,156,191,9),USE(?TestOnMainThreadString:6), |
                            FONT('Consolas')
                        BUTTON('&Disconnect DB'),AT(5,174,50,20),USE(?DisconnectButton)
                        STRING('MyTurbo{{PROP:Disconnect}'),AT(63,182,191,9),USE(?TestOnMainThreadString:7) |
                            ,FONT('Consolas')
                        BUTTON('Database &List'),AT(5,198,50,20),USE(?DatabaseListButton)
                        BUTTON('&Quit'),AT(5,222,50,20),USE(?QuitButton)
                        STRING('START(SQLiteTest,25000,''PRAGMA Database_List'')'),AT(63,203,191,9), |
                            USE(?TestOnMainThreadString:8),FONT('Consolas')
                    END

ThreadCounter       LONG

    CODE

        DummyDBOwner = ':memory:' ! Per Federico Navarro, and it works!
        OPEN(MainWindow)
        ACCEPT
            CASE EVENT()
            OF EVENT:CloseWindow
                LOOP ThreadCounter = 2 TO 64 !Some quickie code to close threads. Will re-visit later.
                    POST(EVENT:CloseWindow,,ThreadCounter)
                    YIELD()
                END
            END
            
            CASE ACCEPTED()
            OF ?LookupFileButton
                IF LOWER(CLIP(DummyDBOwner)) = ':memory:'
                    DummyDBOwner = '' !So that file dialog will actually work
                END                
                IF FILEDIALOG('Select Database',DummyDBOwner,'*.*',FILE:LongName+FILE:KeepDir)
                ELSE
                    IF NOT DummyDBOwner
                        DummyDBOwner = ':memory:'
                    END
                END                
                SELECT(?DummyDBOwner)
                DISPLAY(?DummyDBOwner)                
            OF ?TestButton
                SQLiteTest(DEFAULT_QUERY)
            OF ?TestNTButton
                START(SQLiteTest,25000,DEFAULT_QUERY)
            OF ?CreateDBButton
                !not needed for :memory:
                CREATE(MyTurbo) 
                IF c.CheckError('Error on Create')
                END
            OF ?CreateDBPropertyButton
                !not needed for :memory:
                MyTurbo{PROP:CreateDB}
                IF c.CheckError('Error on CreateDB')
                END
            OF ?OpenFileButton
                OPEN(MyTurbo)
                IF c.CheckError('Error on Open')
                END
            OF ?CloseFileButton
                CLOSE(MyTurbo)
                IF c.CheckError('Error on Close')
                END
            OF ?DisconnectButton
                MyTurbo{PROP:Disconnect}
                IF c.CheckError('Error on Disconnect')
                END
            OF ?DatabaseListButton
                START(SQLiteTest,25000,'PRAGMA Database_List<13,10>--See database list below. Also checkout these other queries<13,10>--SELECT * FROM mem.sqlite_master<13,10>--SELECT * FROM main.sqlite_master')
            OF ?QuitButton
                POST(EVENT:CloseWindow)
            END
        END
        CLOSE(MainWindow)

        
SQLiteTest          PROCEDURE(STRING pQuery)

SS                  SystemStringClass  !Utility for manipulating strings, loading files, executing sql
ClipboardSS         SystemStringClass  !For building an Excel friendly string to copy to the clipboard.

Ndx                 LONG  !Generic counter
ColumnNdx           LONG  !Counter for processing through columns



ResultQ             QUEUE
                        LIKE(MyTurbo:Record)
                    END

StartTime           LONG
EndTime             LONG
Elapsed             DECIMAL(10,2)  !For computing time
SQLQueryText        CSTRING(350000)!The query that you type in
Columns             LONG,DIM(20)   !For keeping track of the text length for column sizing
Characters          LONG           !For calculating column widths
ResultColumns       LONG           !Number of result columns (maxes at 20 in this demo)
ResultColumnsQ      QUEUE          !Result column names 
ColumnName              CSTRING(61)
DataType                CSTRING(61) 
                    END
ResultColumnsText   CSTRING(1000)  !Just for copying stuff to clipboard for pasting in query.
IncludeColumnHeadings BYTE(TRUE)   !When copying to clipboard, option to include the labels
IncludeDataType         BYTE(FALSE)
DBOpened         BYTE(FALSE)

Window              WINDOW('Simple SQLite Tester'),AT(,,454,217),CENTER,GRAY,IMM,SYSTEM,MAX, |
                        FONT('Segoe UI',12),RESIZE
                        PROMPT('SQL:'),AT(3,4),USE(?PROMPT1)
                        TEXT,AT(4,17,302,35),USE(SQLQueryText),VSCROLL,FONT('Consolas',11), |
                            ALRT(CtrlEnter)
                        BUTTON,AT(309,17,17,15),USE(?ExecuteSQLButton),KEY(F8Key),ICON(ICON:VCRplay), |
                            TIP('Execute Query (or at least try) - [F8]')
                        PROMPT('Result Columns for Copying:'),AT(299,4,102),USE(?ColumnsPrompt)
                        TEXT,AT(330,17,120,46),USE(ResultColumnsText),VSCROLL,COLOR(COLOR:BTNFACE), |
                            READONLY
                        PROMPT('Result:'),AT(3,56),USE(?ResultPrompt)
                        LIST,AT(4,67,446,128),USE(?SQLResultList),HIDE,HVSCROLL,FONT('Consolas',10), |
                            COLOR(COLOR:BTNFACE),FROM(ResultQ),FORMAT('20L(2)|M@s255@#2#20L(' & |
                            '2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s2' & |
                            '55@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(' & |
                            '2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s2' & |
                            '55@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(' & |
                            '2)|M@s255@')
                        PROMPT('Press F8 or click the button to the right to execute SQL query. ' & |
                            'Max 20 columns will display.'),AT(23,4,274),USE(?PROMPT2)
                        BUTTON('&Copy Result to Clipboard'),AT(3,199,95),USE(?CopyButton), |
                            TIP('Copy to clipboard (Excel Friendly, unless there are tabs in' & |
                            ' the data :-) )')
                        CHECK('Include Column Headings'),AT(103,201),USE(IncludeColumnHeadings), |
                            TIP('Include column headings in the clipboard export.')
                        CHECK('Include Data Type'),AT(199,201),USE(IncludeDataType),TIP('Include' & |
                            ' the data type on the displayed column headers')
                        BUTTON('Cl&ose'),AT(409,199,42,14),USE(?CloseButton),STD(STD:Close)
                    END

WindowText              CSTRING(201)

    CODE
        OPEN(MyTurbo)
        IF c.CheckError('Error on Open')
          !RETURN
        ELSE 
          DBOpened = True
        END
      MyTurbo{PROP:SQL} = 'SELECT COUNT(*) FROM mem.sqlite_master WHERE type = "table" and name = "' & SPECIAL_DUMMY_TABLE & '"' !Now checking mem.sqlite_master for our created table
      NEXT(MyTurbo)
      IF NOT ERRORCODE()
        !MESSAGE('Reusing existing memory table ' & MyTurbo:F1 & ' definition row/s','SQLite In Memory Demo')
          WindowText = 'Simple SQLite Tester - [Using Pre-loaded Data]'
      ELSE
        !this is kept for reusing Mem prefix, but duplicates internal structures
        !using both with DummyDBOwner :memory:
        MyTurbo{PROP:SQL} = 'ATTACH '':memory:'' AS Mem'             !This is the magic for using in-Memory with the SQLite driver
        IF c.CheckError('Error on ATTACH')
            RETURN
        END
        SS.FromFile('.\Create.mem.Data.sql')                         !I created separate .SQL files because combining the CREATE and INSERT commands into one query string didn't seem to work.
        MyTurbo{PROP:SQL} = SS.GetString()                           !Executing the CREATE script.
        IF c.CheckError('Create Statement')
            RETURN
        END
        SS.SetLen(0)                                                 !Clearing out the SystemString object for next use. 
!        SS.FromFile('.\mem.Data.sql')                                !Load the INSERT INTO sql into SystemString object
!        StartTime = CLOCK()                                          !Get the start time
!        MyTurbo{PROP:SQL} = SS.GetString()                           !Execute the INSERT
!        IF c.CheckError('INSERT Statement')
!            RETURN
!        END
        StartTime = CLOCK()                                          !Get the start time
        MyTurbo{PROP:SQL} = 'ATTACH ".\MockData.db" as MockData'     !Open the MockData database
        IF c.CheckError('Error on ATTACH MockData')
            RETURN
        END
        MyTurbo{PROP:SQL} = 'INSERT INTO mem.data SELECT * FROM MockData.data' !Load the data from Mockdata into memory.
        IF c.CheckError('Insert from MockData')
            RETURN
        END
        EndTime = CLOCK()
        MyTurbo{PROP:SQL} = 'DETACH MockData'                        !Detach from the database. We could leave it attached if desired, but we don't need it.
        IF c.CheckError('Detach MockData')
            RETURN
        END
            
        Elapsed = (EndTime - StartTime) *.01                         !Calculated elapsed time into a DECIMAL so it would format correctly in MESSAGE()
        MyTurbo{PROP:SQL} = 'SELECT COUNT(*) FROM Mem.Data'          !Getting the row count
        NEXT(MyTurbo)
        IF c.CheckError('RECORDS() Error')
            RETURN
        ELSE
            WindowText = 'Simple SQLite Tester - [Loaded ' & MyTurbo:F1 & ' rows from .\MockData.db to :memory: in ' & Elapsed & ' secs.]'
        END
        
        SS.SetLen(0)                                                 !Clear out SystemString
      END
        
        OPEN(Window)
        0{PROP:Hide} = TRUE
        0{PROP:Text} = WindowText
        SQLQueryText = pQuery
        ACCEPT
            CASE EVENT()
            OF EVENT:AlertKey
                CASE FIELD()
                OF ?SQLQueryText
                    CASE KEYCODE()
                    OF CtrlEnter
                        POST(EVENT:Accepted,?ExecuteSQLButton)
                    END                    
                END
            OF EVENT:OpenWindow
                POST(EVENT:Accepted,?ExecuteSQLButton) !We'll run the default query on startup
                POST(EVENT:Sized)
            OF EVENT:Sized
                0{PROP:Pixels} = TRUE
                ?SQLQueryText{PROP:Width        } =  0{PROP:Width}  - ?ExecuteSQLButton{PROP:Width} - ?SQLQueryText{PROP:XPOS} - ?ResultColumnsText{PROP:Width} - 14
                ?ExecuteSQLButton{PROP:XPos     } =  0{PROP:Width}  - ?ExecuteSQLButton{PROP:Width} - ?ResultColumnsText{PROP:Width} - 10
                ?ResultColumnsText{PROP:XPos    } =  0{PROP:Width}  - ?ResultColumnsText{PROP:Width} - 8
                ?ColumnsPrompt{PROP:XPos        } =  ?ResultColumnsText{PROP:XPos}
                ?SQLResultList{PROP:Width       } =  0{PROP:Width}  - ?SQLResultList{PROP:XPos}     - 8
                ?SQLResultList{PROP:Height      } =  0{PROP:Height} -  ?SQLResultList{PROP:YPos} - ?CloseButton{PROP:Height} - 8
                ?CloseButton{PROP:XPos          } =  0{PROP:Width}  - ?CloseButton{PROP:Width}      - 4
                ?CloseButton{PROP:YPos          } =  0{PROP:Height} - ?CloseButton{PROP:Height}     - 4
                ?CopyButton{PROP:Ypos           } =  ?CloseButton{PROP:YPos}
                ?IncludeColumnHeadings{PROP:YPos} =  ?CloseButton{PROP:YPos} + 4
                ?IncludeDataType{PROP:Ypos      } =  ?IncludeColumnHeadings{PROP:YPos}
                0{PROP:Pixels                   } =  FALSE
                IF 0{PROP:Hide} = TRUE
                    0{PROP:Hide} =  FALSE
                END
                c.InvalidateWindow
            END
            
            CASE ACCEPTED()
            OF ?IncludeDataType
                POST(EVENT:Accepted,?ExecuteSQLButton)
            OF ?CopyButton !Loop through queue and generate an Excel Friendly thing you can paste. NOTE: Should probably strip out TABS if present in data or this will break.
                ClipboardSS.SetLen(0)
                IF IncludeColumnHeadings !Process Column Headings
                    LOOP ColumnNdx = 1 TO ResultColumns !Get The Headers
                        GET(ResultColumnsQ,ColumnNdx)
                        ClipboardSS.Append(ResultColumnsQ.ColumnName)
                        IF ColumnNdx = ResultColumns
                            ClipboardSS.Append('<13,10>')
                        ELSE
                            ClipboardSS.Append('<9>')
                        END
                    END
                END                
                LOOP Ndx = 1 TO RECORDS(ResultQ) !Process Data
                    GET(ResultQ,Ndx)
                    LOOP ColumnNdx = 1 TO ResultColumns
                        ClipboardSS.Append(WHAT(ResultQ,ColumnNdx+1))
                        IF ColumnNdx = ResultColumns
                            ClipboardSS.Append('<13,10>')
                        ELSE
                            ClipboardSS.Append('<9>')
                        END
                    END
                END
                SETCLIPBOARD(ClipboardSS.GetString())
            OF ?ExecuteSQLButton 
                UPDATE(?SQLQueryText) !Just in case, we'll update the data from the control.
                FREE(ResultQ)                    !We'll process the results into the result queue
                CLEAR(ResultQ)
                DISPLAY(?SQLResultList)
                MyTurbo{PROP:SQL} = SQLQueryText !Execute SQL
                IF ERRORCODE() = 33
                    MESSAGE('No Results to Display')
                ELSE                    
                    IF c.CheckError('Executing Query') 
                    END
                END
                
                HIDE(?SQLResultList)
                LOOP
                    CLEAR(MyTurbo:RECORD)                        
                    NEXT(MyTurbo)
                    IF ERRORCODE()
                        BREAK
                    END      
                    ResultQ = MyTurbo:RECORD
                    ADD(ResultQ)
                END
                IF RECORDS(ResultQ) !We'll fill a special dummy table with the columns from the query result and parse them out.
                    SS.SetLen(0)
                    MyTurbo{PROP:SQL} = 'DROP TABLE IF EXISTS mem.' & SPECIAL_DUMMY_TABLE & ';' !If the dummy table exists, we'll kill it
                    IF NOT c.CheckError('Error Dropping Dummy')
                        MyTurbo{PROP:SQL} = 'CREATE TABLE mem.' & SPECIAL_DUMMY_TABLE & ' AS ' & SQLQueryText & ' LIMIT 0' !Gathering the column names/types
                        FREE(ResultColumnsQ)
                        ResultColumnsText = ''
                        IF NOT c.CheckError('Error Creating Dummy')
                            MyTurbo{PROP:SQL} = 'PRAGMA mem.table_info(' & SPECIAL_DUMMY_TABLE & ')'
                            IF NOT ERRORCODE()!c.CheckError('Error Selecting Dummy SQL')
                                LOOP
                                    NEXT(MyTurbo)
                                    IF ERRORCODE()
                                        BREAK
                                    END
                                    ResultColumnsQ.ColumnName = MyTurbo:F2
                                    ResultColumnsQ.DataType   = MyTurbo:F3
                                    ADD(ResultColumnsQ)
                                    ResultColumnsText = ResultColumnsText & ResultColumnsQ.ColumnName 
                                    IF IncludeDataType
                                        ResultColumnsText = ResultColumnsText & ' ' & ResultColumnsQ.DataType
                                    END
                                    ResultColumnsText = ResultColumnsText & '<13,10>'
                                END
                                
                                ResultColumns = RECORDS(ResultColumnsQ)
                                IF ResultColumns > MAXIMUM(Columns,1)
                                    ResultColumns = MAXIMUM(Columns,1)
                                END 
                                DISPLAY(?ResultColumnsText)
                                    
                            END
                        END
                        DISPLAY(?ResultColumnsText)
                    END
                END
                    
                CLEAR(Columns)
                IF RECORDS(ResultQ)
                    LOOP Ndx = 1 TO RECORDS(ResultQ) !attempting to get text width of each column
                        GET(ResultQ,Ndx)
                        LOOP ColumnNdx = 1 TO MAXIMUM(Columns,1) !Checking the widths of the data in each column of row
                            IF LEN(CLIP(WHAT(ResultQ,ColumnNdx+1))) > Columns[ColumnNdx]
                                Columns[ColumnNdx] = LEN(CLIP(WHAT(ResultQ,ColumnNdx+1))) + 1 !Setting new max widths
                            END
                        END
                    END
                ELSE
                    FREE(ResultColumnsQ)
                END
                
                0{PROP:Pixels} = TRUE !So that the column widths can be set more accurately
                LOOP ColumnNdx = 1 TO MAXIMUM(Columns,1)
                    Characters = Columns[ColumnNdx] !Getting number of characters 
                    CLEAR(ResultColumnsQ)
                    GET(ResultColumnsQ,ColumnNdx)
                    IF NOT ERRORCODE()
                        ?SQLResultList{PROPList:Header,ColumnNdx} = ResultColumnsQ.ColumnName & CHOOSE(IncludeDataType=TRUE, '<10>' & ResultColumnsQ.DataType,'')
                    ELSE
                        ?SQLResultList{PROPList:Header,ColumnNdx} = ''
                    END
                    
                    IF Characters < LEN(ResultColumnsQ.ColumnName) !If header text is bigger than the data
                        Characters = LEN(ResultColumnsQ.ColumnName) !Then we'll use the header text length for sizing the columns
                    END
                    ?SQLResultList{PROPLIST:width,ColumnNdx} = Characters * 8 !Pretty close to the correct width, usually
                END   
                0{PROP:Pixels} = FALSE
                UNHIDE(?SQLResultList)
                ?ResultPrompt{PROP:Text} = 'Result: ' & RECORDS(ResultQ) & ' rows.'  
                SELECT(?SQLQueryText)
                DISPLAY
            END
        END
        IF DBOpened 
            CLOSE(MyTurbo)
        END
        CLOSE(Window)
    
c.CheckError        PROCEDURE(STRING pMessage)

    CODE
!        
        IF ERRORCODE()
            MESSAGE(ERRORCODE() & ' ' & FILEERRORCODE() & '|' & ERROR() & '|' & FILEERROR(),pMessage) 
            RETURN ERRORCODE()
        END
        RETURN 0
        
c.InvalidateWindow  PROCEDURE

    CODE

    InValidateRect(INT(0{PROP:ClientHandle}),0,1)
        