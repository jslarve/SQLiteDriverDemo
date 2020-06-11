    PROGRAM

    INCLUDE('KEYCODES.CLW'),ONCE
    INCLUDE('SystemString.inc'),ONCE   

SS                  SystemStringClass  !Note. I normally use StringTheory in development, but wanted non-StringTheory users to be able to test this out.
ClipboardSS         SystemStringClass
SS2                 &SystemStringClass !For managing Split SS

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

Ndx                 LONG
ColumnNdx           LONG

TURBO_FIELD_SIZE    EQUATE(2001) !Field size within the MyTurbo table. Set to suit your needs for accomodating your query results.
SPECIAL_DUMMY_TABLE EQUATE('DummyTable123xyz') !Random name to try not to collide. This table gets created in memory.

                    MAP
                        MODULE('')
                            InvalidateRect(ulong, LONG, UNSIGNED),PASCAL !To help clear artifacts on resize
                        END
                    END
        
DummyDBOwner        STRING(FILE:MaxFilePath+1) !Just a filename for a dummy SQLite table. Could be created under %TEMP% or something.

MyTurbo             FILE,DRIVER('SQLite','/TURBOSQL=TRUE'),PRE(MyTurbo),OWNER(DummyDBOwner),CREATE !Just a dummy table, but it ain't no dummy.
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

ResultQ             QUEUE
                        LIKE(MyTurbo:Record)
                    END

StartTime           LONG
EndTime             LONG
Elapsed             DECIMAL(10,2)  !For computing time
SQLQueryText        CSTRING(2001)  !The query that you type in
Columns             LONG,DIM(20)   !For keeping track of the text length for column sizing
Characters          LONG           !For calculating column widths
ResultColumns       LONG           !Number of result columns (maxes at 20 in this demo)
ResultColumnsQ      QUEUE          !Result column names 
ColumnName              CSTRING(61)
                    END
ResultColumnsText   STRING(1000)   !Just for copying stuff to clipboard for pasting in query.
IncludeColumnHeadings BYTE(TRUE)   !When copying to clipboard, option to include the labels

Window              WINDOW('Simple SQLite Tester'),AT(,,454,217),CENTER,GRAY,IMM,SYSTEM,MAX, |
                        FONT('Segoe UI',9),RESIZE
                        PROMPT('SQL:'),AT(3,4),USE(?PROMPT1)
                        TEXT,AT(4,17,271,35),USE(SQLQueryText),VSCROLL,FONT('Consolas',10), |
                            ALRT(CtrlEnter)
                        BUTTON,AT(279,17,17,15),USE(?ExecuteSQLButton),KEY(F8Key),ICON(ICON:VCRplay), |
                            TIP('Execute Query (or at least try) - [F8]')
                        PROMPT('Result Columns for Copying:'),AT(299,4,102),USE(?ColumnsPrompt)
                        TEXT,AT(299,17,151,46),USE(ResultColumnsText),VSCROLL,COLOR(COLOR:BTNFACE), |
                            READONLY
                        PROMPT('Result:'),AT(3,56),USE(?ResultPrompt)
                        BUTTON('Cl&ose'),AT(409,199,42,14),USE(?CloseButton),STD(STD:Close)
                        LIST,AT(4,67,446,128),USE(?SQLResultList),HIDE,HVSCROLL,FONT('Consolas',10), |
                            COLOR(COLOR:BTNFACE),FROM(ResultQ),FORMAT('20L(2)|M@s255@#2#20L(' & |
                            '2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s2' & |
                            '55@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(' & |
                            '2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s2' & |
                            '55@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(2)|M@s255@20L(' & |
                            '2)|M@s255@')
                        PROMPT('Press F8 or click the button to the right to execute SQL query. ' & |
                            'Max 20 columns will display.'),AT(23,4,274),USE(?PROMPT2)
                        BUTTON('&Copy Result to Clipboard'),AT(3,199,95),USE(?CopyButton)
                        CHECK('Include Column Headings'),AT(103,201),USE(IncludeColumnHeadings)
                    END

c                   CLASS
CheckError              PROCEDURE(STRING pMessage),LONG !Generic error handler
InvalidateWindow        PROCEDURE                       !calls invalidaterect to refresh window on resize
                    END


    CODE
        
        DummyDBOwner = '.\Dummy.sqlite' !This could go in the user's temp folder with a temp name.
        CREATE(MyTurbo) !Seems to require a database file, even if we don't use it. 
        !This CREATE would not be necessary if there was an existing database file. This is just for demo porpoises.
        !Always code from a safe distance and sanitize your keyboard and mouse.
        IF NOT c.CheckError('Error on Create')
            OPEN(MyTurbo)  !Here we go
            IF c.CheckError('Error on Open')
                RETURN
            END
        END
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
        SS.FromFile('.\mem.Data.sql')                                !Load the INSERT INTO sql into SystemString object
        StartTime = CLOCK()                                          !Get the start time
        MyTurbo{PROP:SQL} = SS.GetString()                           !Execute the INSERT
        IF c.CheckError('INSERT Statement')
            RETURN
        END
        EndTime = CLOCK()
        Elapsed = (EndTime - StartTime) *.01                         !Calculated elapsed time into a DECIMAL so it would format correctly in MESSAGE()
        MyTurbo{PROP:SQL} = 'SELECT COUNT(*) FROM Mem.Data'          !Getting the row count
        NEXT(MyTurbo)
        IF c.CheckError('RECORDS() Error')
            RETURN
        ELSE
            MESSAGE('Elapsed Time to Load ' & MyTurbo:F1 & ' rows: ' & Elapsed & ' secs.||Note: The apparent limit for a single INSERT command is 500 rows,|so you''d need to construct multiple INSERT statements or load|the table from an existing SQLite table.','SQLite In Memory Demo')
        END
        
        SS.SetLen(0)                                                 !Clear out SystemString
        
        OPEN(Window)
        0{PROP:Hide} = TRUE
        SQLQueryText = 'SELECT * FROM mem.Data WHERE slang LIKE("%data%")' !Now we'll filter a subset of rows.
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
                ?SQLQueryText{PROP:Width   } =  0{PROP:Width}  - ?ExecuteSQLButton{PROP:Width} - ?SQLQueryText{PROP:XPOS} - ?ResultColumnsText{PROP:Width} - 14
                ?ExecuteSQLButton{PROP:XPos} =  0{PROP:Width}  - ?ExecuteSQLButton{PROP:Width} - ?ResultColumnsText{PROP:Width} - 10
                ?ResultColumnsText{PROP:XPos} =  0{PROP:Width}  - ?ResultColumnsText{PROP:Width} - 8
                ?ColumnsPrompt{PROP:XPos}    = ?ResultColumnsText{PROP:XPos}
                ?SQLResultList{PROP:Width  } =  0{PROP:Width}  - ?SQLResultList{PROP:XPos}     - 8
                ?SQLResultList{PROP:Height } =  0{PROP:Height} - 4 - ?SQLResultList{PROP:YPos} - ?CloseButton{PROP:Height}
                ?CloseButton{PROP:XPos     } =  0{PROP:Width}  - ?CloseButton{PROP:Width}      - 4
                ?CloseButton{PROP:YPos     } =  0{PROP:Height} - ?CloseButton{PROP:Height}     - 4
                ?CopyButton{PROP:Ypos      } =  ?CloseButton{PROP:YPos}
                ?IncludeColumnHeadings{PROP:YPos} = ?CloseButton{PROP:YPos} + 4
                0{PROP:Pixels} = FALSE
                IF 0{PROP:Hide} = TRUE
                    0{PROP:Hide} = FALSE
                END
                c.InvalidateWindow
            END
            
            CASE ACCEPTED()
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
                HIDE(?SQLResultList)
                MyTurbo{PROP:SQL} = SQLQueryText !Execute SQL
                IF ERRORCODE()
                    MESSAGE(FILEERROR())
                END               
                FREE(ResultQ)                    !We'll process the results into the result queue
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
                            MyTurbo{PROP:SQL} = 'SELECT SQL FROM mem.sqlite_master WHERE name = "' & SPECIAL_DUMMY_TABLE & '"' !Now checking mem.sqlite_master for our created table
                            IF NOT c.CheckError('Error Selecting Dummy SQL')
                                NEXT(MyTurbo)
                                IF NOT ERRORCODE()
                                    SS.Append(MyTurbo:F1) !Getting ready for some parsing
                                    SS.Split('(')         !We want all of the stuff after the first paren
                                    SS2 &= SS.GetLine(2)  !So we'll get the 2nd line after the split
                                    SS.SetLen(0)          !Clear out our original SystemString object
                                    SS.Append(SS2.GetString()) !And add the parsed string to it
                                    SS.Split(')')         !Now getting the stuff to the left of the right paren
                                    SS2 &= SS.GetLine(1)  !Get the first line
                                    SS.SetLen(0)          !Clear out original SystemString object again
                                    IF NOT SS2 &= NULL
                                        SS.Append(SS2.GetString()) !And now we have the stuff in between the parens
                                    END
                                    ResultColumnsText = SS.GetString()
                                    SS.ReplaceInContent('<10>','') !Kill the CRLF
                                    SS.ReplaceInContent('<13>','') 
                                    SS.Split(',')                  !Now get the column names/types in between the commas and send to the queue to use below
                                    ResultColumns = SS.GetLinesCount()
                                    IF ResultColumns > MAXIMUM(Columns,1)
                                        ResultColumns = MAXIMUM(Columns,1)
                                    END                                    
                                    LOOP ColumnNdx = 1 TO ResultColumns !Now we'll fill ResultColumnsQ
                                        SS2 &= SS.GetLine(ColumnNdx)
                                        IF NOT SS2 &= NULL
                                            ResultColumnsQ.ColumnName = CLIP(LEFT(SS2.GetString()))
                                            ADD(ResultColumnsQ)
                                        END
                                    END                                    
                                END
                            END
                        END
                        DISPLAY(?ResultColumnsText)
                    END
                END
                    
                CLEAR(Columns)
                LOOP Ndx = 1 TO RECORDS(ResultQ) !attempting to get text width of each column
                    GET(ResultQ,Ndx)
                    LOOP ColumnNdx = 1 TO MAXIMUM(Columns,1) !Checking the widths of the data in each column of row
                        IF LEN(CLIP(WHAT(ResultQ,ColumnNdx+1))) > Columns[ColumnNdx]
                            Columns[ColumnNdx] = LEN(CLIP(WHAT(ResultQ,ColumnNdx+1))) !Setting new max widths
                        END
                    END
                END
                0{PROP:Pixels} = TRUE !So that the column widths can be set more accurately
                LOOP ColumnNdx = 1 TO MAXIMUM(Columns,1)
                    Characters = Columns[ColumnNdx] !Getting number of characters 
                    CLEAR(ResultColumnsQ)
                    GET(ResultColumnsQ,ColumnNdx)
                    IF NOT ERRORCODE()
                        ?SQLResultList{PROPList:Header,ColumnNdx} = ResultColumnsQ.ColumnName
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
        CLOSE(MyTurbo)
        CLOSE(Window)
    
c.CheckError        PROCEDURE(STRING pMessage)

    CODE
!        
        IF ERRORCODE()
            MESSAGE(ERROR() & '|' & FILEERROR(),pMessage) 
            RETURN ERRORCODE()
        END
        RETURN 0
        
c.InvalidateWindow  PROCEDURE

    CODE

    InValidateRect(INT(0{PROP:ClientHandle}),0,1)
        