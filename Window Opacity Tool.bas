': This program uses
': InForm - GUI library for QB64 - Beta version 8
': Fellippe Heitor, 2016-2018 - fellippe@qb64.org - @fellippeheitor
': https://github.com/FellippeHeitor/InForm
'-----------------------------------------------------------
$VERSIONINFO:CompanyName=Zachary Spriggs
$VERSIONINFO:FILEVERSION#=1,0,0,0
$VERSIONINFO:ProductName=Window Opacity Tool
$VERSIONINFO:LegalCopyright=(c)2019 Zachary Spriggs
$VERSIONINFO:Comments=Sets opacity level of windows using PowerShell commands
$VERSIONINFO:FileDescription=Window Opacity Tool

': Controls' IDs: ------------------------------------------------------------------
DIM SHARED WindowOpacityTool AS LONG
DIM SHARED ListBox1 AS LONG
DIM SHARED TrackBar1 AS LONG
DIM SHARED valLB AS LONG
DIM SHARED ApplyBT AS LONG
DIM SHARED LoadOpenWindowsInfoBT AS LONG
DIM SHARED ProgressBar1 AS LONG
DIM SHARED DontSeeYourWindowTitleEnterItLB AS LONG
DIM SHARED WindowTitleTB AS LONG
DECLARE DYNAMIC LIBRARY "user32"
    FUNCTION FindWindowA%& (BYVAL lpClassName%&, BYVAL lpWindowName%&)
    FUNCTION LoadIconA%& (BYVAL hInstance%&, BYVAL lpIconName%&)
    FUNCTION SetLayeredWindowAttributes& (BYVAL hwnd AS LONG, BYVAL crKey AS LONG, BYVAL bAlpha AS _UNSIGNED _BYTE, BYVAL dwFlags AS LONG)
    FUNCTION GetWindowLong& ALIAS "GetWindowLongA" (BYVAL hwnd AS LONG, BYVAL nIndex AS LONG)
    FUNCTION SetWindowLong& ALIAS "SetWindowLongA" (BYVAL hwnd AS LONG, BYVAL nIndex AS LONG, BYVAL dwNewLong AS LONG)
    FUNCTION GetWindowLongA& (BYVAL hwnd AS LONG, BYVAL nIndex AS LONG)
    FUNCTION SetWindowLongA& (BYVAL hwnd AS LONG, BYVAL nIndex AS LONG, BYVAL dwNewLong AS LONG)
    FUNCTION SetWindowPos& (BYVAL hwnd AS LONG, BYVAL hWndInsertAfter AS LONG, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL wFlags AS LONG)
END DECLARE
DECLARE LIBRARY
    FUNCTION FindWindow& (BYVAL ClassName AS _OFFSET, WindowName$) ' To get hWnd handle
END DECLARE
DIM MyHwnd AS LONG
DIM SHARED WindowVal
DIM SHARED opacity
$EXEICON:'Nishad2m8-Hologram-Dock-Windows.ico'
DECLARE DYNAMIC LIBRARY "kernel32"
    FUNCTION GetLastError~& ()
END DECLARE
': External modules: ---------------------------------------------------------------
'$INCLUDE:'InForm\InForm.ui'
'$INCLUDE:'InForm\xp.uitheme'
'$INCLUDE:'Window Opacity Tool.frm'

': Event procedures: ---------------------------------------------------------------
SUB __UI_BeforeInit
    IF _FILEEXISTS("policycheck.txt") = 0 THEN 'Checks if policycheck.txt exists
        SHELL _HIDE "PowerShell -NoProfile -ExecutionPolicy Bypass -Command " + CHR$(34) + "& {Start-Process PowerShell -ArgumentList 'Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force' -Verb RunAs}" + CHR$(34) 'Sets PowerShell script execution policy to remotesigned
        OPEN "policycheck.txt" FOR OUTPUT AS #10 'Opens policycheck.txt to read in data
        PRINT #10, "POLICY SET @ " + Clock$ + " ON " + DATE$ 'Prints a time and date stamp in the file policycheck.txt
        CLOSE #10
        SHELL _HIDE _DONTWAIT "attrib +h policycheck.txt" 'Hides the file from user
    END IF
    IF _FILEEXISTS("getWindowTitles.ps1") = 0 THEN
        OPEN "getWindowTitles.ps1" FOR OUTPUT AS #1
        PRINT #1, "$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'"
        PRINT #1, "Get-Process | Where-Object {$_.MainWindowTitle -ne " + CHR$(34) + CHR$(34) + "} | Select-Object MainWindowTitle > openWindows.txt"
        CLOSE #1
        SHELL _HIDE _DONTWAIT "attrib +h getWindowTitles.ps1" 'Hides the file from user
    END IF
    SHELL _HIDE "powershell " + CHR$(34) + "&" + CHR$(34) + CHR$(34) + _STARTDIR$ + "\getWindowTitles.ps1" + CHR$(34) + CHR$(34) + CHR$(34)
END SUB

SUB __UI_OnLoad
    SetFrameRate 60
    _TITLE "Window Opacity Tool"
    OPEN "openWindows.txt" FOR BINARY AS #1
    ResetList ListBox1
    IF EOF(1) = 0 THEN
        LINE INPUT #1, foo$
        LINE INPUT #1, foo$
        LINE INPUT #1, foo$
        DO
            LINE INPUT #1, Window$
            IF Window$ <> "" THEN
                FOR i = 0 TO 31
                    IF INSTR(Window$, CHR$(i)) THEN
                        Window$ = ReplaceStringItem$(Window$, CHR$(i), "")
                    END IF
                NEXT
                FOR i = 127 TO 255
                    IF INSTR(Window$, CHR$(i)) THEN
                        Window$ = ReplaceStringItem$(Window$, CHR$(i), "")
                    END IF
                NEXT
                AddItem ListBox1, Window$
            END IF
        LOOP UNTIL EOF(1)
    END IF
    CLOSE #1
    Control(TrackBar1).Value = 255
END SUB

SUB __UI_BeforeUpdateDisplay
    'This event occurs at approximately 30 frames per second.
    'You can change the update frequency by calling SetFrameRate DesiredRate%

END SUB

SUB __UI_BeforeUnload
    'If you set __UI_UnloadSignal = False here you can
    'cancel the user's request to close.

END SUB

SUB __UI_Click (id AS LONG)
    SELECT CASE id
        CASE WindowOpacityTool

        CASE ListBox1

        CASE TrackBar1

        CASE valLB

        CASE ApplyBT
            IF VAL(Caption(valLB)) <= 40 THEN
                Answer = MessageBox("Values 40 and below will be very hard to see.\nAre you sure you wish to continue?", "NOTICE", MsgBox_YesNo + MsgBox_Information)
                IF Answer = MsgBox_Yes THEN
                    IF Text(WindowTitleTB) <> "" THEN
                        Window$ = Text(WindowTitleTB)
                        FOR x = 1 TO 10
                            MyHwnd = FindWindow(0, Window$ + CHR$(0))
                            x = x + 1
                            Control(ProgressBar1).Value = x * 10
                            _DELAY 0.25
                        NEXT
                        IF MyHwnd THEN
                            SetWindowOpacity MyHwnd, opacity
                            Answer = MessageBox("Opacity set to" + STR$(opacity) + " for window : " + Window$, "Opacity Successfully Set!", MsgBox_OkOnly + MsgBox_Information)
                            Control(ProgressBar1).Value = 0
                        ELSE
                            Answer = MessageBox("Failed to change opacity", "Couldn't change opacity", MsgBox_OkOnly + MsgBox_Exclamation)
                            Control(ProgressBar1).Value = 0
                        END IF
                    ELSE
                        IF Control(ListBox1).Value <> 0 THEN
                            Window$ = GetItem$(ListBox1, WindowVal)
                            Window$ = LTRIM$(Window$)
                            Window$ = RTRIM$(Window$)
                            FOR x = 1 TO 10
                                MyHwnd = FindWindow(0, Window$ + CHR$(0))
                                x = x + 1
                                Control(ProgressBar1).Value = x * 10
                                _DELAY 0.25
                            NEXT
                            IF MyHwnd THEN
                                SetWindowOpacity MyHwnd, opacity
                                Answer = MessageBox("Opacity set to" + STR$(opacity) + " for window : " + Window$, "Opacity Successfully Set!", MsgBox_OkOnly + MsgBox_Information)
                                Control(ProgressBar1).Value = 0
                            ELSE
                                Answer = MessageBox("Failed to change opacity", "Couldn't change opacity", MsgBox_OkOnly + MsgBox_Exclamation)
                                Control(ProgressBar1).Value = 0
                            END IF
                        END IF
                    END IF
                END IF
            ELSE

                IF Text(WindowTitleTB) <> "" THEN
                    Window$ = Text(WindowTitleTB)
                    FOR x = 1 TO 10
                        MyHwnd = FindWindow(0, Window$ + CHR$(0))
                        x = x + 1
                        Control(ProgressBar1).Value = x * 10
                        _DELAY 0.25
                    NEXT
                    IF MyHwnd THEN
                        SetWindowOpacity MyHwnd, opacity
                        Answer = MessageBox("Opacity set to" + STR$(opacity) + " for window : " + Window$, "Opacity Successfully Set!", MsgBox_OkOnly + MsgBox_Information)
                        Control(ProgressBar1).Value = 0
                    ELSE
                        Answer = MessageBox("Failed to change opacity", "Couldn't change opacity", MsgBox_OkOnly + MsgBox_Exclamation)
                        Control(ProgressBar1).Value = 0
                    END IF
                ELSE
                    IF Control(ListBox1).Value <> 0 THEN
                        Window$ = GetItem$(ListBox1, WindowVal)
                        Window$ = LTRIM$(Window$)
                        Window$ = RTRIM$(Window$)
                        FOR x = 1 TO 10
                            MyHwnd = FindWindow(0, Window$ + CHR$(0))
                            x = x + 1
                            Control(ProgressBar1).Value = x * 10
                            _DELAY 0.25
                        NEXT
                        IF MyHwnd THEN
                            SetWindowOpacity MyHwnd, opacity
                            Answer = MessageBox("Opacity set to" + STR$(opacity) + " for window : " + Window$, "Opacity Successfully Set!", MsgBox_OkOnly + MsgBox_Information)
                            Control(ProgressBar1).Value = 0
                        ELSE
                            Answer = MessageBox("Failed to change opacity", "Couldn't change opacity", MsgBox_OkOnly + MsgBox_Exclamation)
                            Control(ProgressBar1).Value = 0
                        END IF
                    END IF
                END IF
            END IF
        CASE LoadOpenWindowsInfoBT
            Text(WindowTitleTB) = ""
            SHELL _HIDE "powershell " + CHR$(34) + "&" + CHR$(34) + CHR$(34) + _STARTDIR$ + "\getWindowTitles.ps1" + CHR$(34) + CHR$(34) + CHR$(34)
            OPEN "openWindows.txt" FOR BINARY AS #1
            ResetList ListBox1
            IF EOF(1) = 0 THEN
                LINE INPUT #1, foo$
                LINE INPUT #1, foo$
                LINE INPUT #1, foo$
                DO
                    LINE INPUT #1, Window$
                    IF Window$ <> "" THEN
                        FOR i = 0 TO 31
                            IF INSTR(Window$, CHR$(i)) THEN
                                Window$ = ReplaceStringItem$(Window$, CHR$(i), "")
                            END IF
                        NEXT
                        FOR i = 127 TO 255
                            IF INSTR(Window$, CHR$(i)) THEN
                                Window$ = ReplaceStringItem$(Window$, CHR$(i), "")
                            END IF
                        NEXT
                        AddItem ListBox1, Window$
                    END IF
                LOOP UNTIL EOF(1)
            END IF
            CLOSE #1
        CASE ProgressBar1
    END SELECT
END SUB

SUB __UI_MouseEnter (id AS LONG)
    SELECT CASE id
        CASE WindowOpacityTool

        CASE ListBox1

        CASE TrackBar1

        CASE valLB

        CASE ApplyBT

        CASE LoadOpenWindowsInfoBT

        CASE ProgressBar1

    END SELECT
END SUB

SUB __UI_MouseLeave (id AS LONG)
    SELECT CASE id
        CASE WindowOpacityTool

        CASE ListBox1

        CASE TrackBar1

        CASE valLB

        CASE ApplyBT

        CASE LoadOpenWindowsInfoBT

        CASE ProgressBar1

    END SELECT
END SUB

SUB __UI_FocusIn (id AS LONG)
    SELECT CASE id
        CASE ListBox1

        CASE TrackBar1

        CASE ApplyBT

        CASE LoadOpenWindowsInfoBT

    END SELECT
END SUB

SUB __UI_FocusOut (id AS LONG)
    'This event occurs right before a control loses focus.
    'To prevent a control from losing focus, set __UI_KeepFocus = True below.
    SELECT CASE id
        CASE ListBox1

        CASE TrackBar1

        CASE ApplyBT

        CASE LoadOpenWindowsInfoBT

    END SELECT
END SUB

SUB __UI_MouseDown (id AS LONG)
    SELECT CASE id
        CASE WindowOpacityTool

        CASE ListBox1

        CASE TrackBar1

        CASE valLB

        CASE ApplyBT

        CASE LoadOpenWindowsInfoBT

        CASE ProgressBar1

    END SELECT
END SUB

SUB __UI_MouseUp (id AS LONG)
    SELECT CASE id
        CASE WindowOpacityTool

        CASE ListBox1

        CASE TrackBar1

        CASE valLB

        CASE ApplyBT

        CASE LoadOpenWindowsInfoBT

        CASE ProgressBar1

    END SELECT
END SUB

SUB __UI_KeyPress (id AS LONG)
    'When this event is fired, __UI_KeyHit will contain the code of the key hit.
    'You can change it and even cancel it by making it = 0
    SELECT CASE id
        CASE ListBox1

        CASE TrackBar1
            IF __UI_KeyHit = 13 THEN
                __UI_Click (ApplyBT)
            END IF
        CASE ApplyBT
            IF __UI_KeyHit = 13 THEN
                __UI_Click (ApplyBT)
            END IF
        CASE LoadOpenWindowsInfoBT

    END SELECT
END SUB

SUB __UI_TextChanged (id AS LONG)
    SELECT CASE id
        CASE WindowTitleTB
    END SELECT
END SUB

SUB __UI_ValueChanged (id AS LONG)
    SELECT CASE id
        CASE ListBox1
            WindowVal = Control(ListBox1).Value
            Control(TrackBar1).Value = 255
            IF WindowVal <> 0 THEN
                ToolTip(ListBox1) = tip$
                Text(WindowTitleTB) = ""
            ELSE
                ToolTip(ListBox1) = ""
            END IF
        CASE TrackBar1
            Caption(valLB) = STR$(Control(TrackBar1).Value)
            opacity = Control(TrackBar1).Value
    END SELECT
END SUB

SUB __UI_FormResized

END SUB
SUB SetWindowOpacity (hWnd AS LONG, Level)
    DIM Msg AS LONG
    CONST G = -20
    CONST LWA_ALPHA = &H2
    CONST WS_EX_LAYERED = &H80000

    Msg = GetWindowLong(hWnd, G)
    Msg = Msg OR WS_EX_LAYERED
    Crap = SetWindowLong(hWnd, G, Msg)
    Crap = SetLayeredWindowAttributes(hWnd, 0, Level, LWA_ALPHA)
END SUB

FUNCTION ReplaceStringItem$ (text$, old$, new$)
    DO
        find = INSTR(start + 1, text$, old$) 'find location of a word in text
        IF find THEN
            first$ = LEFT$(text$, find - 1) 'text before word including spaces
            last$ = RIGHT$(text$, LEN(text$) - (find + LEN(old$) - 1)) 'text after word
            text$ = first$ + new$ + last$
        END IF
        start = find
    LOOP WHILE find
    ReplaceStringItem$ = text$
END FUNCTION
FUNCTION Clock$
    hour$ = LEFT$(TIME$, 2): H% = VAL(hour$)
    min$ = MID$(TIME$, 3, 3)
    IF H% >= 12 THEN ampm$ = " PM" ELSE ampm$ = " AM"
    IF H% > 12 THEN
        IF H% - 12 < 10 THEN hour$ = STR$(H% - 12) ELSE hour$ = LTRIM$(STR$(H% - 12))
    ELSEIF H% = 0 THEN hour$ = "12" ' midnight hour
    ELSE: IF H% < 10 THEN hour$ = STR$(H%) ' eliminate leading zeros
    END IF
    Clock$ = hour$ + min$ + ampm$
END FUNCTION

