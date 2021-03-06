': This program uses
': InForm - GUI library for QB64 - Beta version 8
': Fellippe Heitor, 2016-2018 - fellippe@qb64.org - @fellippeheitor
': https://github.com/FellippeHeitor/InForm
'-----------------------------------------------------------
$VERSIONINFO:CompanyName=SpriggsySpriggs
$VERSIONINFO:FILEVERSION#=1,0,0,0
$VERSIONINFO:ProductName=Window Opacity Tool
$VERSIONINFO:LegalCopyright=(c)2020 SpriggsySpriggs
$VERSIONINFO:Comments=Sets opacity level of windows using PowerShell commands
$VERSIONINFO:FileDescription=Window Opacity Tool

': Controls' IDs: ------------------------------------------------------------------
DIM SHARED WindowOpacityTool AS LONG
DIM SHARED ListBox1 AS LONG
DIM SHARED TrackBar1 AS LONG
DIM SHARED valLB AS LONG
DIM SHARED ApplyBT AS LONG
DIM SHARED LoadOpenWindowsInfoBT AS LONG
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
_ICON

DECLARE DYNAMIC LIBRARY "kernel32"
    FUNCTION GetLastError~& ()
END DECLARE

': External modules: ---------------------------------------------------------------
'$INCLUDE:'WPFMessageBox.BI'
'$INCLUDE:'getWindowTitles.ps1.bin.bas'
'$INCLUDE:'InForm\InForm.ui'
'$INCLUDE:'InForm\xp.uitheme'
'$INCLUDE:'Window Opacity Tool.frm'

': Event procedures: ---------------------------------------------------------------
SUB __UI_BeforeInit
    SHELL _HIDE "PowerShell -ExecutionPolicy Bypass -Command " + CHR$(34) + "&" + CHR$(34) + CHR$(34) + _STARTDIR$ + "\getWindowTitles.ps1" + CHR$(34) + CHR$(34) + CHR$(34)
END SUB

SUB __UI_OnLoad
    SetFrameRate 60
    _TITLE "Window Opacity Tool"
    OPEN "openWindows.txt" FOR BINARY AS #1
    ResetList ListBox1
    IF EOF(1) = 0 THEN
        DO
            LINE INPUT #1, Window$
            IF Window$ <> "" THEN
                FOR i = 0 TO 31
                    IF INSTR(Window$, CHR$(i)) THEN
                        Window$ = Remove(Window$, CHR$(i))
                    END IF
                NEXT
                FOR i = 127 TO 255
                    IF INSTR(Window$, CHR$(i)) THEN
                        Window$ = Remove(Window$, CHR$(i))
                    END IF
                NEXT
                AddItem ListBox1, Window$
            END IF
        LOOP UNTIL EOF(1)
    END IF
    CLOSE #1
    Control(TrackBar1).Value = 255
    _SCREENMOVE _MIDDLE
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
                Params.TitleBackground = Hue.Yellow
                Params.Title = "NOTICE"
                Params.Content = "Values 40 and below will be very hard to see.\nAre you sure you wish to continue?"
                Params.TitleFontSize = 28
                Params.TitleFontWeight = FontWeight.Bold
                Params.ContentFontSize = 18
                Params.ContentFontWeight = FontWeight.Medium
                Params.ButtonType = Button.Yes_No
                Params.ContentBackground = Hue.LightYellow
                Params.Sound = Alert.Exclamation
                Answer = WPFMessageBox(Params)
                'Answer = MessageBox("Values 40 and below will be very hard to see.\nAre you sure you wish to continue?", "NOTICE", MsgBox_YesNo + MsgBox_Information)
                IF Answer = ButtonClicked.Yes THEN 'MsgBox_Yes THEN
                    IF Text(WindowTitleTB) <> "" THEN
                        Window$ = Text(WindowTitleTB)
                        'FOR x = 1 TO 10
                        IF Window$ = _TITLE$ THEN
                            MyHwnd = _WINDOWHANDLE
                        ELSE
                            MyHwnd = FindWindow(0, Window$ + CHR$(0))
                        END IF
                        IF MyHwnd THEN
                            SetWindowOpacity MyHwnd, opacity
                            Params.TitleBackground = Hue.Green
                            Params.Title = "Opacity Successfully Set!"
                            Params.Content = "Opacity set to " + LTRIM$(STR$(opacity)) + " for window :\n" + Window$
                            Params.TitleFontSize = 28
                            Params.TitleFontWeight = FontWeight.Bold
                            Params.TitleTextForeground = Hue.White
                            Params.ContentFontSize = 18
                            Params.ContentFontWeight = FontWeight.Medium
                            Params.ContentBackground = Hue.MintCream
                            Answer = WPFMessageBox(Params)
                            'Answer = MessageBox("Opacity set to" + STR$(opacity) + " for window : " + Window$, "Opacity Successfully Set!", MsgBox_OkOnly + MsgBox_Information)
                            'Control(ProgressBar1).Value = 0
                        ELSE
                            Params.TitleBackground = Hue.Red
                            Params.Title = "Couldn't change opacity"
                            Params.Content = "Failed to change opacity"
                            Params.TitleFontSize = 28
                            Params.TitleFontWeight = FontWeight.Bold
                            Params.ContentFontSize = 18
                            Params.ContentFontWeight = FontWeight.Medium
                            Params.ContentBackground = Hue.MintCream
                            Params.Sound = Alert.Error
                            Answer = WPFMessageBox(Params)
                            'Answer = MessageBox("Failed to change opacity", "Couldn't change opacity", MsgBox_OkOnly + MsgBox_Exclamation)
                            'Control(ProgressBar1).Value = 0
                        END IF
                    ELSE
                        IF Control(ListBox1).Value <> 0 THEN
                            Window$ = GetItem$(ListBox1, WindowVal)
                            Window$ = LTRIM$(Window$)
                            Window$ = RTRIM$(Window$)
                            'FOR x = 1 TO 10
                            IF Window$ = _TITLE$ THEN
                                MyHwnd = _WINDOWHANDLE
                            ELSE
                                MyHwnd = FindWindow(0, Window$ + CHR$(0))
                            END IF
                            IF MyHwnd THEN
                                SetWindowOpacity MyHwnd, opacity
                                Params.TitleBackground = Hue.Green
                                Params.Title = "Opacity Successfully Set!"
                                Params.Content = "Opacity set to " + LTRIM$(STR$(opacity)) + " for window :\n" + Window$
                                Params.TitleFontSize = 28
                                Params.TitleFontWeight = FontWeight.Bold
                                Params.TitleTextForeground = Hue.White
                                Params.ContentFontSize = 18
                                Params.ContentFontWeight = FontWeight.Medium
                                Params.ContentBackground = Hue.MintCream
                                Answer = WPFMessageBox(Params)
                                'Answer = MessageBox("Opacity set to" + STR$(opacity) + " for window : " + Window$, "Opacity Successfully Set!", MsgBox_OkOnly + MsgBox_Information)
                                'Control(ProgressBar1).Value = 0
                            ELSE
                                Params.TitleBackground = Hue.Red
                                Params.Title = "Couldn't change opacity"
                                Params.Content = "Failed to change opacity"
                                Params.TitleFontSize = 28
                                Params.TitleFontWeight = FontWeight.Bold
                                Params.ContentFontSize = 18
                                Params.ContentFontWeight = FontWeight.Medium
                                Params.ContentBackground = Hue.MintCream
                                Params.Sound = Alert.Error
                                Answer = WPFMessageBox(Params)
                                'Answer = MessageBox("Failed to change opacity", "Couldn't change opacity", MsgBox_OkOnly + MsgBox_Exclamation)
                                'Control(ProgressBar1).Value = 0
                            END IF
                        END IF
                    END IF
                END IF
            ELSE

                IF Text(WindowTitleTB) <> "" THEN
                    Window$ = Text(WindowTitleTB)
                    'FOR x = 1 TO 10
                    IF Window$ = _TITLE$ THEN
                        MyHwnd = _WINDOWHANDLE
                    ELSE
                        MyHwnd = FindWindow(0, Window$ + CHR$(0))
                    END IF
                    'x = x + 1
                    'Control(ProgressBar1).Value = x * 10
                    '_DELAY 0.25
                    'NEXT
                    IF MyHwnd THEN
                        SetWindowOpacity MyHwnd, opacity
                        Params.TitleBackground = Hue.Green
                        Params.Title = "Opacity Successfully Set!"
                        Params.Content = "Opacity set to " + LTRIM$(STR$(opacity)) + " for window :\n" + Window$
                        Params.TitleFontSize = 28
                        Params.TitleFontWeight = FontWeight.Bold
                        Params.TitleTextForeground = Hue.White
                        Params.ContentFontSize = 18
                        Params.ContentFontWeight = FontWeight.Medium
                        Params.ContentBackground = Hue.MintCream
                        Answer = WPFMessageBox(Params)
                        'Answer = MessageBox("Opacity set to" + STR$(opacity) + " for window : " + Window$, "Opacity Successfully Set!", MsgBox_OkOnly + MsgBox_Information)
                        'Control(ProgressBar1).Value = 0
                    ELSE
                        Params.TitleBackground = Hue.Red
                        Params.Title = "Couldn't change opacity"
                        Params.Content = "Failed to change opacity"
                        Params.TitleFontSize = 28
                        Params.TitleFontWeight = FontWeight.Bold
                        Params.ContentFontSize = 18
                        Params.ContentFontWeight = FontWeight.Medium
                        Params.ContentBackground = Hue.MintCream
                        Params.Sound = Alert.Error
                        Answer = WPFMessageBox(Params)
                        'Answer = MessageBox("Failed to change opacity", "Couldn't change opacity", MsgBox_OkOnly + MsgBox_Exclamation)
                        'Control(ProgressBar1).Value = 0
                    END IF
                ELSE
                    IF Control(ListBox1).Value <> 0 THEN
                        Window$ = GetItem$(ListBox1, WindowVal)
                        Window$ = LTRIM$(Window$)
                        Window$ = RTRIM$(Window$)
                        'FOR x = 1 TO 10
                        IF Window$ = _TITLE$ THEN
                            MyHwnd = _WINDOWHANDLE
                        ELSE
                            MyHwnd = FindWindow(0, Window$ + CHR$(0))
                        END IF
                        'x = x + 1
                        'Control(ProgressBar1).Value = x * 10
                        '_DELAY 0.25
                        'NEXT
                        IF MyHwnd THEN
                            SetWindowOpacity MyHwnd, opacity
                            Params.TitleBackground = Hue.Green
                            Params.Title = "Opacity Successfully Set!"
                            Params.Content = "Opacity set to " + LTRIM$(STR$(opacity)) + " for window :\n" + Window$
                            Params.TitleFontSize = 28
                            Params.TitleFontWeight = FontWeight.Bold
                            Params.TitleTextForeground = Hue.White
                            Params.ContentFontSize = 18
                            Params.ContentFontWeight = FontWeight.Medium
                            Params.ContentBackground = Hue.MintCream
                            Answer = WPFMessageBox(Params)
                            'Answer = MessageBox("Opacity set to" + STR$(opacity) + " for window : " + Window$, "Opacity Successfully Set!", MsgBox_OkOnly + MsgBox_Information)
                            'Control(ProgressBar1).Value = 0
                        ELSE
                            Params.TitleBackground = Hue.Red
                            Params.Title = "Couldn't change opacity"
                            Params.Content = "Failed to change opacity"
                            Params.TitleFontSize = 28
                            Params.TitleFontWeight = FontWeight.Bold
                            Params.ContentFontSize = 18
                            Params.ContentFontWeight = FontWeight.Medium
                            Params.ContentBackground = Hue.MintCream
                            Params.Sound = Alert.Error
                            Answer = WPFMessageBox(Params)
                            'Answer = MessageBox("Failed to change opacity", "Couldn't change opacity", MsgBox_OkOnly + MsgBox_Exclamation)
                            'Control(ProgressBar1).Value = 0
                        END IF
                    END IF
                END IF
            END IF
        CASE LoadOpenWindowsInfoBT
            Text(WindowTitleTB) = ""
            SHELL _HIDE "PowerShell -ExecutionPolicy Bypass -Command " + CHR$(34) + "&" + CHR$(34) + CHR$(34) + _STARTDIR$ + "\getWindowTitles.ps1" + CHR$(34) + CHR$(34) + CHR$(34)
            OPEN "openWindows.txt" FOR BINARY AS #1
            ResetList ListBox1
            IF EOF(1) = 0 THEN
                DO
                    LINE INPUT #1, Window$
                    IF Window$ <> "" THEN
                        FOR i = 0 TO 31
                            IF INSTR(Window$, CHR$(i)) THEN
                                Window$ = Remove(Window$, CHR$(i))
                            END IF
                        NEXT
                        FOR i = 127 TO 255
                            IF INSTR(Window$, CHR$(i)) THEN
                                Window$ = Remove(Window$, CHR$(i))
                            END IF
                        NEXT
                        AddItem ListBox1, Window$
                    END IF
                LOOP UNTIL EOF(1)
            END IF
            CLOSE #1
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
    a = SetWindowLong(hWnd, G, Msg)
    'a = SetLayeredWindowAttributes(hWnd, &HFFFFFF, Level, &H1)
    a = SetLayeredWindowAttributes(hWnd, 0, Level, &H2)
END SUB

FUNCTION Remove$ (text$, old$)
    DIM find
    DIM start
    DIM first$
    DIM last$
    DIM new$
    new$ = ""
    DO
        find = INSTR(start + 1, text$, old$) 'find location of a word in text
        IF find THEN
            first$ = LEFT$(text$, find - 1) 'text before word including spaces
            last$ = RIGHT$(text$, LEN(text$) - (find + LEN(old$) - 1)) 'text after word
            text$ = first$ + new$ + last$
        END IF
        start = find
    LOOP WHILE find
    Remove = text$
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

'$INCLUDE:'WPFMessageBox.BM'
