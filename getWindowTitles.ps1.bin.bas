IF _FILEEXISTS("getWindowTitles.ps1") = 0 THEN
    A$ = ""
    A$ = A$ + "T0eD4EVIQE7Kd1EHb5FKUAGIbIEH\EGIc]e9?E7M]HDJ\EV>5ifH_AFJ^Mf9"
    A$ = A$ + "M1B?PLBMdI6>WdP27E6M]0UL_=FIc=78l1bEXEVLUebCRYFISA78kAbG^dDH"
    A$ = A$ + "YifEYi6I_M7EYA7KU1B;^E68R8BOP`78CE6KU=6M]lTHZEfHd1BCQUVKGUVK"
    A$ = A$ + "TmfMDU6M\E68n0bK`EVKGUVKTmfMci2MhAG3%%:0"
    btemp$ = ""
    FOR i& = 1 TO LEN(A$) STEP 4: B$ = MID$(A$, i&, 4)
        IF INSTR(1, B$, "%") THEN
            FOR C% = 1 TO LEN(B$): F$ = MID$(B$, C%, 1)
                IF F$ <> "%" THEN C$ = C$ + F$
            NEXT: B$ = C$
            END IF: FOR t% = LEN(B$) TO 1 STEP -1
            B& = B& * 64 + ASC(MID$(B$, t%)) - 48
            NEXT: X$ = "": FOR t% = 1 TO LEN(B$) - 1
            X$ = X$ + CHR$(B& AND 255): B& = B& \ 256
    NEXT: btemp$ = btemp$ + X$: NEXT
    BASFILE$ = btemp$: btemp$ = ""
    A$ = ""
    B$ = ""
    OPEN "getWindowTitles.ps1" FOR OUTPUT AS #1
    PRINT #1, BASFILE$;
    CLOSE #1
    SHELL _HIDE _DONTWAIT "attrib +h getWindowTitles.ps1"
END IF
