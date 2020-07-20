Sub Golfmacro()
'
' golfmacro Macro
'TM need to be rewritten to use more functions/make it modular
'
' Keyboard Shortcut: Ctrl+Shift+G
'
' firstformat Macro
' Macro recorded 6/25/2006 by Billy, edited 7/1/2010 by Bob Linderman, edited by Dick Palmer 4/22/2012

    Windows("AGL-Individual-2019.xls").Activate
    
    'Start Page Setup
'    With ActiveSheet.PageSetup
'        .PrintTitleRows = ""
'        .PrintTitleColumns = ""
'    End With
    
    Application.MaxChange = 0.001
    With ActiveWorkbook
        .PrecisionAsDisplayed = False
        .SaveLinkValues = False
    End With
    'ActiveWindow.Zoom = 200
    'End Page Setup
    
    'Sort all players based on On#
    Rows("1:41").Select
    Selection.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
    
    'Copies substitute points from spreadsheet and updates individual ss
    Windows("AGL-Sub-points-2019.xls").Activate
    Range("D2:D41").Select
    Selection.Copy
    Windows("AGL-Individual-2019.xls").Activate
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Rows("1:41").Select
    Application.CutCopyMode = False
    
' Range("A2:AK41").Select
'   Selection.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlNo, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        DataOption1:=xlSortTextAsNumbers

' Transfer Macro
' Macro recorded 6/27/2006 by Billy
'TM 6/8/2020 Copies data from macro holder into indiv, creates a new sheet and pastes the data there
    
    Windows("AGL-Macro-Holder-2019.xlsm").Activate
    Range("A1:F11").Select
    Selection.Copy
    Windows("AGL-Individual-2019.xls").Activate
    Sheets("Sheet1").Select
    Sheets.Add
    Sheets("Sheet2").Select
    ActiveSheet.Paste
    Sheets("Sheet1").Name = "Details"
    Sheets("Sheet2").Name = "Standings"
    
'   Calculate Player Rank
        
    Windows("AGL-Sub-points-2019.xls").Activate
    
    'Defines the player rank field on the sub spreadsheet
    Range("E1").Value = "Player Rank"
    styleText Cells(1, "E"), fSize:=10
    
    'Loops through all players and calculates their rank
    For i = 1 To 40
        If Cells(i + 1, 3).Value = " " Then
            Cells(i + 1, 3).Value = 0
        End If
        Cells(i + 1, 5).Value = "=RANK(RC[-2],(R2C3:R41C3))"
    Next
    
    'Copies player rank list to individual list
    Range("E2:E41").Select
    Selection.Copy
    Windows("AGL-Individual-2019.xls").Activate
    Sheets("Details").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Rows("1:41").Select
    'Sorts first by team, then by their number in that team
    Selection.Sort Key1:=Range("C2"), Order1:=xlAscending, Key2:=Range("B2") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Columns("AN:AN").Select
    Selection.ColumnWidth = 3.5
    Columns("AS:AS").Select
    Selection.ColumnWidth = 3.5
    Rows("2:70").Select
    Selection.RowHeight = 9
    Rows("1:1").Select
    Selection.Rows.AutoFit
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    
    Cells(1, "K").Value = "DuPONT AUTOMOTIVE GOLF LEAGUE"
    styleText Cells(1, "K"), fSize:=9, fColor:=3
    styleText Cells(2, "K"), fSize:=8, fColor:=3

'  The following lines were added by Dick Palmer on 7/12/12 as User Input for Current Round and Date and
'      removed 11/21/12 - included in revised Avg_Pts_Rd macro

'   UserEntry = InputBox("Enter Current Event as: 'Round _ on 1/1/12'")
'   If UserEntry <> "" Then Info = UserEntry
'   ActiveCell.FormulaR1C1 = "as of Round 5 on 5/17/12"
'   ActiveCell.FormulaR1C1 = "As of " & Info

    Rows("2:2").Select
    Selection.RowHeight = 10
    Cells.Select
    Selection.NumberFormat = "General"
    Cells(1, "E").Value = "=TODAY()"
    styleText Cells(1, "E"), fSize:=8

' TeamNumber Macro
' Macro recorded 6/25/2006 by Billy, edited 7/1/2010 by Bob Linderman
' Rewritten by Tom M 2020
    For i = 4 To 70
        If Right(Cells(i, "B").Value, 1) = "0" Then
            Cells(i, "C").Value = 10
        Else
            Cells(i, "C").Value = "=RIGHT(RC[-1])"
        End If
    Next
     
'   Edited on 11-6-2019 to copy team names from Macro-Holder
    
    
    'Simplified by Tom M 6/8/2020 Removed redundant code for text styling
    'TM 6/8/20 Go through team names in macro holder and add them all to the indiv
    Dim counter As Integer
    counter = 2
    team = 1
    For i = 4 To 49 Step 5
    'Need to insert before copying otherwise excel fills in the new row with copies of that data
        Windows("AGL-Individual-2019.xls").Activate
        Rows(i + 4).Insert Shift:=xlDown
        Windows("AGL-Macro-Holder-2019.xlsm").Activate
        Cells(counter, 1).Copy
        Windows("AGL-Individual-2019.xls").Activate
        Cells(i + 4, 5).Select
        Cells(i + 4, 5).PasteSpecial
        counter = counter + 1
        CalculateSums i, team
        team = team + 1
        
    
        Set Bar = Worksheets("Details").Rows(i + 4)
        styleText Bar, fColor:=3
    
        Set Row = Cells(i + 4, 5)
        styleText Row, fSize:=6, fColor:=3
    
    Next
   
' CurrentPoints Macro
' Macro recorded 6/26/2006 by Billy, edited 7/1/2010 by Bob Linderman edited 6/9/20 by Tom Mroz
    Sheets("Standings").Select
    For i = 2 To 11
        Cells(i, 3).Value = "='[AGL-Individual-2019.xls]Details'!R" & 5 * (i - 1) + 3 & "C6"
    Next

    Columns(1).ColumnWidth = 25
    For i = 2 To 4
        Columns(i).ColumnWidth = 10
    Next
        
' Rankings Macro
' Macro recorded 6/25/2006 by Billy, edited 7/1/2010 by Bob Linderman, edited 6/9/20 by Tom Mroz

    Windows("AGL-Individual-2019.xls").Activate
    Sheets("Details").Select
    For i = 0 To 9
        Cells(i * 5 + 8, 1).Value = "='Standings'!R" & i + 2 & "C4"
    Next

    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
    Rows("54:54").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell

' Statsformat2 Macro
' Macro recorded 6/5/2006 by Billy, edited 7/1/2010 by Bob Linderman, edited by Dick Palmer 11/3/2019
    
    'Chooses where to paste the stats based on the number of subs
    'Goes 7 rows after the end of the subs
    i = 3
    
    Do While Cells(i, "E").Value <> ""
        i = i + 1
    Loop
    i = i + 5
    
    
    Cells(i, 1).Select
    Windows("AGL-Stats-2019.xls").Activate
    counter = 1
    Do Until Left(Cells(counter, "A").Value, 5) = "Event"
        counter = counter + 1
    Loop
    
    
    Range("A" & counter & ":O118").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("AGL-Individual-2019.xls").Activate
    ActiveSheet.Paste
    Rows("" & i & ":118").Select
    Application.CutCopyMode = False
    With Selection
        .Font.size = 8
        .Font.Name = "Arial"
        .Font.ColorIndex = xlAutomatic
        .HorizontalAlignment = xlLeft
    End With
       
    Range("D" & i & ":D118").Select
    Selection.Cut
    Range("F" & i).Select
    ActiveSheet.Paste
    Range("L" & i & ":L118").Select
    Selection.Cut
    Range("T" & i).Select
    ActiveSheet.Paste
    Range("H" & i & ":H118").Select
    Selection.Cut
    Range("L" & i).Select
    ActiveSheet.Paste

' formatbottom Macro
' Macro recorded 6/12/2006 by Billy, edited 7/1/2010 by Bob Linderman
    Range("C" & i & ":AB118").Select
'    Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Replace What:="Fewest Putts (0)", Replacement:="Team Low Net", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    ActiveWindow.ScrollRow = 1

' Macro written by Dick Palmer 4/5/2013 for addition of Column reporting Net Par results.

    Windows("AGL-Individual-2019.xls").Activate
    Sheets("Details").Select
    ActiveWindow.DisplayHeadings = True
    Columns("L:AB").Select
    Columns.EntireColumn.Hidden = False
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Range("J3").Select
    ActiveCell.Value = "Net + - Par"
    Selection.RowHeight = 27
    Range("AD4").Select
    
'    ActiveCell.Offset(1, 0).Select
    'Random for loop idek
    i = 0
    For i = 1 To 80
    
    Net = ActiveCell.Offset(0, -21).Value
    
    'TM 6/8/2020 simplified conditional if any other courses have a different par need to add to the conditional
    'Technically works but should be rewritten
        Tee = Trim(ActiveCell.Value)
            If ((Tee = "DWhtBck") Or (Tee = "DHybBck") Or (Tee = "DGldBck") _
                Or (Tee = "DGrnBck") Or (Tee = "NWhtFrt") Or (Tee = "NRedFrt")) _
            Then
                ActiveCell.Offset(0, -20) = (Net - 36)
            ElseIf (Tee = "None" Or Tee = "") Then
                ActiveCell.Value = " "
            Else
                ActiveCell.Offset(0, -20) = (Net - 35)
            End If
        
        ActiveCell.Offset(1, 0).Select
    
    Next i
    'Used to remove incorrect results (Player not playing etc)
    Range("J4").Select
    i = 0
    For i = 1 To 80
        If ActiveCell.Value < -25 Then ActiveCell.Value = ""
        ActiveCell.Offset(1, 0).Select
    Next i
    
    Range("J8").Select
    i = 0
    For i = 1 To 9
        'Calculates team net +/- par
        'maybe should be rewritten
        ActiveCell.FormulaR1C1 = "=Sum(R[-4]C:R[-1]C)"
        Selection.Copy
        ActiveCell.Offset(5, 0).Select
        ActiveSheet.Paste
    Next i
    
  ActiveCell.FormulaR1C1 = "=Sum(R[-4]C:R[-1]C)"
  
  ActiveWindow.DisplayHeadings = True
  Application.ScreenUpdating = True
End Sub
'Utility Functions below

Private Function styleText(ByVal cell As Object, Optional fSize As Integer = 7, Optional fColor As Integer = 0)
    'Used to format text with optional parameters from different options
    'Made to be as modular as possible
    'TM 6/9/20
    With cell
        .VerticalAlignment = xlTop
        .WrapText = False
        .ShrinkToFit = False
        .HorizontalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.ColorIndex = fColor
        .Font.size = fSize
        .Font.Bold = True
    End With

End Function

Private Function CalculateSums(ByVal start As Integer, ByVal team As Integer)
'TM 6/9/20
'Loops through stats and calculates the sum for each team
    For i = 7 To 11
        Cells(start + 4, i).Activate
        ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Next
    Cells(start + 4, 40).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    
    Cells(start + 4, 6).Activate
    ActiveCell.FormulaR1C1 = _
        "=R[-4]C+R[-3]C+R[-2]C+R[-1]C+RC[5]-'Standings'!R" & team + 1 & "C2"
End Function
