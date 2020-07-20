
Sub AveragePointsPerRound()

' AveragePointsPerRound Macro
' Calculates the average points per round to determine player and team rankings; created by Bob Linderman, 7/1/2010
' Macro works for 10 teams and 4 players per team
' Assumes maximum number of substitute players is 30

' Macro revised by Dick Palmer 3/31/2013 from "Low Net" metric to "Net + - Par".
' Revised further on 11/6/2013 to track Player Low Net Weekly and YTD

' Keyboard Shortcut: Ctrl+a
    
'   Confirm 10 teams and 4 players per team or quit macro
   
    Dim TenAndFour As String
    TenAndFour = MsgBox("This Macro Only Works for a League Manager Report with 10 Teams and 4 Players per team and a season of 18 weeks.  Continue?", vbYesNo, "Average Points Per Round")
    Select Case TenAndFour
        Case Is = vbNo
        Exit Sub
    End Select
    
        
'   Calculate Number of Matches Played To Date by Each Team
    'starts the individual file and updates save type
    
    Windows("AGL-Individual-2019.xls").Activate
    Application.Calculation = xlAutomatic
    
    Columns("A").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    

    'Copy rankings
    'TM 6/16/20
    ActiveSheet.Range(Cells(3, 2), Cells(53, 2)).Select
    Selection.Copy
    Cells(3, 1).Select
    ActiveSheet.Paste
    Cells(3, 2).Value = "Avg. Pts. per Round"
    
    '6/16/20 TM
    'Average points per round for players
    For b = 4 To 49 Step 5
        For c = 0 To 3
            If Cells(b + c, 12).Value <> 0 Then
                Cells(b + c, 2).Value = (Cells(b + c, 7).Value / Cells(b + c, 12))
            Else
                Cells(b + c, 2).Value = 0
            End If
        Next
    Next
    
    '6/29/20 Rank Players
    For i = 4 To 49 Step 5
        For j = 0 To 3
            Cells(i + j, "A").Select
             ActiveCell.FormulaR1C1 = _
        "=RANK(RC[1],(R4C2:R7C2,R9C2:R12C2,R14C2:R17C2,R19C2:R22C2,R24C2:R27C2,R29C2:R32C2,R34C2:R37C2,R39C2:R42C2,R44C2:R47C2,R49C2:R52C2))"
        Next
    Next
    
    '6/17/20 TM
    'Calculates the highest amount of rounds played to account for subs
    'Assumes that any player didnt play additional rounds on their own
    'Calculates the max value in the rounds column to account for any subs
    roundsPlayed = 0
    For i = 4 To 49 Step 5
        For j = 0 To 3
            roundsPlayed = WorksheetFunction.Max(Cells(i + j, 12), roundsPlayed)
        Next
    Next
    
    '6/17/20 TM
    'Average points per round for teams
    'If all four players for a team have a score of 0 then team did not play that week
    'Loops through week 1 to current week and checks if each team played that week
    'Finds the number of weeks that each team played and divides by its total points
    roundsPerTeam = 0
    hasPlayed = True
    For i = 8 To 53 Step 5
    
        'Tests to see if team played during the current week
        For j = 1 To 4
            'if player did not play this week
            If Cells(i - j, 42).Value = 0 Or Cells(i - j, 42).Value = "" Then
                hasPlayed = False
            Else
                hasPlayed = True
                Exit For
            End If
        Next
        If hasPlayed = True Then
            roundsPerTeam = roundsPerTeam + 1
        End If
        
        'Test for weeks 1 to current week
        'If current week is week 1 then the loop doesnt nothing
        
        Dim start As Integer
        start = roundsPlayed - 1
        Do While start > 0
            For j = 1 To 4
            
            'If week has not been reached or players did not play
            'If any player on the team played then it skips to the next week
            'All four players need to have a score of 0 or "" to qualify as the team not having played
                If Cells(i - j, 31 - start).Value = 0 _
                    Or Cells(i - j, 31 - start).Value = "" Then
                    
                    hasPlayed = False
                Else
                    hasPlayed = True
                    Exit For
                End If
                
            Next
            start = start - 1
            If hasPlayed = True Then
                roundsPerTeam = roundsPerTeam + 1
            End If
        
        Loop
        'Worksheets("Standings").Cells(counter + 1, "E").Value
        If roundsPerTeam > 9 Then
            roundsPerTeam = roundsPerTeam - 9
        End If
        Cells(i, 2).Value = (Cells(i, "G").Value / roundsPerTeam)
        roundsPerTeam = 0
        counter = counter + 1
    Next
                
    
'   Color Rankings in Green if Sufficient Matches Played to Qualify for MVP
'   Left Original
  
    Windows("AGL-Macro-Holder-2019.xlsm").Activate
    RoundsNeeded = Range("Minimum_MVP_Matches").Value
    Windows("AGL-Individual-2019.xls").Activate
    Green = 0
    For i = 0 To 9
        For j = 0 To 3
            PlayerRounds = Cells(i * 5 + j + 4, 12).Value
            If PlayerRounds >= RoundsNeeded Then
                Cells(i * 5 + j + 4, 1).Select
                With Selection.Font
                    .Color = -11489280
                    .TintAndShade = 0
                End With
                Green = 1
            End If
        Next
    Next
    If Green = 1 Then
        Cells(2, 1).Value = "Green - played sufficient matches to qualify for MVP award"
        Cells(2, 1).Select
        With Selection
            .Font.Color = -11489280
            .Font.TintAndShade = 0
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            End With
    End If

    '6/17/20 TM
    'Insert Team Rankings
    For i = 8 To 53 Step 5
        Cells(i, 1).Select
        ActiveCell.FormulaR1C1 = _
        "=RANK(RC[1],(R8C2,R13C2,R18C2,R23C2,R28C2,R33C2,R38C2,R43C2,R48C2,R53C2))"
    Next
    
'   Calculate Number of Subs

    Dim rowfirstsub, rowlastsub, numbersubs, RowYTDStats As Integer
    
    Windows("AGL-Individual-2019.xls").Activate
    Anysubs = True
    AfterYTDStats = 0
    rowweekbogeys = 0
    RowYTDBogeys = 0
    rowfirstsub = 54
    Range("B54:B113").Select
    index = 0
    For Each cell In Selection
        If Left(Cells(rowfirstsub + index, "B").Value, 5) = "Event" Then
            If index >= 2 Then
                numbersubs = index - 1
                rowlastsub = rowfirstsub + numbersubs - 1
            Else
                Anysubs = False
                MsgBox ("There are no subs listed in the report.")
            End If
        End If
        If Left(Cells(rowfirstsub + index, "B").Value, 6) = "Season" Then
            RowYTDStats = index + rowfirstsub
            AfterYTDStats = 1
        End If
        If Left(Cells(rowfirstsub + index, "B").Value, 6) = "Most B" Then
            If AfterYTDStats = 0 Then
               rowweekbogeys = index + rowfirstsub
            Else
               RowYTDBogeys = index + rowfirstsub
            End If
        End If
        index = index + 1
    Next cell
    
'   If There are Subs, Clear Average Points per Round Field for Each

    If Anysubs Then
        For index = 0 To numbersubs
            Cells(rowfirstsub + index, "B").ClearContents
        Next
    End If

'   Clean Up Report For Printing

    
    Range("A3", "A" & rowlastsub).Select
    Application.CutCopyMode = False
    With Selection
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With

    Range("N" & rowlastsub + 2, "T" & rowlastsub + 31).Select
    Selection.Cut
    Range("L" & rowlastsub + 2).Select
    ActiveSheet.Paste
    Range("V" & rowlastsub + 2, "AD" & rowlastsub + 31).Select
    Selection.Cut
    Range("AE" & rowlastsub + 2).Select
    ActiveSheet.Paste
    
' Added by Dick Palmer 1/5/2014 to round HIGH POINTS.

    Range("K" & rowlastsub + 3).Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],13,3)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L" & rowlastsub + 3).Value = ("High Points " & Range("K" & rowlastsub + 3) & ")")
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("K" & RowYTDStats + 1).Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],13,3)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L" & RowYTDStats + 1).Value = ("High Points " & Range("K" & RowYTDStats + 1) & ")")
    Application.CutCopyMode = False
    Selection.ClearContents

    Columns("D:D").Select
    Selection.ColumnWidth = 2
    Columns("I:I").Select
    Selection.ColumnWidth = 2.5
    Columns("N:K").Select
    Selection.ColumnWidth = 3
    Columns("AE:AE").Select
    Selection.ColumnWidth = 7.5
    Columns("AO:AO").Select
    Selection.ColumnWidth = 2.5
    Columns("AU:AY").Select
    Selection.ColumnWidth = 2.67
    Rows("1:1").Select
    Selection.RowHeight = 9.75
    Rows("3:3").Select
    Selection.RowHeight = 27
    
    weekNumber = 0
    Windows("AGL-Macro-Holder-2019.xlsm").Activate
    Range("H1:H20").ClearContents
    For i = 2 To 20
        If Cells(i, "I").Value = roundsPlayed Then
            weekNumber = Cells(i, "J")
            Cells(i, "H").Value = "Last Report " & Chr(187) & Chr(187)
            Exit For
        End If
    Next
    Windows("AGL-Individual-2019.xls").Activate
    Cells(2, "J").Value = "As of Round " & roundsPlayed & " on " & weekNumber

    
'   Format Title Line

    Range("J2").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

'   If Week 9, put First Half Team Totals in Macro Holder Spreadsheet

    If roundsPlayed = 9 Then
        For i = 0 To 9
            Windows("AGL-Individual-2019.xls").Activate
            FirstHalfTotal = Range("G" & i * 5 + 8).Value
            Windows("AGL-Macro-Holder-2019.xlsm").Activate
            Range("B" & i + 2).Select
            ActiveCell.FormulaR1C1 = FirstHalfTotal
        Next
'       Range("C2:C11").Select
'       Selection.ClearContents
    End If

'   Update Macro File Team Points with Current Week Totals

    For i = 0 To 9
        Windows("AGL-Individual-2019.xls").Activate
        SecondHalfTotal = Cells(i * 5 + 8, "G").Value
        Windows("AGL-Macro-Holder-2019.xlsm").Activate
        Cells(i + 2, "C").Value = SecondHalfTotal
    Next

' Check for space between rows to clear LM Low Net contents.

 If rowweekbogeys - rowlastsub > 7 Then
     Range("G" & rowlastsub + 4, "G" & rowlastsub + 7).Select
     Selection.ClearContents
 Else
     Range("G" & rowlastsub + 4, "G" & rowlastsub + 5).Select
     Selection.ClearContents
 End If

'   Enter Individual Low Net Name & save for YTD comparison
    
    'Calculates the min par for this round
    'Loops through all golfers and compares the min score
    Windows("AGL-Individual-2019.xls").Activate
    minPar = 15
    For i = 4 To 53 Step 5
        For j = 0 To 3
            minPar = WorksheetFunction.Min(minPar, Cells(i + j, 11))
        Next
    Next
    For i = 54 To rowlastsub
        If Cells(i, 11).Value <> "" Then
            minPar = WorksheetFunction.Min(minPar, Cells(i, 11))
        End If
            
    Next
    
    'Max of 10 people tied for min net+-
    Dim lowNetNum(10) As Integer
    'Loops through all players and adds anyone who has the same net par as the min value to the array
    counter = 0
    For i = 4 To 53 Step 5
        For j = 0 To 3
            If Cells(i + j, 11).Value = minPar Then
                lowNetNum(counter) = Cells(i + j, 3)
                counter = counter + 1
            End If
        Next
    Next
    For i = 54 To rowlastsub
        If Cells(i, 11).Value = minPar Then
            lowNetNum(counter) = Cells(i, 3)
            counter = counter + 1
        End If
    Next
    
    
    ' Insert additional rows if number of winners is more than num of rows
    
    If counter > 2 And rowweekbogeys - rowlastsub < 7 Then
            Range("A" & rowweekbogeys).Select
                For i = 1 To 7 - (rowweekbogeys - rowlastsub)
                    ActiveCell.EntireRow.Insert Shift:=xlDown
                    RowYTDStats = RowYTDStats + 1
                    RowYTDBogeys = RowYTDBogeys + 1
                Next
    End If
    
    'TM 6/19/2020 Removed redudnacy from weekly net par winners
    Cells(rowlastsub + 3, 7).Value = "Low Net  " & "(" & minPar & ")"
    size = counter - 1
    counter = 0
    For i = 0 To size
        Windows("AGL-Individual-2019.xls").Activate
        Cells(rowlastsub + 4 + i, 7).Value = FindPlayer(lowNetNum(counter))
        Windows("AGL-Macro-Holder-2019.xlsm").Activate
        Cells(roundsPlayed + 109, 11 + (2 * i)).Value = lowNetNum(counter)
        Cells(roundsPlayed + 109, 12 + (2 * i)).Value = minPar
        counter = counter + 1
    Next
    
    
    
     
' Insert Weekly Low Net Score

    Windows("AGL-Individual-2019.xls").Activate
    Range("G" & rowlastsub + 3).Select
    With Selection
       .HorizontalAlignment = xlLeft
       .VerticalAlignment = xlCenter
       .WrapText = False
       .Orientation = 0
       .AddIndent = False
       .IndentLevel = 0
       .ReadingOrder = xlContext
       .MergeCells = False
    End With
    
'   Copy and sort Year-To-Date Individual low net player(s)and score(s)
 
    Windows("AGL-Macro-Holder-2019.xlsm").Activate
    Range("K110:R129").Select
    With Selection.Font
        .Bold = False
        .Name = "Arial"
        .size = 12
    End With


    'Finds yearly minimum par and the list of players who had that par
    IndivMin = YTDMin(110, 11)
    YTDNetIndiv = YTDCalculation(110, 11, IndivMin)
    
'   Enter Year-To-Date Individual low net player(s) and score(s) in report

    Windows("AGL-Individual-2019.xls").Activate
    Cells(RowYTDStats + 1, 7).Value = "Low Net  " & "(" & IndivMin & ")"

 'Modified by TM 6/19/2020 Removed redudancy in the conditions
 'TM 6/27/2020 Uses Array of players
 
        For i = 0 To counter - 1
            Windows("AGL-Individual-2019.xls").Activate
            Cells(RowYTDStats + 2 + i, 7).Value = FindPlayer(YTDNetIndiv(i) - 1)
        Next
    
' Format YTD Individual Winners
' Formatting

        Windows("AGL-Individual-2019.xls").Activate
        If YTDTotalIndWinners > 3 Then
            Range("G" & RowYTDStats + 2, "G" & RowYTDStats + YTDTotalIndWinners + 2).Select
        Else:
            Range("G" & RowYTDStats + 2, "G" & RowYTDStats + 4).Select
        End If
          With Selection
            .Font.Bold = False
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
    
        
' DETERMINE TEAM WEEKLY & YTD WINNERS
    
    'Calculates each Team Net Par +/-
    'Creates an array with 10 elements, each of represents 1 team
    'Array starts at 0 so Team 1(Bruins) would have an index of 0 etc
    'First loop adds each players net +- par to their teams total
    'Second loop accounts for any subs
    Dim teamNetPar(9) As Integer
    For i = 0 To 9
        For j = 0 To 4
            If Cells(j + (4 + (5 * i)), 5).Value <> "" Then
                teamNetPar(i) = teamNetPar(i) + Cells(j + 4 + (5 * i), 11)
            End If
        Next
    Next
    For i = 54 To rowlastsub
        If Cells(i, 4).Value <> "" Then
            teamNetPar(Cells(i, 4).Value - 1) = teamNetPar(Cells(i, 4).Value - 1) + Cells(i, 11)
        End If
    Next

    Windows("AGL-Individual-2019.xls").Activate
    
    'Finds Smallest value for team net par
    minTeamVal = 10
    counter = 0
    For i = 0 To 9
        minTeamVal = WorksheetFunction.Min(minTeamVal, teamNetPar(i))
    Next
    Cells(rowlastsub + 3, 31).Value = "Team Low Net  " & "(" & minTeamVal & ")"
    
    'Loops through all teams and prints those that have the min net par
    counter = 0
    For i = 0 To 9
        Windows("AGL-Individual-2019.xls").Activate
        If teamNetPar(i) = minTeamVal Then
            If counter > 2 Then
                Cells(rowlastsub + 5 + counter, 31).Select
                ActiveCell.EntireRow.Insert Shift:=xlDown
                RowYTDStats = RowYTDStats + 1
                RowYTDBogeys = RowYTDBogeys + 1
            End If
            Cells(rowlastsub + 4 + counter, 31).Value = TeamNicknames(i + 1)
            Windows("AGL-Macro-Holder-2019.xlsm").Activate
            Cells(roundsPlayed + 1, 11 + 2 * counter).Value = i + 1
            Cells(roundsPlayed + 1, 12 + 2 * counter).Value = minTeamVal
            counter = counter + 1
        End If
    Next
    
    Columns("L:L").ColumnWidth = 3.5
    Columns("AG:AG").ColumnWidth = 2
    Columns("AI:AI").ColumnWidth = 3
    Columns("AK:AK").ColumnWidth = 3
    Columns("AM:AM").ColumnWidth = 2
    Columns("AO:AO").ColumnWidth = 3
    
    'Calculating YTD Team Net par etc
    TeamMin = YTDMin(2, 11)
    YTDNetTeam = YTDCalculation(2, 11, TeamMin)
    Cells(RowYTDStats + 1, 31) = "Team Low Net  " & "(" & TeamMin & ")"
    For i = 0 To 15
        If YTDNetTeam(i) = 0 Then
            Exit For
        End If
        If i > 2 Then
            Cells(rowlastsub + 5 + counter, 31).Select
            ActiveCell.EntireRow.Insert Shift:=xlDown
            RowYTDBogeys = RowYTDBogeys + 1
        End If
        Cells(RowYTDStats + 2 + i, 31).Value = TeamNicknames(YTDNetTeam(i) - 1)
    Next
        

'   Copy and sort year-to-date low net team(s)and score(s)

    Windows("AGL-Macro-Holder-2019.xlsm").Activate
    Range("K2:R21").Select
    Selection.Font.Bold = False
    With Selection.Font
        .Name = "Arial"
        .size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    
'   Print Report?

    Windows("AGL-Individual-2019.xls").Activate
    
    ' Determine last line of report

    Lastline = RowYTDBogeys
    NeedTwoLines = 0
    Do Until NeedTwoLines = 2
        If Range("B" & Lastline) = "" And Range("G" & Lastline) = "" And Range("Y" & Lastline) = "" And Range("AF" & Lastline) = "" And Range("AH" & Lastline) = "" And Range("AL" & Lastline) = "" Then
            NeedTwoLines = NeedTwoLines + 1
        End If
        Lastline = Lastline + 1
    Loop
     
 '  Reformat Statistics Range
 'All reformats are placed at the end
    Range("A:A").Select
    Selection.Font.Bold = True
    Columns("AD:AD").ColumnWidth = 7
    
    Columns("a:a").EntireColumn.AutoFit
    Columns("b:b").ColumnWidth = 6

    Range("C:C").Select
    Selection.EntireColumn.Hidden = False
    Range("Z:D").Select
    Selection.EntireColumn.Hidden = False
    
    Range("B" & rowlastsub + 2 & ":AM" & Lastline).Select
    Selection.Cut
    Range("A" & rowlastsub + 2).Select
    ActiveSheet.Paste
      
    Range("K" & rowlastsub + 2 & ":Q" & Lastline).Select
    Selection.Cut
    Range("H" & rowlastsub + 2).Select
    ActiveSheet.Paste
    
    Range("AD" & rowlastsub + 2 & ":AL" & Lastline).Select
    Selection.Cut
    Range("AE" & rowlastsub + 2).Select
    ActiveSheet.Paste

    Columns("J:J").ColumnWidth = 3
    Columns("K:K").ColumnWidth = 3
    Columns("L:L").ColumnWidth = 3.5
    Columns("M:M").ColumnWidth = 3
    For i = 25 To 30
        Columns(i).ColumnWidth = 2.33
    Next
    Range("N:AD").Select
    Selection.EntireColumn.Hidden = True
    
    Columns("B:B").Select
    Selection.NumberFormat = "0.00"
    Selection.ColumnWidth = 8.83
    Columns("A:A").ColumnWidth = 5
    For i = 1 To 39
        Cells(rowlastsub + 2, i).Interior.ColorIndex = 35
        Cells(RowYTDStats, i).Interior.ColorIndex = 35
    Next
 
'   Print setup
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.2)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintInPlace
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
    End With
    ActiveSheet.PageSetup.PrintArea = Range("A1", "AY" & Lastline - 1).Address

    Dim Report As String
    Report = MsgBox("Print a Copy of the Report?", vbYesNo, "Average Points Per Round")
    Select Case Report
        Case Is = vbYes

'       Used for Bob's computer running Excel 2010 and attached printer

            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False
        
    End Select

'   Clean up Macro-Holder file and print End Statement

    MsgBox ("Be sure to save a copy of this week's Macro-Holder spreadsheet for use next week")

End Sub

Private Function TeamNicknames(ByVal num As Integer) As String
    If num <> 0 Then
        Windows("AGL-Macro-Holder-2019.xlsm").Activate
        TeamNicknames = ("#" & Cells(num + 1, 1))
    Else
        TeamNicknames = ""
    End If
    Windows("AGL-Individual-2019.xls").Activate
End Function
Private Function YTDMin(ByVal startRow As Integer, ByVal startCol As Integer) As Integer
    Windows("AGL-Macro-Holder-2019.xlsm").Activate
    minVal = 25
    For i = startRow To startRow + 17
        If Cells(i, startCol + 1).Value <> "" Then
            minVal = WorksheetFunction.Min(minVal, Cells(i, startCol + 1).Value)
        End If
    Next
    Windows("AGL-Individual-2019.xls").Activate
    YTDMin = minVal
End Function
Private Function YTDCalculation(ByVal startRow As Integer, ByVal startCol As Integer, ByVal minVal As Integer) As Integer()
    Windows("AGL-Macro-Holder-2019.xlsm").Activate
    Dim YTDNet(15) As Integer
    counter = 0
    Repeat = False
    For i = startRow To startRow + 17
        If Cells(i, startCol + 1).Value = minVal Then
            For j = 0 To 3
                If Cells(i, startCol + 2 * j).Value <> "" Then
                    'Loops for duplicates
                    For k = 0 To counter
                        If YTDNet(k) = Cells(i, startCol + 2 * j).Value + 1 Then
                            Repeat = True
                            Exit For
                        End If
                    Next
                    If Repeat = False Then
                        YTDNet(counter) = Cells(i, startCol + 2 * j).Value + 1
                        counter = counter + 1
                    Else
                        Repeat = False
                    End If
                End If
            Next
        End If
    Next
    Windows("AGL-Individual-2019.xls").Activate
    YTDCalculation = YTDNet
End Function
Private Function FindPlayer(ByVal playerNum As Integer) As String
'FindPlayer takes in the number of a specific player and returns the name of that player
'For teams 1-9 the rightmost digit in the player num is the team number
'In team 10 that does not apply so the IF statement accounts for that discrepancy
'The function uses that rightmost digit to calculate the start location for the later loop
'The loop only looks through a single team to find the player
'Using this trick speeds up the code by 90% as it no longer needs to search through 40 or so entries

    If Right(playerNum, 1) = "0" Then
        baseVal = 49
    Else
        baseVal = (CInt(Right(playerNum, 1)) - 1) * 5 + 4
    End If
    For i = 0 To 3
        If Cells(baseVal + i, 3).Value = playerNum Then
            player = Cells(baseVal + i, 6)
            Exit For
        End If
    Next
    FindPlayer = player
End Function
