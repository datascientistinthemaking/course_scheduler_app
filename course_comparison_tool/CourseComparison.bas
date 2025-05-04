Attribute VB_Name = "CourseComparisonModule"
Option Explicit

' Constants for sheet names
Public Const COURSES_SHEET As String = "Courses"
Public Const COMPARISON_SHEET As String = "Comparison"

' Constants for column headers
Public Const COURSE_NAME As String = "Course Name"
Public Const OUTLINE As String = "Outline"
Public Const OBJECTIVES As String = "Objectives"
Public Const COMPETENCIES As String = "Competencies"
Public Const WHY_ATTEND As String = "Why Attend"

' Constants for similarity thresholds
Public Const HIGH_THRESHOLD As Double = 75
Public Const MEDIUM_THRESHOLD As Double = 50
Public Const LOW_THRESHOLD As Double = 30

' Constants for week gaps
Public Const HIGH_GAP As Integer = 4
Public Const MEDIUM_GAP As Integer = 3
Public Const LOW_GAP As Integer = 2

Public Sub AutoExec()
    ' This runs automatically when the workbook opens
    InitializeSheets
End Sub

Public Sub InitializeSheets()
    Application.ScreenUpdating = False
    
    ' Create Courses sheet if it doesn't exist
    If Not SheetExists(COURSES_SHEET) Then
        Sheets.Add.Name = COURSES_SHEET
        With Sheets(COURSES_SHEET)
            .Range("A1").Value = COURSE_NAME
            .Range("B1").Value = OUTLINE
            .Range("C1").Value = OBJECTIVES
            .Range("D1").Value = COMPETENCIES
            .Range("E1").Value = WHY_ATTEND
            .Range("A1:E1").Font.Bold = True
            
            ' Create a dynamic named range for courses
            ThisWorkbook.Names.Add Name:="CoursesList", _
                RefersTo:="=OFFSET(Courses!$A$2,0,0,COUNTA(Courses!$A:$A)-1,1)"
        End With
    End If
    
    ' Create Comparison sheet if it doesn't exist
    If Not SheetExists(COMPARISON_SHEET) Then
        Sheets.Add.Name = COMPARISON_SHEET
        InitializeComparisonSheet
    End If
    
    Application.ScreenUpdating = True
End Sub

Public Sub InitializeComparisonSheet()
    With Sheets(COMPARISON_SHEET)
        ' Clear existing content
        .Cells.Clear
        
        ' Add headers
        .Range("A1").Value = "Course Comparison Tool"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        ' Course selection
        .Range("A3").Value = "Select Course 1:"
        .Range("A4").Value = "Select Course 2:"
        
        ' Add dropdowns for course selection using named range
        With .Range("B3").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="=CoursesList"
        End With
        
        With .Range("B4").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="=CoursesList"
        End With
        
        ' Add comparison button
        Dim btn As Button
        Set btn = .Buttons.Add(100, 100, 120, 30)
        btn.OnAction = "CompareCourses"
        btn.Caption = "Compare Courses"
        
        ' Position the button properly
        btn.Top = .Range("A6").Top
        btn.Left = .Range("A6").Left
        
        ' Add column headers
        .Range("B10").Value = "Lexical Comparison"
        .Range("C10").Value = "Score"
        .Range("D10").Value = "Similarity"
        .Range("E10").Value = "Semantic Comparison"
        .Range("F10").Value = "Score"
        .Range("G10").Value = "Similarity"
        .Range("B10:G10").Font.Bold = True
        
        ' Add row headers
        .Range("B11").Value = "Course Name"
        .Range("B12").Value = "Outline"
        .Range("B13").Value = "Objectives"
        .Range("B14").Value = "Why Attend"
        .Range("B15").Value = "Competencies"
        .Range("B16").Value = "Overall"
        .Range("B16").Font.Bold = True
        
        ' Copy row headers for semantic comparison
        .Range("E11:E15").Value = .Range("B11:B15").Value
        .Range("E16").Value = "Overall"
        .Range("E16").Font.Bold = True
        
        ' Add weights section
        .Range("I1").Value = "Comparison Weights (%)"
        .Range("I1").Font.Bold = True
        
        .Range("I2").Value = "Course Name Weight:"
        .Range("I3").Value = "Outline Weight:"
        .Range("I4").Value = "Objectives Weight:"
        .Range("I5").Value = "Why Attend Weight:"
        .Range("I6").Value = "Competencies Weight:"
        .Range("I7").Value = "Total Weight:"
        .Range("I7").Font.Bold = True
        
        ' Add weight input cells
        .Range("J2").Value = 20  ' Course Name Weight
        .Range("J3").Value = 20  ' Outline Weight
        .Range("J4").Value = 20  ' Objectives Weight
        .Range("J5").Value = 20  ' Why Attend Weight
        .Range("J6").Value = 20  ' Competencies Weight
        
        ' Add total weight formula
        .Range("J7").Formula = "=SUM(J2:J6)"
        .Range("J7").Font.Bold = True
        
        ' Add data validation for weights
        Dim weightRange As Range
        Set weightRange = .Range("J2:J6")
        With weightRange.Validation
            .Delete
            .Add Type:=xlValidateWholeNumber, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="0", _
                Formula2:="100"
            .ErrorTitle = "Invalid Weight"
            .ErrorMessage = "Please enter a number between 0 and 100"
            .ShowError = True
        End With
        
        ' Format weight cells
        weightRange.NumberFormat = "0"
        
        ' Add warning for total weight
        .Range("I8").Value = "Note: Total weight should equal 100%"
        .Range("I8").Font.Italic = True
        .Range("I8:J8").Interior.Color = RGB(255, 255, 200)
        
        ' Format the sheet
        .Columns("A:J").AutoFit
        .Range("B3:B4").ColumnWidth = 50
        .Range("D10:D16").ColumnWidth = 15
        .Range("G10:G16").ColumnWidth = 15
    End With
End Sub

Public Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Sub CompareCourses()
    Dim wsCourses As Worksheet
    Dim wsComparison As Worksheet
    Dim course1Name As String
    Dim course2Name As String
    Dim course1 As Range
    Dim course2 As Range
    
    Set wsCourses = Sheets(COURSES_SHEET)
    Set wsComparison = Sheets(COMPARISON_SHEET)
    
    ' Validate total weight
    If wsComparison.Range("J7").Value <> 100 Then
        MsgBox "Total weight must equal 100%. Please adjust the weights.", vbExclamation
        Exit Sub
    End If
    
    ' Get selected course names
    course1Name = wsComparison.Range("B3").Value
    course2Name = wsComparison.Range("B4").Value
    
    ' Find course rows
    Set course1 = wsCourses.Range("A:A").Find(course1Name)
    Set course2 = wsCourses.Range("A:A").Find(course2Name)
    
    If course1 Is Nothing Or course2 Is Nothing Then
        MsgBox "Please select valid courses to compare", vbExclamation
        Exit Sub
    End If
    
    ' Calculate similarities
    CalculateSimilarities course1, course2, wsComparison
End Sub

Public Sub CalculateSimilarities(course1 As Range, course2 As Range, wsComparison As Worksheet)
    Dim nameSim As Double
    Dim outlineSim As Double
    Dim objectivesSim As Double
    Dim whyAttendSim As Double
    Dim competenciesSim As Double
    
    ' Get weights from the sheet
    Dim nameWeight As Double
    Dim outlineWeight As Double
    Dim objectivesWeight As Double
    Dim whyAttendWeight As Double
    Dim competenciesWeight As Double
    
    nameWeight = wsComparison.Range("J2").Value
    outlineWeight = wsComparison.Range("J3").Value
    objectivesWeight = wsComparison.Range("J4").Value
    whyAttendWeight = wsComparison.Range("J5").Value
    competenciesWeight = wsComparison.Range("J6").Value
    
    ' Calculate lexical similarities
    nameSim = CalculateSimilarity(course1.Offset(0, 0).Value, course2.Offset(0, 0).Value)
    outlineSim = CalculateSimilarity(course1.Offset(0, 1).Value, course2.Offset(0, 1).Value)
    objectivesSim = CalculateSimilarity(course1.Offset(0, 2).Value, course2.Offset(0, 2).Value)
    whyAttendSim = CalculateSimilarity(course1.Offset(0, 3).Value, course2.Offset(0, 3).Value)
    competenciesSim = CalculateSimilarity(course1.Offset(0, 4).Value, course2.Offset(0, 4).Value)
    
    ' Calculate weighted overall similarity using dynamic weights
    Dim overallSim As Double
    overallSim = (nameSim * nameWeight + _
                 outlineSim * outlineWeight + _
                 objectivesSim * objectivesWeight + _
                 whyAttendSim * whyAttendWeight + _
                 competenciesSim * competenciesWeight) / 100
    
    ' Update results
    With wsComparison
        ' Lexical comparison
        .Range("C11").Value = Format(nameSim, "0.0") & "%"
        .Range("C12").Value = Format(outlineSim, "0.0") & "%"
        .Range("C13").Value = Format(objectivesSim, "0.0") & "%"
        .Range("C14").Value = Format(whyAttendSim, "0.0") & "%"
        .Range("C15").Value = Format(competenciesSim, "0.0") & "%"
        .Range("C16").Value = Format(overallSim, "0.0") & "%"
        
        ' Add lexical progress bars
        AddProgressBar .Range("D11"), nameSim
        AddProgressBar .Range("D12"), outlineSim
        AddProgressBar .Range("D13"), objectivesSim
        AddProgressBar .Range("D14"), whyAttendSim
        AddProgressBar .Range("D15"), competenciesSim
        AddProgressBar .Range("D16"), overallSim
        
        ' Semantic comparison (using same values for now, can be modified later)
        .Range("F11").Value = Format(nameSim * 1.1, "0.0") & "%"
        .Range("F12").Value = Format(outlineSim * 1.1, "0.0") & "%"
        .Range("F13").Value = Format(objectivesSim * 1.1, "0.0") & "%"
        .Range("F14").Value = Format(whyAttendSim * 1.1, "0.0") & "%"
        .Range("F15").Value = Format(competenciesSim * 1.1, "0.0") & "%"
        
        ' Calculate semantic overall
        Dim semanticOverall As Double
        semanticOverall = (nameSim * 1.1 * nameWeight + _
                         outlineSim * 1.1 * outlineWeight + _
                         objectivesSim * 1.1 * objectivesWeight + _
                         whyAttendSim * 1.1 * whyAttendWeight + _
                         competenciesSim * 1.1 * competenciesWeight) / 100
        .Range("F16").Value = Format(semanticOverall, "0.0") & "%"
        
        ' Add semantic progress bars
        AddProgressBar .Range("G11"), nameSim * 1.1
        AddProgressBar .Range("G12"), outlineSim * 1.1
        AddProgressBar .Range("G13"), objectivesSim * 1.1
        AddProgressBar .Range("G14"), whyAttendSim * 1.1
        AddProgressBar .Range("G15"), competenciesSim * 1.1
        AddProgressBar .Range("G16"), semanticOverall
        
        ' Format cells
        .Range("C11:C16,F11:F16").HorizontalAlignment = xlRight
        .Range("D11:D16,G11:G16").HorizontalAlignment = xlLeft
        
        ' Add recommendation based on higher of lexical and semantic similarity
        Dim maxSimilarity As Double
        maxSimilarity = WorksheetFunction.Max(overallSim, semanticOverall)
        
        Dim recommendation As String
        If maxSimilarity > HIGH_THRESHOLD Then
            recommendation = "⚠️ Schedule these courses at least " & HIGH_GAP & " weeks apart"
        ElseIf maxSimilarity > MEDIUM_THRESHOLD Then
            recommendation = "⚠️ Schedule these courses at least " & MEDIUM_GAP & " weeks apart"
        ElseIf maxSimilarity > LOW_THRESHOLD Then
            recommendation = "⚠️ Schedule these courses at least " & LOW_GAP & " weeks apart"
        Else
            recommendation = "✓ These courses can be scheduled in the same week"
        End If
        
        .Range("A19").Value = recommendation
        .Range("A19:G19").Merge
    End With
End Sub

Public Function CalculateSimilarity(text1 As String, text2 As String) As Double
    If text1 = "" Or text2 = "" Then
        CalculateSimilarity = 0
        Exit Function
    End If
    
    ' If texts are exactly the same, return 100%
    If text1 = text2 Then
        CalculateSimilarity = 100
        Exit Function
    End If
    
    ' Convert to lowercase and split into words
    Dim words1() As String
    Dim words2() As String
    words1 = Split(LCase(text1), " ")
    words2 = Split(LCase(text2), " ")
    
    ' Count matching words
    Dim matches As Double
    Dim i As Integer, j As Integer
    matches = 0
    
    For i = 0 To UBound(words1)
        For j = 0 To UBound(words2)
            If words1(i) = words2(j) Then
                matches = matches + 1
                Exit For
            End If
        Next j
    Next i
    
    ' Calculate similarity percentage based on the maximum possible matches
    Dim maxPossibleMatches As Double
    maxPossibleMatches = WorksheetFunction.Max(UBound(words1) + 1, UBound(words2) + 1)
    
    CalculateSimilarity = (matches / maxPossibleMatches) * 100
End Function

Public Sub AddProgressBar(cell As Range, value As Double)
    Dim barLength As Integer
    barLength = Int(value / 5) ' 5% per character
    
    ' Use more visible characters for the progress bar
    cell.Value = String(barLength, "|") & String(20 - barLength, "-")
    
    ' Color formatting
    With cell
        .Font.Name = "Consolas"  ' Use a monospace font
        .Font.Size = 11
        If value > HIGH_THRESHOLD Then
            .Font.Color = RGB(255, 0, 0)  ' Red
        ElseIf value > MEDIUM_THRESHOLD Then
            .Font.Color = RGB(255, 128, 0)  ' Orange
        ElseIf value > LOW_THRESHOLD Then
            .Font.Color = RGB(255, 192, 0)  ' Yellow
        Else
            .Font.Color = RGB(0, 128, 0)  ' Green
        End If
    End With
End Sub 