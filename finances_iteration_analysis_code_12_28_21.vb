Option Explicit

Sub InsertRowAbove()

Rows(ActiveCell.Row).Insert

End Sub


Sub Iterate_Over_30day_Periods()
' So thought process:
' 1. Starting at: Evaluation cell: E3, Fill-in Cell: H3
' 2. Select 30-day evaluation range --> THIS INVOLVES SOME NOT-SO-STRAIGHTFORWARD LOGIC WITH DATES!
' 3. Based on category, perform categorical sums and fill in cells accordingly
' 4. Select next 30-day evaluation range
' 5. Repeat until end of range is an empty row

Dim Analysis_Sheet As Worksheet
Set Analysis_Sheet = Worksheets(3)
Analysis_Sheet.Activate

Dim Evaluation_Range As Range

' Start at cell B3
Dim Evaluation_Row_Start As Integer
Dim Evaluation_Row_End As Integer
Dim Evaluation_Col As Integer
Dim Evaluation_Amount_Col As Integer
Dim Date_Col As Integer
Evaluation_Row_Start = 3   ' Start at row 3
Evaluation_Col = 5   ' i.e., Column E --> Category
Evaluation_Amount_Col = 4   ' i.e., Column D --> Amount
Date_Col = 2    ' i.e., Column B --> Date
Set Evaluation_Range = Analysis_Sheet.Range(Cells(Evaluation_Row_Start, Evaluation_Col), Cells(Evaluation_Row_Start + 30, Evaluation_Col))
Evaluation_Range.Select

' Start at cell H3
Dim Fillin_Row As Integer
Dim Fillin_Col As Integer
Fillin_Row = 3   ' Start at row 3
Fillin_Col = 8   ' i.e., Column H

' Working with dates
Dim date_start As Date
Dim date_find As Date
Dim date_end As Date
date_start = CDate(Analysis_Sheet.Cells(Evaluation_Row_Start, Date_Col))
Dim continue_iteration_flag As Integer
continue_iteration_flag = 1
Dim do_once_flag As Integer
do_once_flag = 0


' Running totals and other variables
Dim iteration_num As Integer
iteration_num = 1   ' Start at 1

Dim bills_subs_running_total As Double
bills_subs_running_total = 0
Dim project_running_total As Double
project_running_total = 0
Dim family_running_total As Double
family_running_total = 0
Dim health_running_total As Double
health_running_total = 0
Dim groceries_running_total As Double
groceries_running_total = 0
Dim food_running_total As Double
food_running_total = 0
Dim gift_running_total As Double
gift_running_total = 0
Dim misc_running_total As Double
misc_running_total = 0
Dim car_running_total As Double
car_running_total = 0
Dim travel_running_total As Double
travel_running_total = 0
Dim overall_running_total As Double
overall_running_total = 0


' Operation to find end of range --> CAREFUL OF INFINITE LOOP!
date_end = DateAdd("d", 30, date_start)
Evaluation_Row_End = Evaluation_Row_Start
While date_find <= date_end And continue_iteration_flag = 1
    Evaluation_Row_End = Evaluation_Row_End + 1
    ' To ensure we're not checking past the last entry
    If IsEmpty(Analysis_Sheet.Cells(Evaluation_Row_End, Date_Col)) Then
        do_once_flag = 1    ' Just iterate until last entry
        continue_iteration_flag = 0  ' Stop iterating
    End If
    date_find = CDate(Analysis_Sheet.Cells(Evaluation_Row_End, Date_Col))
Wend

' Set next Evaluation_Range accordingly
Set Evaluation_Range = Analysis_Sheet.Range(Cells(Evaluation_Row_Start, Evaluation_Col), Cells(Evaluation_Row_End, Evaluation_Col))
Evaluation_Range.Select

'While Not IsEmpty(Analysis_Sheet.Cells(Evaluation_Row_Start + 30, Evaluation_Col).Value)        ' if end cell of range is not empty...
While continue_iteration_flag = 1 Or do_once_flag = 1

    Analysis_Sheet.Cells(Fillin_Row, 24).Value = Evaluation_Row_Start
    Analysis_Sheet.Cells(Fillin_Row, 25).Value = Evaluation_Row_End
    Analysis_Sheet.Cells(Fillin_Row, 26) = date_start
    Analysis_Sheet.Cells(Fillin_Row, 27) = date_find
    
    ' Reset all running totals
    bills_subs_running_total = 0
    project_running_total = 0
    family_running_total = 0
    health_running_total = 0
    groceries_running_total = 0
    food_running_total = 0
    gift_running_total = 0
    misc_running_total = 0
    car_running_total = 0
    travel_running_total = 0
    overall_running_total = 0
    
    Dim eval_row As Integer
    Dim eval_category As String
    ' Run through range of 30 cells and execute sums
    For eval_row = Evaluation_Row_Start To Evaluation_Row_End
        
        eval_category = Analysis_Sheet.Cells(eval_row, Evaluation_Col).Value
        Select Case eval_category    ' Based on category
            Case "Bills/Subs"
                bills_subs_running_total = bills_subs_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Project"
                project_running_total = project_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Family"
                family_running_total = family_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Health"
                health_running_total = health_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Groceries"
                groceries_running_total = groceries_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Food"
                food_running_total = food_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Gift"
                gift_running_total = gift_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Misc"
                misc_running_total = misc_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Car"
                car_running_total = car_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
            Case "Travel"
                travel_running_total = travel_running_total + Analysis_Sheet.Cells(eval_row, Evaluation_Amount_Col)
                
        End Select
        
    Next eval_row
    
    overall_running_total = bills_subs_running_total + project_running_total + family_running_total + _
                            health_running_total + groceries_running_total + food_running_total + gift_running_total + _
                            misc_running_total + car_running_total + travel_running_total
    
    ' Fill up results into Fill-in range
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 0).Value = iteration_num
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 1).Value = bills_subs_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 2).Value = project_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 3).Value = family_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 4).Value = health_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 5).Value = groceries_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 6).Value = food_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 7).Value = gift_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 8).Value = misc_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 9).Value = car_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 10).Value = travel_running_total
    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 11).Value = overall_running_total
    
    
    If do_once_flag = 1 Then
        do_once_flag = 0
    Else
        ' Set up next range
        iteration_num = iteration_num + 1
        ' Evaluation_Row_Start = Evaluation_Row_Start + 1
        Fillin_Row = Fillin_Row + 1
        
        ' Operation to establish next start of range
        date_find = date_start
        date_start = DateAdd("d", 1, date_start)
        While date_find < date_start
            Evaluation_Row_Start = Evaluation_Row_Start + 1
            date_find = CDate(Analysis_Sheet.Cells(Evaluation_Row_Start, Date_Col))
        Wend
        
        ' Operation to find end of range --> CAREFUL OF INFINITE LOOP!
        date_end = DateAdd("d", 30, date_start)
        Evaluation_Row_End = Evaluation_Row_Start
        While date_find <= date_end And continue_iteration_flag = 1
            Evaluation_Row_End = Evaluation_Row_End + 1
            ' To ensure we're not checking past the last entry
            If IsEmpty(Analysis_Sheet.Cells(Evaluation_Row_End, Date_Col)) Then
                do_once_flag = 1    ' Just iterate until last entry
                continue_iteration_flag = 0  ' Stop iterating
            End If
            date_find = CDate(Analysis_Sheet.Cells(Evaluation_Row_End, Date_Col))
        Wend
        
        ' Set next Evaluation_Range accordingly
        Set Evaluation_Range = Analysis_Sheet.Range(Cells(Evaluation_Row_Start, Evaluation_Col), Cells(Evaluation_Row_End, Evaluation_Col))
        Evaluation_Range.Select
    
    End If
    
Wend


End Sub

'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 0).Value = iteration_num
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 1).Value = bills_subs_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 2).Value = project_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 3).Value = family_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 4).Value = health_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 5).Value = groceries_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 6).Value = food_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 7).Value = gift_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 8).Value = misc_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 9).Value = car_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 10).Value = travel_running_total
'Analysis_Sheet.Cells(Fillin_Row, Fillin_Col + 11).Value = overall_running_total

'If Not IsEmpty(Analysis_Sheet.Cells(Evaluation_Row_Start + 30, Evaluation_Col)) Then
'    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col).Value = 25
'End If

'If IsEmpty(Analysis_Sheet.Cells(1, 1)) = False Then
'    Analysis_Sheet.Cells(1, 2).Value = 25
'End If

'If IsEmpty(Analysis_Sheet.Cells(18, 10)) Then
'    Analysis_Sheet.Cells(Fillin_Row, Fillin_Col).Value = 25
'End If

'date1 = CDate(Analysis_Sheet.Cells(24, 2))   ' This would be 11/10/21
'date2 = DateAdd("d", 30, date1)
'Analysis_Sheet.Cells(1, 2) = date2
