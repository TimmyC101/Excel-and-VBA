# Excel-and-VBA

## Well Monitoring Tool

During my time at Chevron, I identified several processes that were inefficient and sought to rectify the situation.  To that end, I taught myself VBA and created custom Excel tools to reduce manual labor and streamline processes.  Within this repository is a monitoring tool that I designed to take in a few inputs from a well tracking software and create numerous graphs and charts for quick and simple data analysis.

Before this tool, Chevron's technologists would spend around 30 minutes per day updating an Excel spreadsheet that was difficult to parse and did not provide the data required for well tracking, analysis, and performance optimization.  With this tool, not only was the data analysis component trivialized, but the time required to update the spreadsheet was reduced to fewer than 5 minutes per week.

To view the spreadsheet in its entirety, please click on the Excel spreadsheet within the repository.

# VBA Code

Sub Macro()

'Declare Ranges and Series ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' AFE Variables
Dim AFE_Depth As Range
Dim AFE_Days As Range
Dim AFE_Cost As Range
Dim P10_AFE_Days As Range
Dim P10_AFE_Cost As Range
Dim P90_AFE_Days As Range
Dim P90_AFE_Cost As Range

' Actual Variables
Dim Actual_Depth As Range
Dim Actual_Days As Range
Dim Actual_Cost As Range

' Projection Variables
Dim Projected_Depth As Range
Dim P10_Days As Range
Dim Mean_Days As Range
Dim P90_Days As Range
Dim P10_Cost As Range
Dim Mean_Cost As Range
Dim P90_Cost As Range

' Phase Variables
Dim Phase_Names As Range
Dim Phase_AFE_Days As Range
Dim Phase_Actual_TF_Days As Range
Dim Phase_Actual_NPT_Days As Range
Dim Phase_Time_Plus_Minus As Range
Dim Phase_AFE_Cost As Range
Dim Phase_Actual_Cost As Range
Dim Phase_Cost_Plus_Minus As Range
Dim Phase_Spacer_1 As Range
Dim Phase_Spacer_2 As Range
Dim LastPhase_Days As Range
Dim LastPhase_Cost As Range

' AFE Cost Variables
Dim Category_Names As Range
Dim Category_AFE_Cost As Range
Dim Category_Actual_Cost As Range
Dim AFE_Sort As Range
Dim AFE_Cost_Sort As Range

' GO-36 Variables
Dim GO_DHC As Range
Dim GO_DHC_Depth As Range
Dim GO_Actual_Cost As Range
Dim GO_P10 As Range
Dim GO_Mean As Range
Dim GO_P90 As Range

'Phase Vs Days
Dim Phase_Vs_Days As Range
Dim Phase_Vs_Days_Adj As Range
Dim Phase_Names_NonWW As Range
Dim P10_Days_PvD As Range
Dim Mean_Days_PvD As Range
Dim P90_Days_PvD As Range

'Expected Dry Hole Days
Dim DHD As Double
Dim Num1 As Double
Dim Num2 As Double
Dim Spud As Double

'Scope Change
Dim SC1 As Integer
Dim SC2 As Integer
Dim SC3 As Integer
Dim SC4 As Integer

'''''''''''''''''''''''''''''''''''''''''''''' End Variable Declaration '''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Application.ScreenUpdating = False


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''' DvD & CvD & Days v Cost''''''''''''''''''''''''''''''''''''''''''


'FOR loops and IF statements to create ranges needed for DvD and CvD Graphs '''''''''''''''''''''''''''''''''''
'This is necessary to remove empty cells from the chart series (not doing this results in incorrect scatter plot)

Worksheets("Original Plan").Activate


' AFE
Set AFE_Depth = Range("AA36")
'Above line is to set first data point to water depth
For Each Cell In Range("AA2", "AA35")
    If Not Cell.Value = "" Then
        Set AFE_Depth = Union(AFE_Depth, Cell)
    End If
Next

Set AFE_Days = Range("Z36")
'Above line is to set first data point to 0
For Each Cell In Range("AX2", "AX35")
    If Not Cell.Value = "" Then
        Set AFE_Days = Union(AFE_Days, Cell)
    End If
Next

Set AFE_Cost = Range("Z36")
'Above line is to set first data point to 0
For Each Cell In Range("BA2", "BA35")
    If Not Cell.Value = "" Then
        Set AFE_Cost = Union(AFE_Cost, Cell)
    End If
Next

'P10
Set P10_AFE_Days = Range("Z36")
'Above line is to set first data point to 0
For Each Cell In Range("AW2", "AW35")
    If Not Cell.Value = "" Then
        Set P10_AFE_Days = Union(P10_AFE_Days, Cell)
    End If
Next

Set P10_AFE_Cost = Range("Z36")
'Above line is to set first data point to 0
For Each Cell In Range("AZ2", "AZ35")
    If Not Cell.Value = "" Then
        Set P10_AFE_Cost = Union(P10_AFE_Cost, Cell)
    End If
Next

'P90
Set P90_AFE_Days = Range("Z36")
'Above line is to set first data point to 0
For Each Cell In Range("AY2", "AY35")
    If Not Cell.Value = "" Then
        Set P90_AFE_Days = Union(P90_AFE_Days, Cell)
    End If
Next

Set P90_AFE_Cost = Range("Z36")
'Above line is to set first data point to 0
For Each Cell In Range("BB2", "BB35")
    If Not Cell.Value = "" Then
        Set P90_AFE_Cost = Union(P90_AFE_Cost, Cell)
    End If
Next


' Actuals
Worksheets("DvD Export").Activate

Set Actual_Depth = Range("G1")
'Above line is to set first data point to the well start depth
For Each Cell In Range("D2", "D1000")
    If Not Cell.Value = "" Then
        If Not Actual_Depth Is Nothing Then
            Set Actual_Depth = Union(Actual_Depth, Cell)
        Else
            Set Actual_Depth = Cell
        End If
    End If
Next

Set Actual_Days = Range("F1")
'Above line is to set first data point to 0
For Each Cell In Range("A2", "A1000")
    If Not Cell.Value = "" Then
        If Not Actual_Days Is Nothing Then
            Set Actual_Days = Union(Actual_Days, Cell)
        Else
            Set Actual_Days = Cell
        End If
    End If
Next

Set Actual_Cost = Range("F1")
'Above line is to set first data point to 0
For Each Cell In Range("C2", "C1000")
    If Not Cell.Value = "" Then
        If Not Actual_Cost Is Nothing Then
            Set Actual_Cost = Union(Actual_Cost, Cell)
        Else
            Set Actual_Cost = Cell
        End If
    End If
Next



''''''''''''''''''''''''''''''''''''''''''' Projections '''''''''''''''''''''''''''''''''''''''''''

Worksheets("Phase Export").Activate

For Each Cell In Range("AD2", "AD35")
    If Not Cell.Value = "" Then
        If Not Projected_Depth Is Nothing Then
            Set Projected_Depth = Union(Projected_Depth, Cell)
        Else
            Set Projected_Depth = Cell
        End If
    End If
Next

For Each Cell In Range("AE2", "AE35")
    If Not Cell.Value = "" Then
        If Not P10_Days Is Nothing Then
            Set P10_Days = Union(P10_Days, Cell)
        Else
            Set P10_Days = Cell
        End If
    End If
Next

For Each Cell In Range("AF2", "AF35")
    If Not Cell.Value = "" Then
        If Not Mean_Days Is Nothing Then
            Set Mean_Days = Union(Mean_Days, Cell)
        Else
            Set Mean_Days = Cell
        End If
    End If
Next

For Each Cell In Range("AG2", "AG35")
    If Not Cell.Value = "" Then
        If Not P90_Days Is Nothing Then
            Set P90_Days = Union(P90_Days, Cell)
        Else
            Set P90_Days = Cell
        End If
    End If
Next

For Each Cell In Range("AH2", "AH35")
    If Not Cell.Value = "" Then
        If Not P10_Cost Is Nothing Then
            Set P10_Cost = Union(P10_Cost, Cell)
        Else
            Set P10_Cost = Cell
        End If
    End If
Next

For Each Cell In Range("AI2", "AI35")
    If Not Cell.Value = "" Then
        If Not Mean_Cost Is Nothing Then
            Set Mean_Cost = Union(Mean_Cost, Cell)
        Else
            Set Mean_Cost = Cell
        End If
    End If
Next

For Each Cell In Range("AJ2", "AJ35")
    If Not Cell.Value = "" Then
        If Not P90_Cost Is Nothing Then
            Set P90_Cost = Union(P90_Cost, Cell)
        Else
            Set P90_Cost = Cell
        End If
    End If
Next

'Updates Series in Days v Depth Chart '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("Days v Depth").Activate
ActiveSheet.ChartObjects("DvD").Activate

ActiveChart.SeriesCollection("Actual").Values = Actual_Depth
ActiveChart.SeriesCollection("Actual").XValues = Actual_Days
    
'Add If statement so that these do not try to plot if the well is finished
If Not Projected_Depth Is Nothing Then
    ActiveChart.SeriesCollection("P10 Projection").Values = Projected_Depth
    ActiveChart.SeriesCollection("P10 Projection").XValues = P10_Days

    ActiveChart.SeriesCollection("Mean Projection").Values = Projected_Depth
    ActiveChart.SeriesCollection("Mean Projection").XValues = Mean_Days


    ActiveChart.SeriesCollection("P90 Projection").Values = Projected_Depth
    ActiveChart.SeriesCollection("P90 Projection").XValues = P90_Days
End If

ActiveChart.SeriesCollection("P10").Values = AFE_Depth
ActiveChart.SeriesCollection("P10").XValues = P10_AFE_Days

ActiveChart.SeriesCollection("AFE").Values = AFE_Depth
ActiveChart.SeriesCollection("AFE").XValues = AFE_Days

ActiveChart.SeriesCollection("P90").Values = AFE_Depth
ActiveChart.SeriesCollection("P90").XValues = P90_AFE_Days

'Updates Series in Cost v Depth Chart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("Cost v Depth").Activate
ActiveSheet.ChartObjects("CvD").Activate

ActiveChart.SeriesCollection("Actual").Values = Actual_Depth
ActiveChart.SeriesCollection("Actual").XValues = Actual_Cost
    
'Add If statement so that these do not try to plot if the well is finished
If Not Projected_Depth Is Nothing Then
    ActiveChart.SeriesCollection("P10 Projection").Values = Projected_Depth
    ActiveChart.SeriesCollection("P10 Projection").XValues = P10_Cost

    ActiveChart.SeriesCollection("Mean Projection").Values = Projected_Depth
    ActiveChart.SeriesCollection("Mean Projection").XValues = Mean_Cost

    ActiveChart.SeriesCollection("P90 Projection").Values = Projected_Depth
    ActiveChart.SeriesCollection("P90 Projection").XValues = P90_Cost
End If

ActiveChart.SeriesCollection("P10").Values = AFE_Depth
ActiveChart.SeriesCollection("P10").XValues = P10_AFE_Cost

ActiveChart.SeriesCollection("AFE").Values = AFE_Depth
ActiveChart.SeriesCollection("AFE").XValues = AFE_Cost

ActiveChart.SeriesCollection("P90").Values = AFE_Depth
ActiveChart.SeriesCollection("P90").XValues = P90_AFE_Cost

'Updates Series in Days v Cost Chart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("Days v Cost").Activate
ActiveSheet.ChartObjects("Days v Cost").Activate

ActiveChart.SeriesCollection("Actual").Values = Actual_Cost
ActiveChart.SeriesCollection("Actual").XValues = Actual_Days
    
'Add If statement so that these do not try to plot if the well is finished
If Not Projected_Depth Is Nothing Then
    ActiveChart.SeriesCollection("P10 Projection").Values = P10_Cost
    ActiveChart.SeriesCollection("P10 Projection").XValues = P10_Days

    ActiveChart.SeriesCollection("Mean Projection").Values = Mean_Cost
    ActiveChart.SeriesCollection("Mean Projection").XValues = Mean_Days

    ActiveChart.SeriesCollection("P90 Projection").Values = P90_Cost
    ActiveChart.SeriesCollection("P90 Projection").XValues = P90_Days
End If

ActiveChart.SeriesCollection("AFE").Values = AFE_Cost
ActiveChart.SeriesCollection("AFE").XValues = AFE_Days

''''''''''''''''''''''''''''''''''''''''End DvD & DvC & Days v Cost''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''' Phase Days & Phase Costs ''''''''''''''''''''''''''''''''''''''''


'FOR loops and IF statements to create ranges needed for Phase Days and Phase Costs Charts'''''''''''''''''

Worksheets("Phase Export").Activate

' Determine number of phases for use in Ranges
For Each Cell In Range("A2", "A35")
    If Not Cell.Value = "" Then
        If Not Cell.Value = "NONWW" Then 'Added to account for new MUE methodology
        'Create PhaseNum counter to accurately determine range for other Series
            PhaseNum = PhaseNum + 1
        End If
    End If
Next

For Each Cell In Range("AS2", "AS35")
    If Not Cell.Value = "" Then
        If Not Cell.Value = "NONWW" Then 'Added to account for new MUE methodology
            If Not Phase_Names Is Nothing Then
                Set Phase_Names = Union(Phase_Names, Cell)
            Else
                Set Phase_Names = Cell
            End If
        End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("AM2", "AM35")
    If Counter < PhaseNum Then
        If Not Phase_AFE_Days Is Nothing Then
            Set Phase_AFE_Days = Union(Phase_AFE_Days, Cell)
            Counter = Counter + 1
        Else
            Set Phase_AFE_Days = Cell
            Counter = 1
        End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("AK2", "AK35")
    If Counter < PhaseNum Then
        If Not Phase_Actual_TF_Days Is Nothing Then
            Set Phase_Actual_TF_Days = Union(Phase_Actual_TF_Days, Cell)
            Counter = Counter + 1
        Else
            Set Phase_Actual_TF_Days = Cell
            Counter = 1
        End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("AL2", "AL35")
    If Counter < PhaseNum Then
        If Not Phase_Actual_NPT_Days Is Nothing Then
            Set Phase_Actual_NPT_Days = Union(Phase_Actual_NPT_Days, Cell)
            Counter = Counter + 1
        Else
            Set Phase_Actual_NPT_Days = Cell
            Counter = 1
        End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("BF2", "BF35")
    If Counter < PhaseNum Then
        If Not Phase_Spacer_1 Is Nothing Then
            Set Phase_Spacer_1 = Union(Phase_Spacer_1, Cell)
            Counter = Counter + 1
        Else
            Set Phase_Spacer_1 = Cell
            Counter = 1
        End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("BG2", "BG35")
    If Counter < PhaseNum Then
        If Not Phase_Spacer_2 Is Nothing Then
            Set Phase_Spacer_2 = Union(Phase_Spacer_2, Cell)
            Counter = Counter + 1
        Else
            Set Phase_Spacer_2 = Cell
            Counter = 1
        End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("AO2", "AO35")
    If Counter < PhaseNum Then
        'We do not use the [If Not Cell.Value = "" Then] Text here because we need to plot the zeroes
            If Not Phase_Time_Plus_Minus Is Nothing Then
                Set Phase_Time_Plus_Minus = Union(Phase_Time_Plus_Minus, Cell)
                Counter = Counter + 1
            Else
                Set Phase_Time_Plus_Minus = Cell
                Counter = 1
            End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("AP2", "AP35")
    If Counter < PhaseNum Then
        If Not Phase_AFE_Cost Is Nothing Then
            Set Phase_AFE_Cost = Union(Phase_AFE_Cost, Cell)
            Counter = Counter + 1
        Else
            Set Phase_AFE_Cost = Cell
            Counter = 1
        End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("W2", "W35")
    If Counter < PhaseNum Then
        If Not Phase_Actual_Cost Is Nothing Then
            Set Phase_Actual_Cost = Union(Phase_Actual_Cost, Cell)
            Counter = Counter + 1
        Else
            Set Phase_Actual_Cost = Cell
            Counter = 1
        End If
    End If
Next

' Reset Counter
Counter = 0
For Each Cell In Range("AR2", "AR35")
    If Counter < PhaseNum Then
        'We do not use the [If Not Cell.Value = "" Then] Text here because we need to plot the zeroes
            If Not Phase_Cost_Plus_Minus Is Nothing Then
                Set Phase_Cost_Plus_Minus = Union(Phase_Cost_Plus_Minus, Cell)
                Counter = Counter + 1
            Else
                Set Phase_Cost_Plus_Minus = Cell
                Counter = 1
            End If
    End If
Next

'Find Scope Change Phases so they can be formatted differently on Phase Time and Phase Cost

Counter = 0
For Each Cell In Range("Y2", "Y35")
Counter = Counter + 1
    If Cell.Value = "Y" Then
        If SC1 = 0 Then
            SC1 = Counter
        Else
            If SC2 = 0 Then
                SC2 = Counter
            Else
                If SC3 = 0 Then
                    SC3 = Counter
                Else
                    If SC4 = 0 Then
                        SC4 = Counter
                    End If
                End If
            End If
        End If
    End If
Next
        


'''''Outdated due to no longer tracking cumulative on phase days and phase cost
'For Each Cell In Range("AO2", "AO35")
'    If Not Cell.Value = "" Then
'        Set LastPhase_Days = Cell
'    End If
'Next

'For Each Cell In Range("AR2", "AR35")
'    If Not Cell.Value = "" Then
'        Set LastPhase_Cost = Cell
'    End If
'Next

'Range("AO36") = LastPhase_Days
'Range("AR36") = LastPhase_Cost

'Updates Series in Phase Days Chart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("Phase Time").Activate

ActiveSheet.ChartObjects("Phase Time").Activate

ActiveChart.SeriesCollection("AFE").XValues = Phase_Names
ActiveChart.SeriesCollection("TF").XValues = Phase_Names
ActiveChart.SeriesCollection("NPT").XValues = Phase_Names
ActiveChart.SeriesCollection("Spacer 1").XValues = Phase_Names
ActiveChart.SeriesCollection("Spacer 2").XValues = Phase_Names

ActiveChart.SeriesCollection("AFE").Values = Phase_AFE_Days
ActiveChart.SeriesCollection("TF").Values = Phase_Actual_TF_Days
ActiveChart.SeriesCollection("NPT").Values = Phase_Actual_NPT_Days
ActiveChart.SeriesCollection("Spacer 1").Values = Phase_Spacer_1
ActiveChart.SeriesCollection("Spacer 2").Values = Phase_Spacer_2

'Reset all bars to blue (this is done incase the scope change phases change sequence)
ActiveChart.FullSeriesCollection("AFE").Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With

'Set each of the possible SC phases to orange
If SC1 > 0 Then
ActiveChart.FullSeriesCollection("AFE").Points(SC1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
End If

If SC2 > 0 Then
ActiveChart.FullSeriesCollection("AFE").Points(SC2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
End If

If SC3 > 0 Then
ActiveChart.FullSeriesCollection("AFE").Points(SC3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
End If

If SC4 > 0 Then
ActiveChart.FullSeriesCollection("AFE").Points(SC4).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
End If

ActiveSheet.ChartObjects("Phase Time plus minus").Activate
ActiveChart.SeriesCollection(1).XValues = Phase_Names
ActiveChart.SeriesCollection("Cumulative").Values = Phase_Time_Plus_Minus


'Updates Series in Phase Cost Chart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("Phase Cost").Activate

ActiveSheet.ChartObjects("Phase Cost").Activate
ActiveChart.SeriesCollection(1).XValues = Phase_Names
ActiveChart.SeriesCollection("AFE").Values = Phase_AFE_Cost
ActiveChart.SeriesCollection("Actual").Values = Phase_Actual_Cost

'Reset all bars to blue (this is done incase the scope change phases change sequence)
ActiveChart.FullSeriesCollection("AFE").Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With

'Set each of the possible SC phases to orange
If SC1 > 0 Then
ActiveChart.FullSeriesCollection("AFE").Points(SC1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
End If

If SC2 > 0 Then
ActiveChart.FullSeriesCollection("AFE").Points(SC2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
End If

If SC3 > 0 Then
ActiveChart.FullSeriesCollection("AFE").Points(SC3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
End If

If SC4 > 0 Then
ActiveChart.FullSeriesCollection("AFE").Points(SC4).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
End If


ActiveSheet.ChartObjects("Phase Cost Plus Minus").Activate
ActiveChart.SeriesCollection(1).XValues = Phase_Names
ActiveChart.SeriesCollection("Cumulative").Values = Phase_Cost_Plus_Minus

''''''''''''''''''''''''''''''''''''''End of Phase Days & Phase Costs ''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''' Phase vs Days '''''''''''''''''''''''''''''''''''''''''''''''

'The following work is necessary and complicated by the fact that trying to plot quotes ("") on a scatter plots 0
'So basically we have to recreate the ranges and reselect them with totally blank cells

Worksheets("Phase Export").Activate

'Set range of phase names for x axis of Phase vs Days graph
For Each Cell In Range("AS2", "AS35")
    If Not Cell.Value = "" Then
        If Not Phase_Names_NonWW Is Nothing Then
            Set Phase_Names_NonWW = Union(Phase_Names_NonWW, Cell)
        Else
            Set Phase_Names_NonWW = Cell
        End If
    End If
Next

'Cumulative Actual Phase Durations
For Each Cell In Range("AB2", "AB35")
    If Not Cell.Value = "" Then
        If Not Phase_Vs_Days Is Nothing Then
            Set Phase_Vs_Days = Union(Phase_Vs_Days, Cell)
        Else
            Set Phase_Vs_Days = Cell
        End If
    End If
Next

'Paste Cumulative Actual Phase Durations from above
Range("BA2", "BA35").ClearContents
Phase_Vs_Days.Copy
Range("BA2").PasteSpecial (xlPasteValues)

'Need to modify graph inputs to have blank cells at the end so that scatter plot shows up correctly
'Create counter based on number of phases to adequately select data range
Counter = 0
For Each Cell In Range("BA2", "BA35")
    If Counter <= PhaseNum Then
        If Not Phase_Vs_Days_Adj Is Nothing Then
            Set Phase_Vs_Days_Adj = Union(Phase_Vs_Days_Adj, Cell)
            Counter = Counter + 1
        Else
            Set Phase_Vs_Days_Adj = Cell
            Counter = 1
        End If
    End If
Next

'Remaining code to create projections, made easier by the fact that we have already declared these variables

'Create counter to determine # of phases that have occured already (until reaching current)
Counter = 0
For Each Cell In Range("Z2", "Z35")
    If Cell.Value = "Past" Then
        Counter = Counter + 1
    End If
Next

'Create P10, Mean, and P90 projection columns

'Add If statement so that these do not try to plot if the well is finished
If Not Projected_Depth Is Nothing Then
    Range("BB2", "BD35").ClearContents
    P10_Days.Copy
    Range("BB2").Offset(Counter).PasteSpecial (xlPasteValues)

    Mean_Days.Copy
    Range("BC2").Offset(Counter).PasteSpecial (xlPasteValues)

    P90_Days.Copy
    Range("BD2").Offset(Counter).PasteSpecial (xlPasteValues)
End If

'Create Proper projection ranges for Phase vs Days graph
Counter = 0
For Each Cell In Range("BB2", "BB35")
    If Counter <= PhaseNum Then
        If Not P10_Days_PvD Is Nothing Then
            Set P10_Days_PvD = Union(P10_Days_PvD, Cell)
            Counter = Counter + 1
        Else
            Set P10_Days_PvD = Cell
            Counter = 1
        End If
    End If
Next

Counter = 0
For Each Cell In Range("BC2", "BC35")
    If Counter <= PhaseNum Then
        If Not Mean_Days_PvD Is Nothing Then
            Set Mean_Days_PvD = Union(Mean_Days_PvD, Cell)
            Counter = Counter + 1
        Else
            Set Mean_Days_PvD = Cell
            Counter = 1
        End If
    End If
Next

Counter = 0
For Each Cell In Range("BD2", "BD35")
    If Counter <= PhaseNum Then
        If Not P90_Days_PvD Is Nothing Then
            Set P90_Days_PvD = Union(P90_Days_PvD, Cell)
            Counter = Counter + 1
        Else
            Set P90_Days_PvD = Cell
            Counter = 1
        End If
    End If
Next


Worksheets("Phase v Days").Activate

ActiveSheet.ChartObjects("Phase vs Days").Activate
ActiveChart.SeriesCollection(1).XValues = Phase_Names_NonWW
ActiveChart.SeriesCollection("Cumulative Actual Days").Values = Phase_Vs_Days_Adj
ActiveChart.SeriesCollection("P10 Projection").Values = P10_Days_PvD
ActiveChart.SeriesCollection("Mean Projection").Values = Mean_Days_PvD
ActiveChart.SeriesCollection("P90 Projection").Values = P90_Days_PvD


''''''''''''''''''''''''''''''''''''''''''' End of Phase vs Days ''''''''''''''''''''''''''''''''''''''''''''




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''' AFE Costs '''''''''''''''''''''''''''''''''''''''''''''''''''

'Sort AFE Costs ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Worksheets("AFE Export").Activate
    Set AFE_Table_Sort = Range("A1", Range("A1").End(xlDown).Offset(0, 10))
    Set AFE_Cost_Sort = Range("G1", Range("G1").End(xlDown))
    
    ActiveWorkbook.Worksheets("AFE Export").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("AFE Export").Sort.SortFields.Add Key:=AFE_Cost_Sort, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("AFE Export").Sort
        .SetRange AFE_Table_Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Determine Ranges for AFE Costs Chart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("AFE Cost").Activate

For Each Cell In Range("N3", "N36")
    If Not Cell.Value = "" Then
        If Not Category_Names Is Nothing Then
            Set Category_Names = Union(Category_Names, Cell)
        Else
            Set Category_Names = Cell
        End If
    End If
Next

For Each Cell In Range("O3", "O36")
    If Not Cell.Value = "" Then
        If Not Category_Actual_Cost Is Nothing Then
            Set Category_Actual_Cost = Union(Category_Actual_Cost, Cell)
        Else
            Set Category_Actual_Cost = Cell
        End If
    End If
Next

For Each Cell In Range("P3", "P36")
    If Not Cell.Value = "" Then
        If Not Category_AFE_Cost Is Nothing Then
            Set Category_AFE_Cost = Union(Category_AFE_Cost, Cell)
        Else
            Set Category_AFE_Cost = Cell
        End If
    End If
Next

'Updates Series in AFE Cost Chart '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("AFE Cost").Activate

ActiveSheet.ChartObjects("AFE Cost").Activate
ActiveChart.SeriesCollection(1).XValues = Category_Names
ActiveChart.SeriesCollection(2).XValues = Category_Names
ActiveChart.SeriesCollection("Total AFE").Values = Category_AFE_Cost
ActiveChart.SeriesCollection("Actual to Date").Values = Category_Actual_Cost

'''''''''''''''''''''''''''''''''''''''''''''' End of AFE Costs '''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''Updates Series in DvD vs Offsets Chart '''''''''''''''''''''''''''''''''

Worksheets("DvD vs Offsets").Activate
ActiveSheet.ChartObjects("DvD vs Offsets").Activate

ActiveChart.SeriesCollection("Actual").Values = Actual_Depth
ActiveChart.SeriesCollection("Actual").XValues = Actual_Days

''''''''''''''''''''''''''''''''''''''''''''' End of DvD vs Offsets '''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''' GO-36 '''''''''''''''''''''''''''''''''''''''''''''''''''

'Create GO-36 Data Ranges''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("DvD Export").Activate

'Set start value equal to 0
Set GO_Actual_Cost = Range("F1")
For Each Cell In Range("E2", "E1000")
    If Not Cell.Value = "" Then
        If Not GO_Actual_Cost Is Nothing Then
            Set GO_Actual_Cost = Union(GO_Actual_Cost, Cell)
        Else
            Set GO_Actual_Cost = Cell
        End If
    End If
Next


Worksheets("Phase Export").Activate

'Set start depth equal to 0 ****(Delete if Sidetrack)****
Set GO_DHC_Depth = Range("Z36")
For Each Cell In Range("F2", "F35")
    If Not Cell.Value = "" Then
        If Not GO_DHC_Depth Is Nothing Then
            Set GO_DHC_Depth = Union(GO_DHC_Depth, Cell)
        Else
            Set GO_DHC_Depth = Cell
        End If
    End If
Next


Worksheets("GO-36 Data").Activate

'Set start value equal to 0 ****(Delete if Sidetrack)****
Set GO_DHC = Range("I8")
For Each Cell In Range("D2", "D35")
    If Not Cell.Value = "" Then
        If Not GO_DHC Is Nothing Then
            Set GO_DHC = Union(GO_DHC, Cell)
        Else
            Set GO_DHC = Cell
        End If
    End If
Next


For Each Cell In Range("F2", "F35")
    If Not Cell.Value = "" Then
        If Not GO_P10 Is Nothing Then
            Set GO_P10 = Union(GO_P10, Cell)
        Else
            Set GO_P10 = Cell
        End If
    End If
Next

For Each Cell In Range("G2", "G35")
    If Not Cell.Value = "" Then
        If Not GO_Mean Is Nothing Then
            Set GO_Mean = Union(GO_Mean, Cell)
        Else
            Set GO_Mean = Cell
        End If
    End If
Next

For Each Cell In Range("H2", "H35")
    If Not Cell.Value = "" Then
        If Not GO_P90 Is Nothing Then
            Set GO_P90 = Union(GO_P90, Cell)
        Else
            Set GO_P90 = Cell
        End If
    End If
Next


'Update Series in GO-36 Chart''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("GO-36 Graph").Activate
ActiveSheet.ChartObjects("GO-36").Activate

ActiveChart.SeriesCollection("Actual").Values = Actual_Depth
ActiveChart.SeriesCollection("Actual").XValues = GO_Actual_Cost
    
'Add If statement so that these do not try to plot if the well is finished
If Not Projected_Depth Is Nothing Then
    ActiveChart.SeriesCollection("Forecast Min").Values = Projected_Depth
    ActiveChart.SeriesCollection("Forecast Min").XValues = GO_P10

    ActiveChart.SeriesCollection("Net Actual + Projection").Values = Projected_Depth
    ActiveChart.SeriesCollection("Net Actual + Projection").XValues = GO_Mean

    ActiveChart.SeriesCollection("Forecast Max").Values = Projected_Depth
    ActiveChart.SeriesCollection("Forecast Max").XValues = GO_P90
End If

ActiveChart.SeriesCollection("GO-36 DHC").Values = AFE_Depth
ActiveChart.SeriesCollection("GO-36 DHC").XValues = GO_DHC

''''''''''''''''''''''''''''''''''''''''''' End GO-36 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''' DvD '''''''''''''''''''''''''''''''''''''''''''''''''''''

'Delete Extra Offsets Series if less than 12 are selected''''''''''''''''''''''''''''''''''''''''''''''''''
    
Worksheets("DvD vs Offsets").Activate
ActiveSheet.ChartObjects("DvD vs Offsets").Activate

For Each Series In ActiveSheet.ChartObjects("DvD vs Offsets").Chart.SeriesCollection
    If Series.Name = "0" Then
        Series.Delete
    End If
Next
    
'''''''''''''''''''''''''''''''''''''''''''''''End DvD  '''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''Expected Dry Hole Days ''''''''''''''''''''''''''''''''''''''''''


'Find difference between spud date and moveon date
Worksheets("User Inputs").Activate
Spud = Range("B27").Value

Worksheets("Phase Export").Activate
Spud = Spud - Range("O2").Value

'Finds the highest value of actual days at the end of the final drill or evaluation
DHD = 0
For Each Cell In Range("C2", "C35")
    If Cell.Value = "DRILL" Or Cell.Value = "EVAL" Then
        If Not Cell.Offset(0, 19) = "" Then
            Num1 = Cell.Offset(0, 19).Value
            If Num1 > DHD Then
                DHD = Num1
            End If
        End If
        If Not Cell.Offset(0, 29) = "" Then
            Num2 = Cell.Offset(0, 29).Value
            If Num2 > DHD Then
                DHD = Num2
            End If
        End If
    End If
Next
    
Worksheets("Days v Depth").Activate
Range("F8").Value = DHD - Spud

''''''''''''''''''''''''''''''''''''''End Expected Dry Hole Days  '''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("User Inputs").Activate

End Sub
