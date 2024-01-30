Attribute VB_Name = "myUtilityMacros"
Option Explicit
Sub u_Create_Table_Of_Contents()

' Purpose: To create a table of contents for the Active Workbook.
' Trigger: Manual
' Updated: 2/18/2020

' Change Log:
'          2/18/20: Updated the code to make it a bit prettier

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim TOC_ws As Worksheet
        Application.DisplayAlerts = False
            If Evaluate("ISREF(TOC!A1)") = True Then ActiveWorkbook.Sheets("TOC").Delete
        Application.DisplayAlerts = True

        ActiveWorkbook.Worksheets.Add(Before:=Worksheets(1)).Name = "TOC"
            Set TOC_ws = ActiveWorkbook.Sheets("TOC")
    
    Dim ws As Worksheet

    Dim curRow As Long
    
    Dim bolBanded
        bolBanded = MsgBox("Banded rows cool bro?", vbYesNo + vbQuestion, "Move")

' -----------
' Run your code
' -----------

    With TOC_ws.Range("A1")
        .Value = "Table of Contents"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Underline = True
        .IndentLevel = 1
    End With
    
    With ActiveWindow
        .DisplayGridlines = False
        .Zoom = 150
        .DisplayHeadings = False
    End With
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "TOC" And ws.Visible = xlSheetVisible Then
            With TOC_ws
                curRow = .Range("A:A").Find("").Row
                .Range("A" & curRow).Value = ws.Name
                .Hyperlinks.Add Anchor:=.Range("A" & curRow), Address:="", SubAddress:="'" & ws.Name & "'!A1"
                .Range("A" & curRow).IndentLevel = 2
                .Rows(curRow).RowHeight = 17.5
                If curRow Mod 2 = 0 And bolBanded = 6 Then
                    .Range("A" & curRow).Interior.Color = RGB(240, 240, 240)
                    .Range("A" & curRow).Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
                    .Range("A" & curRow).Borders(xlEdgeTop).Color = RGB(190, 190, 190)
                End If
            End With
        End If
    Next ws
    
    TOC_ws.Columns("A").EntireColumn.AutoFit
    
End Sub
Sub u_Unhide_All_Worksheets()

' Purpose: To unhide all of the sheets in a workbook.
' Trigger: Manual
' Updated: 11/14/2019

' ***********************************************************************************************************************************

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
Next ws

End Sub
Sub u_Hide_All_Worksheets_Except_ActiveSheet()

' Purpose: To hide all of the sheets in a workbook.
' Trigger: Manual
' Updated: 11/14/2019

' ***********************************************************************************************************************************

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> ActiveSheet.Name Then ws.Visible = xlSheetHidden
Next ws

End Sub
Sub u_Hide_Charts()

' Purpose: To hide all of the chart sheets that are dark or light red.
' Trigger: Manual
' Updated: 12/19/2020

' Change Log:
'       7/24/2020: Added the dark red charts to be hidden
'       12/19/2020: Converted the colors to variables

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim ws As Worksheet
    
    ' Dim Reds
    Dim clrRed1 As Long
        clrRed1 = RGB(242, 220, 219)
        
    Dim clrRed15 As Long
        clrRed15 = RGB(236, 202, 201)
        
    Dim clrRed2 As Long
        clrRed2 = RGB(230, 184, 183)

' -----------------
' Hide the charts
' -----------------
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Tab.Color = clrRed1 Or ws.Tab.Color = clrRed2 Then ws.Visible = xlSheetHidden
    Next ws

End Sub
Sub u_Delete_Hidden_Sheets()

' Purpose: To delete all of the hidden Worksheets in a Workbook.
' Trigger: Manual
' Updated: 1/14/2020

' ***********************************************************************************************************************************

    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetHidden Then ws.Delete
    Next ws
    
    Application.DisplayAlerts = True

End Sub
Sub u_Unhide_All_Rows_Columns()

' Purpose: To unhide all of the rows / columns in each sheet in a workbook.
' Trigger: Manual
' Updated: 11/14/2019

' ***********************************************************************************************************************************

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Columns.EntireColumn.Hidden = False
    ws.Rows.EntireRow.Hidden = False
Next ws

End Sub
Sub u_Insert_Alternate_Rows()

' Purpose: To insert a row between each row in the selection.
' Trigger: Manual
' Updated: 11/14/2019

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim rng As Range
        Set rng = Selection
    
    Dim CountRow As Long
        CountRow = rng.EntireRow.Count
    
    Dim i As Long

' -----------
' Run your code
' -----------

    For i = 1 To CountRow
        ActiveCell.EntireRow.Insert
        ActiveCell.Offset(2, 0).Select
    Next i

End Sub
Sub u_Resize_Visible_Columns()
    
' Purpose: To autosize all of the columns in the ActiveSheet.
' Trigger: Manual
' Updated: 9/16/2019

' ***********************************************************************************************************************************
    
    ThisWorkbook.ActiveSheet.Cells.SpecialCells(xlCellTypeVisible).Columns.AutoFit
    
End Sub
Sub u_Clear_Slicers(Optional bolClearAllSlicers As Boolean, Optional strSinglePivotClearName)

' Purpose: To clear out the filters on the slicers in the pivot tables in the Active workbook.
' Trigger: Called
' Updated: 7/8/2022
'
' Change Log:
'       7/8/2022:   Initial Creation
'       7/11/2022:  Converted to a Utility Macro
'
' USE EXAMPLE: _
'   Call Macros.u_Clear_Slicers(strSinglePivotClearName:="Jason_Pivot")
'   Call Macros.u_Clear_Slicers(bolClearAllSlicers:=True)
'
' ***********************************************************************************************************************************************

Application.ScreenUpdating = False
On Error Resume Next

' -----------------
' Declare Variables
' -----------------

    Dim Slcr As SlicerCache

' ----------------
' Clear the pivots
' ----------------

    If bolClearAllSlicers = True Then
        For Each Slcr In ActiveWorkbook.SlicerCaches
            Slcr.ClearManualFilter
        Next Slcr
    Else
        ActiveSheet.PivotTables(strSinglePivotClearName).ClearAllFilters
    End If

Application.ScreenUpdating = True

End Sub
Sub u_Wipe_Checklist()

' Purpose: To remove the X marks in the Checklist ws.
' Trigger: Button
' Updated: 8/10/2021

' Change Log:
'       8/10/2021: Initial Creation

' ***********************************************************************************************************************************

myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    ' Assign Worksheets
    
    Dim wsChecklist As Worksheet
    Set wsChecklist = ThisWorkbook.Sheets("CHECKLIST")

    ' Declare Integers
       
    Dim int_LastRow As Long
        int_LastRow = wsChecklist.Cells(Rows.Count, "C").End(xlUp).Row

    Dim i As Long
    
' ------------
' Wipe the X's
' ------------

    With wsChecklist
                
        For i = 2 To int_LastRow
            If .Range("C" & i).Value = "X" Then .Range("C" & i).Value = ""
        Next i
        
    End With

myPrivateMacros.DisableForEfficiencyOff

End Sub

