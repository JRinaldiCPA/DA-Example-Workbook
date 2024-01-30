Attribute VB_Name = "XXX_VBA_TEMPLATE"
' Declare Workbooks
Dim wbSource As Workbook

' Declare Worksheets
Dim wsData As Worksheet
Dim wsLists As Worksheet
Dim wsDataChangeLog As Worksheet
Dim wsValidation As Worksheet

' Declare Cell References
Dim int_CurRow As Long
Dim int_LastRow As Long
Dim int_LastCol As Long

Dim int_CurRowValidation As Long

' Declare Strings
Dim strCustomer As String
Dim strNewFileFullPath As String
Dim strLastCol_wsData As String
Dim strUserID As String

' Declare Arrays
Dim arry_Header_Data() As Variant
Dim arry_Header_Lists() As Variant

Dim ary_Customers
Dim ary_SelectedCustomers
Dim ary_PM

' Declare Dictionaries
Dim dict_PMs As Scripting.Dictionary

' Declare Field References
Dim col_Customer As Long
Dim col_LOB As Long
Dim col_Region As Long

' Declare "Booleans"
Dim bol_ReliefStatus As String
Dim bol_Edit_Filter As Boolean
Dim bol_QC_Flags As Boolean

Option Explicit
Sub Cool_Story_Bro()

' ***********************************************************************************************************************************
'
' Author:       James Rinaldi
' Created Date: XX/XX/2021
' Last Updated: XX/XX/2021
'
' ----------------------------------------------------------------------------
'
' Purpose: To do some cool stuff brah...obvi.
'
' Trigger: Keyboard Shortcut - Ctrl + Shift + V
' Trigger: Ribbon Icon - Personal Macros > Functions
' Trigger: Called: o_01_MAIN_PROCEDURE
' Trigger: Daily_Review_UserForm > cmd_Done_Click
'
' Updated: XX/XX/2021

' Change Log:
'       XX/XX/2021

' ***********************************************************************************************************************************

' Purpose:  To give a bunch of my best options for doing different tasks in VBA.
' Trigger:  uf_Run_Process > cmd_Import_BB_Exceptions
' Updated:  1/3/2022

' Change Log:
'       10/20/2021: Initial Creation
'       1/3/2022:   Created the code to import the BB Ticklers

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler

Call myPrivateMacros.DisableForEfficiency

' ----------------------------------
' Set the current directory and path
' ----------------------------------

    ' Assign Objects
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    

On Error Resume Next
    ChDrive ThisWorkbook.Path
        If objFSO.FolderExists(ThisWorkbook.Path & "\(Source Data)") = True Then
            ChDir ThisWorkbook.Path & "\(Source Data)"
        Else
            ChDir ThisWorkbook.Path
        End If
On Error GoTo 0

' -----------------
' Declare Variables
' -----------------

    Dim NewRef As String
        Select Case Asc(Right(OldRef, 1))
            Case 65 To 90, 97 To 122
                NewRef = Left(OldRef, Len(OldRef) - 1) & Chr(Asc(Right(OldRef, 1)) + 1)
            Case Else
                NewRef = Left(OldRef, InStrRev(OldRef, ".")) & Right(OldRef, (Len(OldRef) - InStrRev(OldRef, "."))) + 1
            End Select

' ---------------------
' Assign your variables
' ---------------------

    ' Assign Workbooks
    
    ' Use GetOpenFileName to prompt the user, plus it shows them only Excel Workbooks
    Dim wbTracker As Workbook
        Set wbTracker = Workbooks.Open(Application.GetOpenFilename( _
        Title:="Select the current CMML CV Tracker", FileFilter:="Excel Workbooks (*.xls*;*.csv),*.xls*;*.csv"), _
        UpdateLinks:=False, ReadOnly:=False)
        
    ' Propmt the user for a PDF file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the applicable Cover Page in PDF"
        .AllowMultiSelect = False
        .Filters.Add "PDF Files", "*.pdf", 1
        .InitialFileName = ThisWorkbook.Path & "" \ ""
        .Show
        strSourceFilePath = .SelectedItems.Item(1)
    End With
        
    ' Assign Worksheets
    
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("Dashboard Review")
    
    Dim wsValidation As Worksheet
    Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
    
    ' Assign Integers
    
    Dim int_PM_Attest_Complete As Long: int_PM_Attest_Complete = 1 ' Declare / Assign in one line
    
    ' Assign Cell References
       
    Dim int_LastRow As Long
        int_LastRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row
        int_LastRow = Application.WorksheetFunction.Max(ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, ws.Cells(ws.Rows.Count, "C").End(xlUp).Row)
        
        int_LastRow = WorksheetFunction.Max( _
            ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
            ws.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row, _
            ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row)
            
        If int_LastRow = 1 Then int_LastRow = 2
       
    Dim int_LastCol As Long
        int_LastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
        
        int_LastCol = WorksheetFunction.Max( _
            ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column, _
            ws.Cells(int_LastRow, ws.Columns.Count).End(xlToLeft).Column, _
            ws.Rows(1).Find("").Column - 1)
            
        If int_LastCol = 1 Then int_LastCol = 2
       
    Dim int_CurRow As Long
        int_CurRow = ThisWorkbook.ActiveSheet.Range("A:A").Find("").Row
        int_CurRow = wsValidation.Cells(wsValidation.Rows.Count, "A").End(xlUp).Row + 1
        int_CurRow = Evaluate("MATCH(TRUE,INDEX(ISBLANK('[To Do.xlsm]DA Requests'!A:A),0),0)")
        int_CurRow = fx_Find_Row(ws:=ws_Target, str_Target:="")
                
    Dim intLastUsedRow_Dest As Long
        intLastUsedRow_Dest = wsDest.Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    ' Assign Field References
    Dim arry_Header_wsData() As Variant
        arry_Header_wsData = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, int_LastCol)))
    
    Dim col_LOB As Long
        col_LOB = fx_Create_Headers("LOB", arry_Header_wsData)
                
    ' Declare Colors
        
    Dim clrGreen1 As Long
        clrGreen1 = RGB(235, 241, 222)
        
    Dim clrGreen2 As Long
        clrGreen2 = RGB(216, 228, 188)
    
    Dim clrOrange2 As Long
        clrOrange2 = RGB(252, 213, 180)
            
    Dim clrRed1 As Long
        clrRed1 = RGB(242, 220, 219)
    
    Dim clrRed15 As Long
        clrRed15 = RGB(236, 202, 201)
    
    'Dim Loop Variables
    
    Dim i As Long

    Dim intRunCMML As Long
        intRunCMML = MsgBox("Do you want to update the CMML data?", vbYesNo + vbQuestion, "Update CMML")
        If intRunCMML = vbYes Then GoTo RunCode
    
    ' Use the V_RunDate named range that lives in wsValidation as part of a calc
    'Dim dtModMaturity As Date
        dtModMaturity = DateAdd("M", 1, DateValue(Month([v_RunDate]) & "/01/2020")) - 1
    
' --------------------------------------------------
' Fill the Dictionary with the correct Officer names
' --------------------------------------------------
        
    ' Assign Dictionary
    
    'Dim dict_Officers As Scripting.Dictionary
    'Set dict_Officers = New Scripting.Dictionary
    Dim dict_Officers As New Scripting.Dictionary ' Dim / Set in one step
    
        dict_Officers.CompareMode = TextCompare
        
    'Fill the Dictionary
    
    With wsLists
            
        For i = 2 To int_LastRow_Lists
            dict_Officers.Add Key:=.Cells(i, col_OldOfficer).Value2, Item:=.Cells(i, col_NewOfficer).Value2
        Next i
    
    End With
    
    ' Replace the Officer with a more accurate name

    With wsCurrent
        For i = 2 To int_LastRow

            If dict_Officers.Exists(.Cells(i, col_Officer).Value2) Then
                .Cells(i, col_Officer) = dict_Officers.Item(.Cells(i, col_Officer).Value2)
            End If

        Next i
    End With

' -------------------------------------------
' Loop through one Dictionary to fill another
' -------------------------------------------

    For Each val In dict_File
        If InStr(1, val, cmb_DynamicSearch.Value, vbTextCompare) Then 'If the name is similar then add to the list
            Me.lst_RefX.AddItem val
        End If
    Next val

' -----------------
' Fix Header Values
' -----------------

    wsSource.Range(wsSource.Cells(intHeaderRow, 1), wsSource.Cells(intHeaderRow, 99)).Find("RTB High(Branch)").Value2 = "RTB High"
    wsSource.Range(wsSource.Cells(intHeaderRow, 1), wsSource.Cells(intHeaderRow, 99)).Find("RTB Low (User Defined 14)").Value2 = "RTB Low"

' ----------
' Cool Stuff
' ----------

    ' For Each Cell loop
    For Each cell In Selection
        If cell.Font.Bold = "False" Then cell.InsertIndent 1
    Next cell

    ' For i loop
    For i = 2 To int_LastRow
        If wsCurrent.Range("P" & i).Value = "Yes" Then
            wsCurrent.Range(Cells(i, 1), Cells(i, int_LastCol)).Interior.Color = RGB(208, 197, 221)
        End If

    Next i

    ' If the Risk Trend is still open close it
    If Evaluate("ISREF(" & "'[" & wbSource.Name & "]" & wbSource.Sheets(1).Name & "'" & "!A1)") = True Then
        wbSource.Close savechanges:=False
    End If

Call myPrivateMacros.DisableForEfficiencyOff

'ThisWorkbook.Worksheets("CHECKLIST").Range("chk_o3_Import_BB_Data").Value = "X"

Exit Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

ErrorHandler:
    MsgBox ("Something went awry, if you hit Cancel then redo the process." & Chr(10) & "Otherwise, reach out to James Rinaldi in Audit for a fix.")
    MsgBox Err.Source & ": The following error occured: " & Err.Description
    
    Call myPrivateMacros.DisableForEfficiencyOff
    Exit Sub

' --------------------
' Message Box Examples
' --------------------

    MsgBox Title:="Validation Totals Match", _
        Buttons:=vbOKOnly, _
        Prompt:="The validation totals match, you're golden. " & Chr(10) & Chr(10) _
        & "1st Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "1st Validation Count: " & Format(int1stCount, "0,0")
       
    MsgBox Title:="Validation Totals Don't Match", _
        Buttons:=vbCritical, _
        Prompt:="The validation totals from the Sageworks Dashboard Dump don't match what was imported. " _
        & "Please review the totals in the Validation worksheet to determine what went awry. " _
        & "Once the issues has been identifed talk to James to fix / reimport the data." & Chr(10) & Chr(10) _
        & "1st Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "1st Validation Total Variance: " & Format(int1stTotal - int2ndTotal, "$#,##0") & Chr(10) & Chr(10) _
        & "1st Validation Count: " & Format(int1stCount, "0,0") & Chr(10) _
        & "1st Validation Count Variance: " & Format(int1stCount - int2ndCount, "0,0")

    MsgBox Title:="Validation Totals Match", _
        Buttons:=vbOKOnly + vbInformation, _
        Prompt:="The validation totals match, you're golden. " & Chr(10) & Chr(10) _
        & "Validation Total: " & Format(int1stTotal, "$#,##0")


    If dict_Anomalies("Anomalies Found Count") > dict_Anomalies("Unique Anomalies Found Count") Then
        MsgBox Title:="Ah Ah Ah", Buttons:=vbExclamation, _
        Prompt:="You need to address the highlighted fields before attesting, including: " & Chr(10) & fx_Remaining_Anomalies_List(Target.Row)
    ElseIf wsData.Cells(Target.Row, col_CovCompliance) = "" Then
        MsgBox Title:="Ah Ah Ah", Buttons:=vbExclamation, Prompt:="You need to complete the Covenant Compliance attestation first."
    Else
        uf_PM_Attestation.Show
    End If

End Sub
Sub o_01_MAIN_PROCEDURE()

' Purpose:  To initialize the userform, including adding in the data from the arrays.
' Trigger:  Workbook Open
' Updated:  11/16/2020
' Author:   James Rinaldi

' Change Log:
'       3/23/2020:  Intial Creation
'       8/19/2020:  Added the logic to exclude the exempt customers
'       11/16/2020: I added the autofilter to hide the CRE and ABL LOBs

' ***********************************************************************************************************************************

Debug.Print "o_01_MAIN_PROCEDURE"

Call myPrivateMacros.DisableForEfficiency
    
Call XXX_VBA_TEMPLATE.o_02_Assign_Private_Variables
    
' -----------------
' Declare Variables
' -----------------

    ' Use this to prompt the user for which action to take, so multiples actions can be in the same Sub
    Dim intRun_PrepforStephanie As Long
        intRun_PrepforStephanie = MsgBox("Do you want to Prep the file for Stephanie?", vbYesNo + vbQuestion, "Prep for Stephanie")
        If intRun_PrepforStephanie = vbYes Then GoTo RunCode
    
    'If nothing else has been triggered, abort
    GoTo Abort
        
' ----------------------
' Run the update process
' ----------------------
        
RunCode:
        
    If intRun_PrepforStephanie = vbYes Then
        Call o_21_Prep_For_Stephanie
    End If

' --------------------------------------
' Let the user know the process finished
' --------------------------------------

    MsgBox _
        Title:="Import Prior Past Due", _
        Buttons:=vbOKOnly + vbInformation, _
        Prompt:="The prior Past Due ACH Review data was imported, please review the control totals in the VALIDATION tab to ensure all data was imported."
    
    'Application.GoTo Reference:=wsValidation.Range("A1"), Scroll:=False

Abort:

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To assign the Private Variables that were Declared above the line.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 5/1/2021

' Change Log:
'       5/1/2021:   Intial Creation

' ***********************************************************************************************************************************
    
Debug.Print "o_02_Assign_Private_Variables"

    ChDrive ThisWorkbook.Path
        ChDir ThisWorkbook.Path & "\(Source Data)"

' ---------------------
' Assign your variables
' ---------------------
    
    ' Assign Workbooks
    
    Set wbSource = fx_Open_Workbook(strPromptTitle:="Select the current Adjusted Risk Trend")
    
    ' Assign Worksheets

    If Evaluate("ISREF(" & "'Data'" & "!A1)") = True Then
        Set wsData = ThisWorkbook.Sheets("Data")
    End If
    
    If Evaluate("ISREF(" & "'Lists'" & "!A1)") = True Then
        Set wsLists = ThisWorkbook.Sheets("LISTS")
    End If
    
    If Evaluate("ISREF(" & "'Change Log'" & "!A1)") = True Then
        Set wsDataChangeLog = ThisWorkbook.Sheets("DATA CHANGE LOG")
    End If
    
    If Evaluate("ISREF(" & "'VALIDATION'" & "!A1)") = True Then
        Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
    End If
    
    If Evaluate("ISREF(" & "'[" & ThisWorkbook.Name & "]" & "Validation'" & "!A1)") = True Then
        Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
    End If
    
    ' Assign Cell References
    
    int_LastCol = WorksheetFunction.Max( _
        wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column, _
        wsData.Rows(1).Find("").Column - 1)
        
    int_LastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    int_LastCol_wsLists = wsLists.Cells(1, wsLists.Columns.Count).End(xlToLeft).Column
        
    int_LastRow_wsArrays = wsArrays.Cells(wsArrays.Rows.Count, "A").End(xlUp).Row
        If int_LastRow_wsArrays = 1 Then int_LastRow_wsArrays = 2

    ' Assign Strings
    strUserID = Application.UserName
    strLastCol_wsData = Split(Cells(1, int_LastCol).Address, "$")(1)

    ' Assign Arrays
    
    arry_Header_Data = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, int_LastCol)))
    arry_Header_Lists = Application.Transpose(wsLists.Range(wsLists.Cells(1, 1), wsLists.Cells(1, int_LastCol_wsLists)))

    ' Assign "Ranges"
    
    col_Customer = fx_Create_Headers_v2("Customer", arry_Header_Data)
    col_LOB = fx_Create_Headers_v2("LOB", arry_Header_Data)
    col_Region = fx_Create_Headers_v2("Region", arry_Header_Data)
    
    col_Exposure = fx_Create_Headers_v2("Webster Commitment (000's) - Gross Exposure", arry_Header_Data)
    col_Outstanding = fx_Create_Headers_v2("Webster Outstanding (000's) - Book Balance", arry_Header_Data)
    
    ' Assign List "Ranges"
    col_LOB1_Lists = fx_Create_Headers_v2("LOB1", arry_Header_Lists)
    col_Cust1_Lists = fx_Create_Headers_v2("Customer1", arry_Header_Lists)
    col_PM1_Lists = fx_Create_Headers_v2("PM1", arry_Header_Lists)

    ' Assign Booleans
    bolPrivilegedUser = fx_Privileged_User
    
End Sub
Sub o_01_XXX_Assign_Variables()

    ' Use this if using my standard folder structure


    ' Dim Objects
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Dim Ranges
    
    Dim rngData As Range
    Set rngData = wsData.Range(wsData.Cells(1, 1), wsData.Cells(9999, 10))
    
    ' Dim Integers
    
    Dim int_PM_Attest_Complete As Long: int_PM_Attest_Complete = 1 ' Declare / Assign in one line

    'Dim Dates
    
    Dim dtOpenDate As Date
        dtOpenDate = CDate("12/31/2099")

    ' Assign Strings
        
End Sub

Sub o_11_Wipe_Old_Data()

' Purpose: To import the current data into the Template.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 12/18/2021

' Change Log:
'       12/18/2021: Intial Creation

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Integers
    
    Dim int_LastRow As Long
            
    Dim int_LastCol As Long

' -------------
' Wipe Old Data
' -------------

    ' Wipe data from QC Review ws
    With wsQCReview
        int_LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            If int_LastRow = 1 Then int_LastRow = 2
        int_LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    
        .Range(.Cells(2, 1), .Cells(int_LastRow, int_LastCol)).ClearContents
        .Range(.Cells(2, 1), .Cells(int_LastRow, int_LastCol)).Interior.Color = xlNone
    End With ' wsQCReview
    
End Sub
Sub o_12_Import_Data()

' Purpose: To import the current data into the Template.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 12/18/2021

' Change Log:
'       12/18/2021: Intial Creation

' ***********************************************************************************************************************************

    ' Set the current folder path to be where the template is located
        ChDrive ThisWorkbook.Path
            ChDir ThisWorkbook.Path

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Workbooks
    Set wbQCReview = Workbooks.Open(Application.GetOpenFilename _
    (Title:="Select the Excel that contains the data you are reviewing for the QC Review", FileFilter:="Excel Workbooks (*.xls*;*.csv),*.xls*;*.csv"))

    ' Declare Worksheets

     Set wsQCReview_Orig = wbQCReview.Worksheets(1)

    ' Declare Integers
    
    Dim int_LastRow As Long
    Dim int_LastCol As Long
    Dim intHeaderRow As Long

    Dim int_LastRow_wsValidation As Long
        int_LastRow_wsValidation = wsValidation.Cells(wsValidation.Rows.Count, "A").End(xlUp).Row
        If int_LastRow_wsValidation = 1 Then int_LastRow_wsValidation = 2

' -----------
' Wipe out the old validation data
' -----------
    
    wsValidation.Range("A2:E" & int_LastRow_wsValidation).ClearContents

' -----------
' Remove the Header and Footer Rows
' -----------

    If intFooterRow < int_LastRow Then wsSource.Rows(intFooterRow & ":" & int_LastRow).Delete
    
    If intHeaderRow > 1 Then wsSource.Range("A1:A" & intHeaderRow - 1).EntireRow.Delete

' -----------
' Copy the data from the QC Review workbook
' -----------
    
    With wsQCReview_Orig
    
        ' Declare Variables
    
        wsQCReview_Orig.Activate
    
        intHeaderRow = .Range("A:A").Find("Account Number").Row
            .Rows("1:" & intHeaderRow - 1).Delete 'Delete Header
    
        int_LastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
            .Rows(int_LastRow + 1 & ":" & int_LastRow + 4).Delete 'Delete Footer
        
        int_LastCol = .Cells(intHeaderRow, .Columns.Count).End(xlToLeft).Column
        
        ' Copy in the data and close workbook
            
    Call fx_Copy_in_Data_for_Matching_Fields(wsQCReview_Orig, wsQCReview, "o_1_Import_Data", "Princ Bal")
        
    End With ' wsQCReview_Orig
        
End Sub
Sub o_12_Import_Data_Basic()

' Purpose: To import all of the data from the BO 4155 Report.
' Trigger: Called by Main Procedure
' Updated: 11/3/2021

' Change Log:
'       11/3/2021:  Intial Creation

' ***********************************************************************************************************************************

'Sets the current directory and path
    
    ChDrive ThisWorkbook.Path
        ChDir ThisWorkbook.Path

' ----------------------
' Declare your variables
' ----------------------

    'Dim Workbooks / Worksheets

    Dim wbSource As Workbook
        Set wbSource = Workbooks.Open(Application.GetOpenFilename(Title:="Select the current BO 4155 report (ACH Limits)"))
    
    Dim wsSource As Worksheet
        Set wsSource = wbSource.Sheets(1)
        
    Dim wbDest As Workbook
        Set wbDest = ThisWorkbook
    
    Dim wsDest As Worksheet
        Set wsDest = wbDest.Sheets("Data")
            
    'Dim Integers
    
    Dim intHeaderRow As Long
        intHeaderRow = wsSource.Range("A:A").Find("Cost Ctr").Row
    
    Dim intFooterRow As Long
        intFooterRow = wsSource.Range("A:A").Find("GRAND TOTAL").Row
    
    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max(wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row, wsSource.Cells(wsSource.Rows.Count, "F").End(xlUp).Row)
    
    'Dim Strings (for function)
            
    Dim strControlTotalHeader As String
        strControlTotalHeader = "Credit Daily Limit (Application)"

    Dim str_ModuleName As String
        str_ModuleName = "o_1_Import_Data"

' -----------
' Remove the Header and Footer Rows
' -----------

    If intFooterRow < int_LastRow Then wsSource.Rows(intFooterRow & ":" & int_LastRow).Delete
    
    If intHeaderRow > 1 Then wsSource.Range("A1:A" & intHeaderRow - 1).EntireRow.Delete
    
' -----------
' Update the data
' -----------

    Call fx_Copy_in_Data_for_Matching_Fields(wsSource, wsDest, str_ModuleName, strControlTotalHeader)

End Sub
Sub o_21_Apply_Formulas()

' Purpose: To copy the formulas from the FIELDS - FORMULAS ws to the All Customers ws.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 9/3/2021

' Change Log:
'       9/3/2021:   Intial Creation

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------
   
    ' Declare Worksheets
       
    Dim wsAllCust As Worksheet
        Set wsAllCust = ThisWorkbook.Sheets("ALL Customers")
            wsAllCust.AutoFilter.ShowAllData
            
    Dim wsFields As Worksheet
        Set wsFields = ThisWorkbook.Sheets("FIELDS - FORMULAS")
   
    ' Declare Integers
    
    Dim intFirstCol_wsAllCust As Long
        intFirstCol_wsAllCust = wsAllCust.Range("1:1").Find("PPP Balance").Column
        
    Dim int_LastCol_wsAllCust As Long
        int_LastCol_wsAllCust = wsAllCust.Cells(1, wsAllCust.Columns.Count).End(xlToLeft).Column
        
    Dim intFirstRow_wsFields As Long
        intFirstRow_wsFields = wsFields.Range("C:C").Find("PPP Balance").Row

    Dim int_LastRow_wsFields As Long
        int_LastRow_wsFields = wsFields.Range("C:C").Find("Adj. Mod Balance").Row + 1
        
    Dim int_LastRow_wsAllCust As Long
        int_LastRow_wsAllCust = wsAllCust.Cells(wsAllCust.Rows.Count, "F").End(xlUp).Row
            If int_LastRow_wsAllCust = 1 Then int_LastRow_wsAllCust = 2
    
    ' Declare Loop Variables
    
    Dim x As Long
    
    Dim y As Long
    
    Dim strFormula As String
    
' -----------
' Copy the formulas into the All Customers ws
' -----------

    For x = intFirstCol_wsAllCust To int_LastCol_wsAllCust
    
        For y = intFirstRow_wsFields To int_LastRow_wsFields
        
            If wsAllCust.Cells(1, x) = wsFields.Cells(y, 3) Then
                strFormula = wsFields.Cells(y, 5).Value2
                wsAllCust.Range(wsAllCust.Cells(2, x), wsAllCust.Cells(int_LastRow_wsAllCust, x)).Formula = strFormula
    
                Exit For
            End If
        Next y
    
    Next x

' -----------
' Copy the formulas back values only
' -----------

    wsAllCust.Range(wsAllCust.Cells(2, intFirstCol_wsAllCust), wsAllCust.Cells(int_LastRow_wsAllCust, int_LastCol_wsAllCust)).Value2 = _
    wsAllCust.Range(wsAllCust.Cells(2, intFirstCol_wsAllCust), wsAllCust.Cells(int_LastRow_wsAllCust, int_LastCol_wsAllCust)).Value2

End Sub
Sub o_22_Clean_4155_Data()

' Purpose: To clean the data that was imported from the BO 4155.
' Trigger: Called by Main Procedure
' Updated: 11/3/2021

' Change Log:
'   11/3/2021:  Intial Creation

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    'Dim Worksheets
    
    Dim wsData As Worksheet
        Set wsData = ThisWorkbook.Sheets("Data")

    'Dim Integers

    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max(wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row, wsData.Cells(wsData.Rows.Count, "F").End(xlUp).Row)
        
    Dim int_LastCol As Long
        int_LastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    
    'Dim Array
    
    Dim aryData As Variant
    
    'Dim Ranges / "Ranges"
    
    Dim rngHeader As Range
        Set rngHeader = wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, int_LastCol))
    
    Dim col_LOB As Long
        col_LOB = fx_Create_Headers("Line of Business", rngHeader)
        
    Dim col_CustName As Long
        col_CustName = fx_Create_Headers("Customer Name", rngHeader)
        
    Dim col_Officer As Long
        col_Officer = fx_Create_Headers("Officer", rngHeader)
        
    Dim col_IMAcctStatus As Long
        col_IMAcctStatus = fx_Create_Headers("IM Acct Status", rngHeader)
        
    Dim col_TransPoint As Long
        col_TransPoint = fx_Create_Headers("Transmission Point", rngHeader)
        
    'Dim Loop Variables
        
    Dim i As Long
    
' -----------
' Remove the inapplicable accounts
' -----------
    
    With wsData
        For i = int_LastRow To 2 Step -1

            If .Cells(i, col_CustName) <> "" Then
                If .Cells(i, col_Officer) = "Albert Wang" Then
                    .Rows(i).Delete ' Remove Al Wang Accounts
                ElseIf .Cells(i, col_IMAcctStatus) = "I" Or .Cells(i, col_IMAcctStatus) = "" Then
                    .Rows(i).Delete ' Remove Inactive Accounts (I) or Webster GL Accounts (Blanks)
                ElseIf Right(.Cells(i, col_TransPoint), 3) = "PRE" Then
                    .Rows(i).Delete
                End If
            Else
                If .Cells(i, 1) <> "" Then
                    .Rows(i).Delete ' Remove Total Rows
                End If
            End If

        Next i
    End With

' -----------
' Update the LOB for Consumer Deposits
' -----------

    With wsData
        For i = 2 To int_LastRow
            If .Cells(i, col_LOB) = "Consumer Deposits" Then .Cells(i, col_LOB) = "Business Banking"
        Next i
    End With

End Sub
Sub o_21_Copy_Necessary_Data_to_wsData()

' Purpose: To copy only the fields that we need into the Data worksheet from Source.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 7/17/2021

' Change Log:
'   4/8/2021: Intial Creation
'   6/30/2021: Updated to skip the creation of the Source wb
'   7/17/2021: Updated to inculde the flag for PPP based on C/I's criteria, and to overwrite the LOB for PPP loans

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------
    
    'Dim Strings
    
    Dim strColName As String
        
    'Dim Integers
    
    Dim intColNum_Source As Long
    
    Dim intColNum_Data As Long
    
    Dim intNewColNum As Long

    Dim i As Long

    'Dim Arrays
    
    Dim ary_Data_Temp
    
' -----------
' Copy over the data
' -----------
        
    'Clear out the old data
    wsData.Range(wsData.Cells(2, 1), wsData.Cells(int_LastRow_wsData, int_LastCol_wsData)).ClearContents
        
    For intColNum_Source = 1 To int_LastCol_Source
        strColName = ary_Source(1, intColNum_Source)
        
        If Not wsData.Range("1:1").Find(strColName) Is Nothing And strColName <> "D_Officer Primary" Then
            intNewColNum = wsData.Range("1:1").Find(strColName).Column
                    
            wsData.Range(wsData.Cells(1, intNewColNum), wsData.Cells(UBound(ary_Source), intNewColNum)) = _
            Application.Index(ary_Source, 0, intColNum_Source)
                        
        End If
        
    Next intColNum_Source
        
    int_LastRow_wsData = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row
        
' ---------------
' Format the data
' ---------------

    With wsData

    For intColNum_Data = 1 To int_LastCol_wsData

        strColName = .Cells(1, intColNum_Data).Value2

        If strColName = "D_Direct Outstanding" Or _
           strColName = "D_Daily Gross Exposure" _
        Then
            .Range(.Cells(2, intColNum_Data), .Cells(, intColNum_Data)).NumberFormat = "$#,##0"
        ElseIf strColName = "D_Effective Date" Then
            .Range(.Cells(2, intColNum_Data), .Cells(int_LastRow_wsData, intColNum_Data)).NumberFormat = "mm/dd/yyyy"
        ElseIf strColName = "D_Account Number" Then
            .Range(.Cells(2, intColNum_Data), .Cells(int_LastRow_wsData, intColNum_Data)).NumberFormat = "0"
        End If
        
    Next intColNum_Data
    
    End With

' -----------
' Copy the data into an array
' -----------

    ary_Data_Temp = wsData.Range(wsData.Cells(1, 1), wsData.Cells(int_LastRow_wsData, int_LastCol_wsData))

' -----------
' Create the Virtual Fields
' -----------
        
    For i = 2 To UBound(ary_Data_Temp)
        
        ' V_PPP_LOAN_FLAG field
        If ary_Data_Temp(i, col_TypeCode_Data) = "U18" _
        And ary_Data_Temp(i, col_CostCenter_Data) = "161" _
        And ary_Data_Temp(i, col_OpenDt_Data) >= DateValue("4/1/2020") Then
            ary_Data_Temp(i, col_PPPFlag_Data) = "Yes"
        End If
        
        ' If it's a PPP loan then overwrite the values for LOB BEFORE creating the Helper field
        If ary_Data_Temp(i, col_PPPFlag_Data) = "Yes" Then ary_Data_Temp(i, col_LOB_Data) = "PPP"
        
        ' V_HELPER field
        If ary_Data_Temp(i, col_CustName_Data) <> "" Then
            ary_Data_Temp(i, col_Helper_Data) = ary_Data_Temp(i, col_LOB_Data) & " - " & UCase(ary_Data_Temp(i, col_CustName_Data)) 'Upper
        End If
    
    Next i

' -----------
' Output results and sort by Helper
' -----------
    
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(int_LastRow_wsData, int_LastCol_wsData)) = ary_Data_Temp
    
    With wsData
        
        .Range(.Cells(1, 1), .Cells(int_LastRow_wsData, int_LastCol_wsData)).Sort Key1:=.Cells(1, col_Helper_Data), Order1:=xlAscending, Header:=xlYes
    
    End With
    
' -----------
' Output the control totals
' -----------

    With wsValidation
        int_CurRowValidation = .Cells(wsValidation.Rows.Count, "A").End(xlUp).Row
    
        .Range("A" & int_CurRowValidation + 1) = Now
        .Range("B" & int_CurRowValidation + 1) = "o_1_Import_6330_Data"
        .Range("C" & int_CurRowValidation + 1) = "Source Array"
        .Range("D" & int_CurRowValidation + 1) = Format(Application.WorksheetFunction.Sum(Application.Index(ary_Source, 0, col_Outstanding_Source)), "$#,##0")
        .Range("E" & int_CurRowValidation + 1) = Format(UBound(ary_Source) - 1, "0,0")
        
        .Range("A" & int_CurRowValidation + 2) = Now
        .Range("B" & int_CurRowValidation + 2) = "o_1_Import_6330_Data"
        .Range("C" & int_CurRowValidation + 2) = "wsData"
        .Range("D" & int_CurRowValidation + 2) = Format(Application.WorksheetFunction.Sum(wsData.Range("P2:P" & int_LastRow_wsData)), "$#,##0")
        .Range("E" & int_CurRowValidation + 2) = Format(int_LastRow_wsData - 1, "0,0")
    
    End With
    
    Debug.Print "Total from wsSource:   " & Format(Application.WorksheetFunction.Sum(Application.Index(ary_Source, 0, col_Outstanding_Source)), "$#,##0")
    Debug.Print "Total in wsData:       " & Format(Application.WorksheetFunction.Sum(wsData.Range("P2:P" & int_LastRow_wsData)), "$#,##0")
        
End Sub
Sub o_2_Manipulate_Data()

' Purpose: To manipulate the data to calculate the two virtual fields, and sort.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 12/18/2020

' Change Log:
'       12/18/2020: Intial Creation
'       12/18/2020: Disabled the 00000 padding for Account Number, no longer needed now that I am pulling in everything as values
'       12/18/2020: Revamped to replace all of the formulas with code

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Assign Private Integers
    
        int_LastRow_QC_Review = wsQCReview.Cells(wsQCReview.Rows.Count, "A").End(xlUp).Row
        int_LastRow_NewReport = wsNewReport.Cells(wsNewReport.Rows.Count, "A").End(xlUp).Row
        int_LastRow_HUD = wsHUD.Cells(wsHUD.Rows.Count, "B").End(xlUp).Row

    ' Dim Header Arrays
    
        Dim arry_Header_QCReview() As Variant
            arry_Header_QCReview = Application.Transpose(wsQCReview.Range(wsQCReview.Cells(1, 1), wsQCReview.Cells(1, int_LastCol_QC_Review)))

        Dim arry_Header_NewReport() As Variant
            arry_Header_NewReport = Application.Transpose(wsNewReport.Range(wsNewReport.Cells(1, 1), wsNewReport.Cells(1, int_LastCol_NewReport)))

        Dim arry_Header_HUD() As Variant
            arry_Header_HUD = Application.Transpose(wsHUD.Range(wsHUD.Cells(1, 1), wsHUD.Cells(1, int_LastCol_HUD)))

    ' Dim Data Arrays

        Dim arryQCData() As Variant
            arryQCData = wsQCReview.Range(wsQCReview.Cells(1, 1), wsQCReview.Cells(int_LastRow_QC_Review, int_LastCol_QC_Review))

        Dim arryNewData() As Variant
            arryNewData = wsNewReport.Range(wsNewReport.Cells(1, 1), wsNewReport.Cells(int_LastRow_NewReport, int_LastCol_NewReport))
    
        Dim arryHUDData() As Variant
            arryHUDData = wsHUD.Range(wsHUD.Cells(1, 1), wsHUD.Cells(int_LastRow_HUD, int_LastCol_HUD))
    
    ' Dim Dictionary
        
        Dim dict_NewData As Scripting.Dictionary
            Set dict_NewData = New Scripting.Dictionary

    ' Dim "Ranges"
    
        ' QC Review
        Dim col_AcctNum_QCReview As Long
            col_AcctNum_QCReview = fx_Create_Headers("Account Number", arry_Header_QCReview)

        Dim col_NewCommDate_QCReview As Long
            col_NewCommDate_QCReview = fx_Create_Headers("New QC LC Comment Date", arry_Header_QCReview)

        ' New Report
        Dim col_AcctNum_NewReport As Long
            col_AcctNum_NewReport = fx_Create_Headers("Account Number", arry_Header_NewReport)

        Dim col_LCComment_NewReport As Long
            col_LCComment_NewReport = fx_Create_Headers("LC Comment Date", arry_Header_NewReport)
        
        ' HUD
        Dim col_AcctNum_HUD As Long
            col_AcctNum_HUD = fx_Create_Headers("Loan Number", arry_Header_HUD)

        Dim col_NewCommDate_HUD As Long
            col_NewCommDate_HUD = fx_Create_Headers("New QC LC Comment Date", arry_Header_HUD)

        Dim col_Mortgage_Status_Code As Long
            col_Mortgage_Status_Code = fx_Create_Headers("Mortgage Status Code", arry_Header_HUD)

    ' Dim Loop Variables
    
        Dim i As Long

' -----------
' Fill the Dictionary with the New Data, including Acct Num and Last Review Date ("LC Comment Date")
' -----------
    On Error Resume Next
        For i = 2 To int_LastRow_NewReport
            If arryNewData(i, col_LCComment_NewReport) <> "" Then
                dict_NewData.Add Key:=arryNewData(i, col_AcctNum_NewReport), Item:=arryNewData(i, col_LCComment_NewReport)
            End If
        Next i
    
' -----------
' Loop through the Dictionary to pull out the Last Review Date for QC Review
' -----------

        For i = 2 To int_LastRow_QC_Review
            If dict_NewData.Exists(arryQCData(i, col_AcctNum_QCReview)) Then
                arryQCData(i, col_NewCommDate_QCReview) = dict_NewData.Item(arryQCData(i, col_AcctNum_QCReview))
            End If
        Next i

    'Output the values from the array back to the sheet
    wsQCReview.Range(wsQCReview.Cells(1, 1), wsQCReview.Cells(int_LastRow_QC_Review, int_LastCol_QC_Review)) = arryQCData

' -----------
' Loop through the Dictionary to pull out the Last Review Date for HUD
' -----------

        For i = 2 To int_LastRow_HUD
            If dict_NewData.Exists(arryHUDData(i, col_AcctNum_HUD)) Then
                arryHUDData(i, col_NewCommDate_HUD) = dict_NewData.Item(arryHUDData(i, col_AcctNum_HUD))
            End If
        Next i

    'Output the values from the array back to the sheet
    wsHUD.Range(wsHUD.Cells(1, 1), wsHUD.Cells(int_LastRow_HUD, int_LastCol_HUD)) = arryHUDData

    On Error GoTo 0

' -----------
' Sort the data
' -----------
    
    ' Sort HUD

        wsHUD.Range("A1:" & strLastCol_HUD & int_LastRow_HUD).Sort Key1:=wsHUD.Cells(1, col_Mortgage_Status_Code), Order1:=xlAscending, Header:=xlYes
            
    ' QC Review
    
        With wsQCReview.Sort
            
            .SortFields.Clear
            .SortFields.Add Key:=wsQCReview.Cells(1, col_AgingCat_QCReview), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "120 DAYS LATE, 120 DAYS  AND OVER, 120 DAYS AND OVER, 90 DAYS LATE, 60 DAYS LATE, 30 DAYS LATE, AUTO NPL, MAN NPL, UNASSIGNED, NON ACCRUAL", DataOption:=xlSortNormal
            
            .SetRange Range("A1:" & strLastCol_QC_Review & int_LastRow_QC_Review)
            .Header = xlYes
            .Apply
        End With

End Sub

Sub o_31_Populate_Arrays()

    ary_Data = wsData.Range(wsData.Cells(1, 1), wsData.Cells(int_LastRow_wsData, int_LastCol_wsData))
    
    ary_AllCust = wsAllCust.Range(wsAllCust.Cells(1, 1), wsAllCust.Cells(int_LastRow_wsAllCust, int_LastCol_wsAllCust))

End Sub
Sub o_32_Import_Static_Fields()

' Purpose: To import the data that will remain static from the 6330 report.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 7/17/2020

' Change Log:
'          4/15/2020: Intial Creation
'          7/15/2020: Added in TIN
'          7/17/2020: Added in the code for the PPP Flag

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    Dim intRowCustData As Long
        intRowCustData = 2
    
    Dim intRowTempData As Long
    
    Dim strTempCustHelper As String

' -----------
' Copy the static data into the array
' -----------
    
    For intRowTempData = 2 To UBound(ary_AllCust, 1) 'Start with the temp data

LoopStart:
        If ary_AllCust(intRowTempData, col_Helper_AllCust) = ary_Data(intRowCustData, col_Helper_Data) Then 'If the current line matches based on Helper
            
            ary_AllCust(intRowTempData, 1) = intRowTempData - 1                                                     ' Ref #
            
            ary_AllCust(intRowTempData, col_AcctNum_AllCust) = ary_Data(intRowCustData, col_AcctNum_Data)     ' Acct #
            ary_AllCust(intRowTempData, col_LOB_AllCust) = ary_Data(intRowCustData, col_LOB_Data)             ' Line of Business
            ary_AllCust(intRowTempData, col_Region_AllCust) = ary_Data(intRowCustData, col_Region_Data)       ' Region
            ary_AllCust(intRowTempData, col_CustName_AllCust) = ary_Data(intRowCustData, col_CustName_Data)   ' Borrower Name
            ary_AllCust(intRowTempData, col_TIN_AllCust) = ary_Data(intRowCustData, col_TIN_Data)             ' TIN
            ary_AllCust(intRowTempData, col_TINHelper_AllCust) = _
            ary_Data(intRowCustData, col_LOB_Data) & " - " & ary_Data(intRowCustData, col_TIN_Data)           ' TIN Helper
            ary_AllCust(intRowTempData, col_PPPFlag_AllCust) = ary_Data(intRowCustData, col_PPPFlag_Data)     ' PPP Flag
                            
            ary_AllCust(intRowTempData, col_Industry_AllCust) = ary_Data(intRowCustData, col_Industry_Data)   ' Industry
            ary_AllCust(intRowTempData, col_PortMgr_AllCust) = ary_Data(intRowCustData, col_PortMgr_Data)     ' Portfolio Manager
            ary_AllCust(intRowTempData, col_SIC_AllCust) = ary_Data(intRowCustData, col_SIC_Data)             ' SIC Code
            
        Else
            intRowCustData = intRowCustData + 1
                If intRowCustData > int_LastRow_wsData Then Exit For
            GoTo LoopStart
        End If
   
    Next intRowTempData

End Sub
Sub o_33_Import_Dynamic_Fields()

' Purpose: To import the data that is dynamic and needs to be calculated, from the 6330 report.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/30/2020

' Change Log:
'          4/15/2020: Intial Creation
'          5/14/2020: Added new code for the CCRP to flag the 6Ws
'          6/30/2020: Added new code for the PPP loan amount

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    'Dim Integers
    
    Dim intRowCustData As Long:         intRowCustData = 2
    
    Dim intRowTempData As Long
    
    Dim intBRG As Long
    
    Dim intFRG As Long
    
    Dim dblCCRP As Double
    
    Dim intOutstanding As Double
    
    Dim intExposure As Double
    
    Dim intPPPBal As Double
    
    'Dim Strings

    Dim strHelper_AllCust As String
    
    Dim strLFT As String
    
    Dim strWatchInd As String
    
    Dim strNonAccrInd As String
    
    Dim strPPPFlag As String

    'Dim Dates
    
    Dim dtOpenDate As Date
        dtOpenDate = CDate("12/31/2099")

' -----------
' Copy the static data into the array
' -----------

    For intRowTempData = 2 To UBound(ary_AllCust, 1) 'Start with the temp data
    
LoopStart:
        If ary_AllCust(intRowTempData, col_Helper_AllCust) = ary_Data(intRowCustData, col_Helper_Data) Then 'If the current line matches based on Helper
                
            strHelper_AllCust = ary_AllCust(intRowTempData, col_Helper_AllCust)
                
            Do Until strHelper_AllCust <> ary_Data(intRowCustData, col_Helper_Data)
            
                If ary_Data(intRowCustData, col_BRG_Data) > intBRG Then intBRG = ary_Data(intRowCustData, col_BRG_Data)
                If ary_Data(intRowCustData, col_FRG_Data) > intFRG Then intFRG = ary_Data(intRowCustData, col_FRG_Data)
                If ary_Data(intRowCustData, col_CCRP_Data) > dblCCRP Then dblCCRP = ary_Data(intRowCustData, col_CCRP_Data)
                
                ' LFT
                If ary_Data(intRowCustData, col_LFT_Data) = 8 Then
                    strLFT = 8
                Else
                    If strLFT <> "8" Then strLFT = ary_Data(intRowCustData, col_LFT_Data)
                End If
            
                'Watch Flag
                If ary_Data(intRowCustData, col_WatchFlag_Data) = "Y" Then
                    strWatchInd = "Y"
                    If dblCCRP = 6 Then dblCCRP = 6.5 'Make a 6.5 to reflect the 6Ws
                Else
                    If strWatchInd <> "Y" Then strWatchInd = ary_Data(intRowCustData, col_WatchFlag_Data)
                End If
            
                'Non-Accrual Flag
                If ary_Data(intRowCustData, col_NonAccrual_Data) = "Y" Then
                    strNonAccrInd = "Y"
                Else
                    If strNonAccrInd <> "Y" Then strNonAccrInd = ary_Data(intRowCustData, col_NonAccrual_Data)
                End If
            
                'Date Open
                If CDate(ary_Data(intRowCustData, col_OpenDt_Data)) < dtOpenDate Then
                    dtOpenDate = CDate(ary_Data(intRowCustData, col_OpenDt_Data))
                End If
            
                'Outstanding / Exposure
                intOutstanding = intOutstanding + ary_Data(intRowCustData, col_Outstanding_Data)
                
                ' Exposure
                intExposure = intExposure + ary_Data(intRowCustData, col_Exposure_Data)
                
            intRowCustData = intRowCustData + 1
                If intRowCustData > int_LastRow_wsData Then Exit Do
                        
            Loop
            
        Else
            intRowCustData = intRowCustData + 1
            GoTo LoopStart
        End If

' ---------------
' Clear Variables
' ---------------
        
        ary_AllCust(intRowTempData, col_BRG_AllCust) = intBRG
            If intBRG = 0 Then ary_AllCust(intRowTempData, col_BRG_AllCust) = Empty
        ary_AllCust(intRowTempData, col_FRG_AllCust) = intFRG
            If intFRG = 0 Then ary_AllCust(intRowTempData, col_FRG_AllCust) = Empty
        ary_AllCust(intRowTempData, col_CCRP_AllCust) = dblCCRP
            If dblCCRP = 0 Then ary_AllCust(intRowTempData, col_CCRP_AllCust) = Empty
        ary_AllCust(intRowTempData, col_LFT_AllCust) = strLFT
        
        ary_AllCust(intRowTempData, col_WatchFlag_AllCust) = strWatchInd
        ary_AllCust(intRowTempData, col_NonAccrual_AllCust) = strNonAccrInd
        ary_AllCust(intRowTempData, col_OpenDt_AllCust) = dtOpenDate
        
        ary_AllCust(intRowTempData, col_Outstanding_AllCust) = intOutstanding
        ary_AllCust(intRowTempData, col_Exposure_AllCust) = intExposure
        
        intBRG = Empty
        intFRG = Empty
        dblCCRP = Empty
        strLFT = Empty
        strWatchInd = Empty
        strNonAccrInd = Empty
        strPPPFlag = Empty
        dtOpenDate = CDate("12/31/2099")
        intOutstanding = Empty
        intExposure = Empty
        intPPPBal = Empty

        If intRowCustData > int_LastRow_wsData Then Exit For

    Next intRowTempData

End Sub
Sub o_41_Copy_Arrays_to_AllCustws()
    
' Purpose: To copy the data from the arrays into the workbook.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 4/30/2020

' ***********************************************************************************************************************************
    
    wsAllCust.Range(wsAllCust.Cells(1, 1), wsAllCust.Cells(int_LastRow_wsAllCust, int_LastCol_wsAllCust)) = ary_AllCust

' -----------
' Output the control totals
' -----------

    With wsValidation
        int_CurRowValidation = .Cells(.Rows.Count, "A").End(xlUp).Row
    
        .Range("A" & int_CurRowValidation + 1) = Now
        .Range("B" & int_CurRowValidation + 1) = "o_1_Import_RiskTrend_Data > o_41_Copy_Arrays_to_AllCustws"
        .Range("C" & int_CurRowValidation + 1) = "wsAllCust"
        .Range("D" & int_CurRowValidation + 1) = Format(Application.WorksheetFunction.Sum(wsAllCust.Range("T2:T" & int_LastRow_wsAllCust)), "$#,##0")
        .Range("E" & int_CurRowValidation + 1) = Format(wsAllCust.Cells(Rows.Count, "C").End(xlUp).Row - 1, "0,0")

    End With

    Debug.Print "Total in wsAllCust:    " & Format(wsValidation.Range("D" & int_CurRowValidation + 1), "$#,##0")

End Sub
Sub o_42_Format_Rows()

' Purpose: To take the formatting from the first couple rows and apply to the whole sheet.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 4/30/2020

' ***********************************************************************************************************************************
  
' ----------------------
' Declare your variables
' ----------------------

    Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("ALL Customers")

With ws
        
    Dim LastRow As Long
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    Dim LastCol As Long
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column

    Dim rngFormat As Range
        Set rngFormat = .Range(.Cells(2, 1), .Cells(3, LastCol))
    
    Dim rngTarget As Range
        Set rngTarget = .Range(.Cells(4, 1), .Cells(LastRow, LastCol))
    
End With

' -----------
' Copy the formatting from the 2nd and 3rd row to the rest
' -----------
    
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteFormats
    
    Application.CutCopyMode = False

End Sub
Sub o_31_Validate_Control_Totals()

' Purpose: To validate that the control totals for the data imported match.
' Trigger: Called by uf_Sageworks_Regular
' Updated: 2/12/2021

' Change Log:
'       2/12/2021: Intial Creation

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler

' ----------------------
' Declare your variables
' ----------------------
    
    ' Dim Integers
    
    Dim int1stTotal As Double
        int1stTotal = wsValidation.Range("D2").Value
    
    Dim int2ndTotal As Double
        int2ndTotal = wsValidation.Range("D3").Value
    
    Dim int1stCount As Long
        int1stCount = wsValidation.Range("E2").Value
    
    Dim int2ndCount As Long
        int2ndCount = wsValidation.Range("E3").Value
        
    ' Dim Booleans
    
    Dim bolTotalsMatch As Boolean
        If int1stTotal = int2ndTotal And int1stCount = int2ndCount Then
            bolTotalsMatch = True
        Else
            bolTotalsMatch = False
        End If
    
' -----------
' Output the messagebox with the results
' -----------
   
    If bolTotalsMatch = True Then
    MsgBox Title:="Validation Totals Match", _
        Buttons:=vbOKOnly, _
        Prompt:="The validation totals match, you're golden. " & Chr(10) & Chr(10) _
        & "1st Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "1st Validation Count: " & Format(int1stCount, "0,0")
       
    ElseIf bolTotalsMatch = False Then
    MsgBox Title:="Validation Totals Don't Match", _
        Buttons:=vbCritical, _
        Prompt:="The validation totals from the Sageworks Dashboard Dump don't match what was imported. " _
        & "Please review the totals in the Validation worksheet to determine what went awry. " _
        & "Once the issues has been identifed talk to James to fix / reimport the data." & Chr(10) & Chr(10) _
        & "1st Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "1st Validation Total Variance: " & Format(int1stTotal - int2ndTotal, "$#,##0") & Chr(10) & Chr(10) _
        & "1st Validation Count: " & Format(int1stCount, "0,0") & Chr(10) _
        & "1st Validation Count Variance: " & Format(int1stCount - int2ndCount, "0,0")
    
    End If

Exit Sub
    
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
ErrorHandler:

myPrivateMacros.DisableForEfficiencyOff
    
End Sub
Sub o_14_Fix_End_Market()

' Purpose: To resolve the issues with the End Market coming from Sageworks to match to the CV Tracker data.
' Trigger: Called: o_01_MAIN_PROCEDURE_2_Import_Other_Data
' Updated: 3/5/2021

' Change Log:
'       3/5/2021: Intial Creation

' ***********************************************************************************************************************************

''' Sageworks                                       CV Tracker
'''
''' Healthcare - Pharma & Life Science              Healthcare - Pharma & Life Sciences
''' Healthcare - Medical Devices & Products         Healthcare - Medical Device & Products
''' Healthcare - Specailized Physician Services     Healthcare - Specialized Physician Services
''' Healthcare - Other Verticals                    Healthcare - All Other
''' Healthcare - Behavior Health                    Healthcare - Behavioral Health
''' Waste Management - Commercial C&D               Waste Management - Commercial, C&D

' ----------------------
' Declare your variables
' ----------------------
    
    ' Assign Integers
        int_LastRow_wsSageworks = wsSageworks.Cells(wsSageworks.Rows.Count, "A").End(xlUp).Row
            If int_LastRow_wsSageworks = 1 Then int_LastRow_wsSageworks = 2
    
    ' Dim Dictionaries
    Dim dict_EndMarkets As Scripting.Dictionary
        Set dict_EndMarkets = New Scripting.Dictionary
        dict_EndMarkets.CompareMode = TextCompare

    ' Dim Array
    Dim arry_EndMarkets() As Variant
        arry_EndMarkets = Application.Transpose(wsSageworks.Range("F1:F" & int_LastRow_wsSageworks))
        'arry_EndMarkets = wsSageworks.Range("F1:F" & int_LastRow_wsSageworks)

    ' Declare Loop Variables
    Dim i As Long
  
' -----------
' Fill the Dictionary with the Lookup Data
' -----------
    
    dict_EndMarkets.Add Key:="Healthcare - Pharma & Life Science", Item:="Healthcare - Pharma & Life Sciences"
    dict_EndMarkets.Add Key:="Healthcare - Medical Devices & Products", Item:="Healthcare - Medical Device & Products"
    dict_EndMarkets.Add Key:="Healthcare - Specailized Physician Services", Item:="Healthcare - Specialized Physician Services"
    dict_EndMarkets.Add Key:="Healthcare - Other Verticals", Item:="Healthcare - All Other"
    dict_EndMarkets.Add Key:="Healthcare - Behavior Health", Item:="Healthcare - Behavioral Health"
    dict_EndMarkets.Add Key:="Waste Management - Commercial C&D", Item:="Waste Management - Commercial, C&D"
    
' -----------
' Update the Sageworks data using the dictionary
' -----------

    For i = 2 To int_LastRow_wsSageworks
        If dict_EndMarkets.Exists(arry_EndMarkets(i)) Then
            arry_EndMarkets(i) = dict_EndMarkets.Item(arry_EndMarkets(i))
        End If
    Next i

    'Output the values from the array
    wsSageworks.Range("F1:F" & int_LastRow_wsSageworks).Value2 = Application.Transpose(arry_EndMarkets)

End Sub
Sub o_31_Import_BO5356_Data()

' Purpose: To import the data from the BO 5356 Report for the BB Prop Type tab.
' Trigger: Called: o_01_MAIN_PROCEDURE_2_Import_Other_Data
' Updated: 3/15/2021

' Change Log:
'       3/15/2021: Intial Creation

' ***********************************************************************************************************************************

'Sets the current directory and path
    
On Error Resume Next
    
    ChDrive ThisWorkbook.Path
        ChDir ThisWorkbook.Path

On Error GoTo 0

' ----------------------
' Declare your variables
' ----------------------

    ' Assign Workbooks

    Set wbSource = Functions.fx_Open_Workbook(strPromptTitle:="Select the current (3.1) BO 5356 Report")
        
    ' Assign Worksheets
    
    Set wsSource = wbSource.Worksheets(1)
    
    Dim wsBO5356 As Worksheet
    Set wsBO5356 = ThisWorkbook.Sheets("3.1 BB Prop Type TEST")
            
    ' Dim Integers
    
    Dim intFirstRow As Long
        intFirstRow = wsSource.Range("B:B").Find("Account Number").Row
    
    Dim int_LastRow As Long
        int_LastRow = wsSource.Range("A:A").Find("Grand Total").Row - 2
    
    ' Dim Strings (for function)
            
'    Dim str_ModuleName As String
'        str_ModuleName = "o_11_Import_Sageworks_Data"

' -----------
' Update the data from Dashboard Review
' -----------

    Call fx_Copy_in_Data_for_Matching_Fields( _
        wsSource:=wsSource, _
        wsDest:=wsBO5356, _
        intFirstRowtoImport:=intFirstRow, _
        int_LastRowtoImport:=int_LastRow, _
        bol_CloseSourceWb:=True)

End Sub
Sub Misc_Code()


    'Delete existing workook, and add a new one
    Application.DisplayAlerts = False
        If Evaluate("ISREF('Temp Arrays'!A1)") = True Then ThisWorkbook.Sheets("Temp Arrays").Delete
        ThisWorkbook.Worksheets.Add().Name = "Temp Arrays"
    Application.DisplayAlerts = True

    ' If the workbook is still open (aka you aborted) then close it
        If Not wbDailyReview Is Nothing Then wbDailyReview.Close savechanges:=True

        Select Case ws.Name
            Case "Current"
                ActiveWindow.DisplayHeadings = False
            Case "Projects"
                ActiveWindow.DisplayHeadings = False
            Case "Tasks"
                ActiveWindow.DisplayHeadings = False
            Case "Waiting"
                ActiveWindow.DisplayHeadings = False
            Case "Questions"
                ActiveWindow.DisplayHeadings = False
            Case "Recurring"
                ActiveWindow.DisplayHeadings = False
            Case "Someday-Maybe"
                ActiveWindow.DisplayHeadings = False
            Case "Temp"
                ActiveWindow.DisplayHeadings = True
            Case "Daily"
                ActiveWindow.DisplayHeadings = True
            Case "Lists"
                ActiveWindow.DisplayHeadings = True
        End Select

End Sub
Sub o_11_Add_a_New_Project()

' Purpose: To allow me to quickly input a new Project into my To Do.
' Trigger: Keyboard Shortcut - Ctrl + Shift + P
' Updated: 12/16/2020

' Change Log
'       3/31/2018: Initial Creation was sometime in Q1 2018
'       12/16/2020: Added the code to SetFocus on Project on initialization

' ***********************************************************************************************************************************
        
    uf_New_Project.Show vbModeless
        Call Macros.o_13_Filter_Projects

    ' Force the Task Descrption object to take Focus
    
    uf_New_Project.frm_Project.Enabled = False
        uf_New_Project.frm_Project.Enabled = True
        
    uf_New_Project.txt_Project.SetFocus

End Sub
Sub o_13_Filter_Projects()

' Purpose: To allow me to filter down to just the Active Projects for my To Do
' Trigger: Called: o_51_Reset_To_Do
' Updated: 11/9/2021

' Change Log:
'       9/30/2020: Updated to remove some redundant code
'       12/28/2020: Moved the sort code from the To Do Reset
'       6/15/2021: Changed the sort to make it match the Weekly Reset
'       7/27/2021: Removed 'D/A Strategy'
'       11/1/2021: Removed Assigning the Public Variables, and Disabling For Efficiency, it was redundant
'       11/9/2021: Updated to reflect new Personal Areas of Focus

' ***********************************************************************************************************************************
    
Call Macros.o_02_Assign_Private_Variables
Call Macros.o_03_Assign_Private_Variables_wsProjects
    
' --------------------------
' Filter and sort wsProjects
' --------------------------
    
    With wsProjects

    ' Filter down to only incomplete Projects
        If .AutoFilterMode = True Then .AutoFilter.ShowAllData
            int_LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        .Range("A1").Sort _
            Key1:=.Cells(1, col_Area), Order1:=xlAscending, _
            Key2:=.Cells(1, col_Project), Order2:=xlAscending, _
            Header:=xlYes
            
    ' Sort the data
        
        .Sort.SortFields.Clear

        #If Personal <> 1 Then
            .Sort.SortFields.Add Key:=Range("C:C"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
                "D/A Requests, Projects, Continuous, Personal", DataOption:=xlSortNormal
        #Else
            .Sort.SortFields.Add Key:=Range("C:C"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
                "Family, House / Yard, Personal, Financial, Continuous", DataOption:=xlSortNormal
        #End If

        .Sort.SetRange Rows("1:" & int_LastRow)
        .Sort.Header = xlYes
        .Sort.Apply
                                                   
    ' Apply the filter for Active only
        .Range("A1").AutoFilter Field:=col_Status, Criteria1:=Array("Active", "="), Operator:=xlFilterValues
                                                   
    ' Call the Apply Area Formatting Macro to apply the color formatting
        Call Macros.o_16_Apply_Projects_Formatting
            
    ' Hide the unused rows in the Next Actions tab
        .Rows((int_LastRow + 1) & ":" & (int_LastRow + 1)).Hidden = False
        .Rows((int_LastRow + 2) & ":" & (.Rows.Count)).Hidden = True
    
    ' Wipe out the selection
        Application.Goto .Range("A1"), False
        
    End With
              
End Sub
Function fx_Update_Single_Field_EXAMPLE(wsSource As Worksheet, wsDest As Worksheet, _
    str_Source_TargetField As String, str_Source_MatchField As String, str_Dest_TargetField As String, str_Dest_MatchField As String, _
    Optional int_SourceHeaderRow As Long, Optional bol_ConvertMatchSourcetoValues As Boolean, _
    Optional bol_CloseSourceWb As Boolean, Optional bol_SkipDuplicates As Boolean, Optional bol_BlanksOnly As Boolean, Optional str_OnlyUseValue As String, _
    Optional bol_MissingLookupData_MsgBox As Boolean, Optional bol_MissingLookupData_UseExistingData As Boolean, _
    Optional str_MissingLookupData_ValuetoUse, Optional str_WsNameLookup As String, _
    Optional str_FilterField_Dest As String, Optional str_FilterValue As String)

' Purpose: To update the data in the Target Field in the Destination, based on data from the Target Field in the Source.
    ' 1) Add the data to be referenced into the dictionary
    ' 2) Look for the data to be updated for matches to the reference dictionary
    ' 3) Replace the existing data with the reference data

' Trigger: Called
' Updated: 6/13/2022

' Change Log:
'       2/16/2021:  Initial creation, based on fx_Update_Data_SIC
'       2/17/2021:  Updated to convert over to pulling in the applicable ranges.
'       2/26/2021:  Tweaked the names of the paramaters
'       2/26/2021:  Rewrote to include as much of the code as possible in the function.
'       6/22/2021:  Added the bol_CloseSourceWb variable and related code.
'       6/30/2021:  Added the code to ignore duplicates, just output the value once
'       7/14/2021:  Added the option to convert the Match_Source to values (for Acct #s w/ leading 0s)
'       7/14/2021:  Added the option to pass the int_SourceHeaderRow and the related code
'       10/5/2021:  Updated to use the passed Target & Match fields to determine the int_LastRow
'       10/11/2021: Added the option for bol_BlanksOnly
'       10/12/2021: Added the option to ONLY update with a single value, if present (Ex. updating NPL flag for a borrower)
'       10/13/2021: Updated to convert the range from Text => General formatting, if bol_ConvertMatchSourcetoValues = True
'       4/18/2022:  Added the MsgBox for any missing data, and the bol to use it
'                   Updated the code to determine if there are missing fields to use a dictionary, and created a process to handle blanks
'       4/19/2022:  Added 'str_MissingLookupData_ValuetoUse' to allow a user to pass a value that will be used for blanks
'                   Updated the names of some of the variables to help clarify
'       6/13/2022:  Added the 'str_WsNameLookup' and related code when a value is missing.
'                   Added code to remove the leading line break in str_missingvalues

' ***********************************************************************************************************************************

' USE EXAMPLE: _
    Call fx_Update_Single_Field( _
        wsSource:=wsDetailDash, wsDest:=wsSageworks, _
        int_SourceHeaderRow:=4, _
        str_Source_TargetField:="14 Digit Acct#", _
        str_Source_MatchField:="Full Customer #", _
        str_Dest_TargetField:="Account Number", _
        str_Dest_MatchField:="Full Customer #", _
        bol_ConvertMatchSourcetoValues:=True, _
        str_WsNameLookup:="County Lookup", _
        bol_SkipDuplicates:=True, _
        bol_CloseSourceWb:=True, _
        bol_BlanksOnly:= True, _
        str_OnlyUseValue:= "Y", _
        bol_MissingLookupData_MsgBox:=True, _
        bol_MissingLookupData_UseExistingData:=True)

' LEGEND MANDATORY:
'   wsSource:
'   wsDest:
'   str_Source_TargetField:
'   str_Source_MatchField:
'   str_Dest_TargetField:
'   str_Dest_MatchField:

' LEGEND OPTIONAL:
'   int_SourceHeaderRow: The header row for the Source ws, if blank will default to 1
'   bol_ConvertMatchSourcetoValues: Converts to values only from the Source ws for the Match fields
'   bol_CloseSourceWb: Closes the Source wb when the code has finished
'   bol_SkipDuplicates: Removes duplicate values so the looked up data will only be used once
'   bol_BlanksOnly: Only updates the data if the field is currently blank
'   str_OnlyUseValue: Used to allow ONLY a single value to be used
'   bol_MissingLookupData_MsgBox: Outputs a message box with a list of fields that are missing from the lookup
'   bol_MissingLookupData_UseExistingData: Will use the existing data instead of the lookup value
'   str_MissingLookupData_ValuetoUse: If the value isn't in the lookup, and I didn't include a blank in the lookups, will use this value instead
'   str_FilterField_Dest: Used to filter down the values to be imported on
'   str_FilterValue: Used to filter down the values to be imported on

' ***********************************************************************************************************************************

' -------------------------------
' Declare your wsSource variables
' -------------------------------
        
With wsSource
        
    ' Dim wsSource Range References
    
    Dim intHeaderRow_wsSource As Long
    
        If int_SourceHeaderRow <> 0 Then
            intHeaderRow_wsSource = int_SourceHeaderRow
        Else
            intHeaderRow_wsSource = 1
        End If
    
    Dim int_LastCol_wsSource As Long
        int_LastCol_wsSource = .Cells(intHeaderRow_wsSource, Columns.Count).End(xlToLeft).Column
        
    ' Dim wsSource Column References
        
    Dim arry_Header_wsSource() As Variant
        arry_Header_wsSource = Application.Transpose(.Range(.Cells(intHeaderRow_wsSource, 1), .Cells(intHeaderRow_wsSource, int_LastCol_wsSource)))
        
    Dim col_Source_Target As Integer
        col_Source_Target = fx_Create_Headers(str_Source_TargetField, arry_Header_wsSource)

    Dim col_Source_Match As Integer
        col_Source_Match = fx_Create_Headers(str_Source_MatchField, arry_Header_wsSource)

    ' Dim wsSource Range References
        
    Dim int_LastRow_wsSource As Long
        int_LastRow_wsSource = WorksheetFunction.Max( _
        .Cells(Rows.Count, col_Source_Target).End(xlUp).Row, _
        .Cells(Rows.Count, col_Source_Match).End(xlUp).Row)
        
        If int_LastRow_wsSource = 1 Then int_LastRow_wsSource = 2
        
    ' Dim wsSource Ranges
        
    Dim rng_Source_Target As Range
    Set rng_Source_Target = .Range(.Cells(1, col_Source_Target), .Cells(int_LastRow_wsSource, col_Source_Target))
        
    Dim rng_Source_Match As Range
    Set rng_Source_Match = .Range(.Cells(1, col_Source_Match), .Cells(int_LastRow_wsSource, col_Source_Match))
        
    If bol_ConvertMatchSourcetoValues = True Then
        rng_Source_Match.NumberFormat = "General"
        rng_Source_Match.Value = rng_Source_Match.Value
        Set rng_Source_Match = .Range(.Cells(1, col_Source_Match), .Cells(int_LastRow_wsSource, col_Source_Match))
    End If
        
    ' Dim wsSource Arrays
    
    Dim arry_Source_Target() As Variant
        arry_Source_Target = Application.Transpose(rng_Source_Target)
    
    Dim arry_Source_Match() As Variant
        arry_Source_Match = Application.Transpose(rng_Source_Match)
        
End With
        
' -----------------------------
' Declare your wsDest variables
' -----------------------------
        
With wsDest
        
    ' Dim wsDest Range References
    
    Dim int_LastCol_wsDest As Long
        int_LastCol_wsDest = wsDest.Cells(1, Columns.Count).End(xlToLeft).Column
 
    ' Dim wsDest Column References
    
    Dim arry_Header_wsDest() As Variant
        arry_Header_wsDest = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsDest)))
        
    Dim col_Dest_Target As Integer
        col_Dest_Target = fx_Create_Headers(str_Dest_TargetField, arry_Header_wsDest)

    Dim col_Dest_Match As Integer
        col_Dest_Match = fx_Create_Headers(str_Dest_MatchField, arry_Header_wsDest)
        
    Dim col_Dest_FilterField As Integer
        col_Dest_FilterField = fx_Create_Headers(str_FilterField_Dest, arry_Header_wsDest)
        If col_Dest_FilterField = 0 Then col_Dest_FilterField = 999
 
    ' Dim wsDest Range References
 
    Dim int_LastRow_wsDest As Long
        int_LastRow_wsDest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.Count, col_Dest_Target).End(xlUp).Row, _
        wsDest.Cells(Rows.Count, col_Dest_Match).End(xlUp).Row)
        
        If int_LastRow_wsDest = 1 Then int_LastRow_wsDest = 2
 
    ' Dim Ranges
        
    Dim rng_Dest_Target As Range
    Set rng_Dest_Target = .Range(.Cells(1, col_Dest_Target), .Cells(int_LastRow_wsDest, col_Dest_Target))
        
    Dim rng_Dest_Match As Range
    Set rng_Dest_Match = .Range(.Cells(1, col_Dest_Match), .Cells(int_LastRow_wsDest, col_Dest_Match))
        
    Dim rng_Dest_FilterField As Range
    Set rng_Dest_FilterField = .Range(.Cells(1, col_Dest_FilterField), .Cells(int_LastRow_wsDest, col_Dest_FilterField))
        
    ' Dim Arrays
    
    Dim arry_Dest_Target() As Variant
        arry_Dest_Target = Application.Transpose(rng_Dest_Target)
    
    Dim arry_Dest_Match() As Variant
        arry_Dest_Match = Application.Transpose(rng_Dest_Match)

    Dim arry_Dest_Filter() As Variant
        arry_Dest_Filter = Application.Transpose(rng_Dest_FilterField)
 
End With
        
' ----------------------------
' Declare your Other variables
' ----------------------------
        
    ' Dim Dictionaries
    
    Dim dict_LookupData As Scripting.Dictionary
        Set dict_LookupData = New Scripting.Dictionary
            dict_LookupData.CompareMode = TextCompare
        
    Dim dict_MissingFields As Scripting.Dictionary
        Set dict_MissingFields = New Scripting.Dictionary
            dict_MissingFields.CompareMode = TextCompare
        
    ' Declare Loop Variables
    
    Dim i As Long
    
    Dim cntr_MissingFields As Integer
    
    Dim str_MissingValues As String
    
    Dim val As Variant
    
    ' Declare Message Variables
    
    Dim strMissingDataMessage As String
    
    If str_WsNameLookup <> "" Then
        strMissingDataMessage = str_WsNameLookup
    Else
        strMissingDataMessage = "(ex. 'Collateral Lookup')"
    End If

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------
    
On Error Resume Next
        
    For i = 1 To UBound(arry_Source_Target)
        If arry_Source_Target(i) <> "" And arry_Source_Match(i) <> "" Then
            If str_OnlyUseValue <> "" Then
                If arry_Source_Target(i) = str_OnlyUseValue Then ' Only import if the target matches the passed value
                    dict_LookupData.Add Key:=arry_Source_Match(i), Item:=arry_Source_Target(i)
                End If
            Else
                dict_LookupData.Add Key:=arry_Source_Match(i), Item:=arry_Source_Target(i)
            End If
        End If
    Next i

On Error GoTo 0
    
' -----------------------------------------------------------------------------
' Loop through the Lookup Data to pull in the updated data for the target field
' -----------------------------------------------------------------------------

    For i = 2 To UBound(arry_Dest_Match)
        
        If str_FilterField_Dest = "" Or arry_Dest_Filter(i) = str_FilterValue Then
        
            If dict_LookupData.Exists(arry_Dest_Match(i)) Then
                
                If bol_BlanksOnly = True Then
                    If arry_Dest_Target(i) = "" Or arry_Dest_Target(i) = 0 Then
                        arry_Dest_Target(i) = dict_LookupData.Item(arry_Dest_Match(i))
                    End If
                Else
                    arry_Dest_Target(i) = dict_LookupData.Item(arry_Dest_Match(i))
                End If
                
                If bol_SkipDuplicates = True Then dict_LookupData.Remove (arry_Dest_Match(i)) ' Remove so it can only be imported once
            ElseIf arry_Dest_Match(i) = Empty Then
            
                ' If I have a record for a blank in the lookups use that, or use the str_MissingLookupData_ValuetoUse if that was passed, otherwise abort
                If dict_LookupData.Exists(" ") = True Then
                    arry_Dest_Target(i) = dict_LookupData.Item(" ")
                
                ElseIf IsMissing(str_MissingLookupData_ValuetoUse) = False Then
                    arry_Dest_Target(i) = str_MissingLookupData_ValuetoUse
                    GoTo MissingDataMsgBox
                Else
                    GoTo MissingDataMsgBox
                End If
            
            Else

MissingDataMsgBox:

                If bol_MissingLookupData_MsgBox = True Then ' Let the user know that the data is missing
                    
                    ' Load the dictionary with each of the exceptions noted
                    On Error Resume Next
                        dict_MissingFields.Add Key:=arry_Dest_Match(i), Item:=arry_Dest_Match(i)
                        cntr_MissingFields = cntr_MissingFields + 1
                    On Error GoTo 0
                    
                End If
                
                If bol_MissingLookupData_UseExistingData = True Then ' Use the existing data to fill in the blank
                    arry_Dest_Target(i) = arry_Dest_Match(i)
                End If
                
                'wsDest.Cells(i, col_Dest_Target).Interior.Color = RGB(252, 213, 180) ' Highlight the missing data (disabled 6/13/2022)
            End If
        
        End If
        
    Next i

' -------------------------------------------------------
' Create the MsgBox if bol_MissingLookupData_MsgBox = True
' -------------------------------------------------------

        ' Create the list of fields
        For Each val In dict_MissingFields
            str_MissingValues = str_MissingValues & Chr(10) & "  > " & CStr(dict_MissingFields(val))
        Next val
        
        If Left(str_MissingValues, 1) = vbLf Then
            str_MissingValues = Right(str_MissingValues, Len(str_MissingValues) - 2)
        End If
        
        ' Output the Messagebox if there were any results
        If cntr_MissingFields > 0 Then
                
            MsgBox Title:="Missing Values in Lookup", _
                Buttons:=vbOKOnly + vbExclamation, _
                Prompt:="There is a missing record in the lookup table for:" & Chr(10) _
                & "'" & str_MissingValues & "'" & Chr(10) & Chr(10) _
                & "Please review the lookups in the applicable worksheet: " & Chr(10) _
                & strMissingDataMessage & Chr(10) & Chr(10) _
                & "Once the data has been reviewed re-run this process, or manually update the data."
        
        End If

    'Output the values from the array
    rng_Dest_Target.Value2 = Application.Transpose(arry_Dest_Target)

' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function
Sub o_4_Save_Workbook()

' Purpose: To save the workboko as a new version with today's date.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 12/19/2020

' Change Log:
'       12/19/2020: Intial Creation

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    'Dim Strings

        Dim strNewFileName As String
            strNewFileName = Format(Date, "YYYY-MM-DD") & " QC - R&R Quality Control Monthly Testing" & ".xlsx"
        
        Dim strNewFileFullPath As String
            strNewFileFullPath = ThisWorkbook.Path & "\" & strNewFileName

' -----------------
' Save the workbook
' -----------------

Application.DisplayAlerts = False
        
    ThisWorkbook.SaveAs Filename:=strNewFileFullPath, FileFormat:=xlOpenXMLWorkbook
        ThisWorkbook.Save

Application.DisplayAlerts = True

Exit Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
ErrorHandler:
    
    Call fx_Error_Handler(str_SubName:="o_4_Save_Workbook", bol_BreakCode:=True)
    Call PrivateMacros.DisableForEfficiencyOff

End Sub

